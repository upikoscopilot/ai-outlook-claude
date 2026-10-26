/**
 * @file aipane.ts
 * @description The AI pane module for the Outlook add-in.
 * @author Ronan LE MEILLAT
 * @copyright 2024 Ronan LE MEILLAT for SCTG Development
 * @license AGPLv3
 */

import { config } from "./config";
import type { AIAnswer, AIModel, AIPrompt, AIProvider } from "./AIPrompt.js";
import DOMPurify from "dompurify";

const ERROR_MESSAGE: string = "Error: Unable to insert AI answer.";
const ANTHROPIC_VERSION: string = "2023-06-01";

/**
 * Sends a request to the Anthropic Messages API directly from the task pane.
 * No third-party proxy is involved: e-mail text and API key go only to api.anthropic.com.
 * @param {AIProvider} provider - The AI provider configuration.
 * @param {AIModel} model - The AI model configuration (max_tokens = maximum output tokens).
 * @param {string} apiKey - The Anthropic API key.
 * @param {string} systemText - The system prompt text.
 * @param {string} userText - The user input text.
 * @returns {Promise<string>} - The AI-generated response.
 */
async function aiRequest(
  provider: AIProvider,
  model: AIModel,
  apiKey: string,
  systemText: string,
  userText: string
): Promise<string> {
  const response = await fetch(`${provider.baseUrl}${provider.basePath}/messages`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": ANTHROPIC_VERSION,
      // Required for calls from a browser context such as an Office task pane
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: model.id,
      max_tokens: model.max_tokens,
      system: systemText,
      messages: [{ role: "user", content: userText }],
    }),
  });

  const data = await response.json().catch(() => null);
  if (!response.ok) {
    const message = data?.error?.message || response.statusText;
    throw new Error(`Anthropic API ${response.status}: ${message}`);
  }

  return (data?.content || [])
    .filter((block: { type: string }) => block.type === "text")
    .map((block: { text: string }) => block.text)
    .join("")
    .trim();
}

/**
 * Converts plain text into safe HTML for the mail body.
 * @param {string} text - Plain text.
 * @returns {string} - Escaped HTML with line breaks.
 */
function textToHtml(text: string): string {
  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
  return DOMPurify.sanitize(escaped);
}

/**
 * Replaces the current selection in the mail body with HTML.
 * @param {string} html - The HTML to insert.
 * @returns {Promise<void>}
 */
function setSelectedHtml(html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error.message));
        } else {
          resolve();
        }
      }
    );
  });
}

/**
 * Retrieves the prompt configuration by its ID.
 * @param {string} id - The ID of the prompt.
 * @returns {AIPrompt} - The prompt configuration.
 */
async function getPrompt(id: string): Promise<AIPrompt> {
  const prompts: AIPrompt[] = config.prompts;
  const isOutlook = await isOutlookClient();
  const prompt: AIPrompt | undefined = prompts.find(
    (prompt) => prompt.id === id && (!prompt.standalone || !isOutlook)
  );
  if (!prompt) {
    console.error("getPrompt: Prompt not found");
    throw new Error("Prompt not found");
  }
  return prompt;
}

/**
 * Retrieves the text selected in the email body.
 * @returns Promise<string> - The selected text.
 */
export function getSelectedText(): Promise<string> {
  return new Promise((resolve, reject) => {
    isOutlookClient()
      .then(() => {
        Office.context.mailbox.item?.getSelectedDataAsync(
          Office.CoercionType.Text,
          (asyncResult: Office.AsyncResult<any>) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(asyncResult.error);
            } else {
              const text = asyncResult.value.data;
              const prop = asyncResult.value.sourceProperty;
              console.log("Selected text in " + prop + ": " + text);
              resolve(text);
            }
          }
        );
      })
      .catch((error) => {
        reject(error);
      });
  });
}

/**
 * Inserts the AI-generated answer into the email body.
 * @param {AIProvider} provider - The AI provider configuration.
 * @param {AIModel} model - The AI model configuration.
 * @param {string} apiKey - The API key for authentication.
 * @param {string} id - The ID of the prompt.
 * @param {string} userText - The user input text.
 * @returns {Promise<AIAnswer>} - The AI-generated response and any errors.
 */
export async function insertAIAnswer(
  provider: AIProvider,
  model: AIModel,
  apiKey: string,
  id: string,
  userText: string
): Promise<AIAnswer> {
  const { system, user } = await getPrompt(id);

  // Validate and sanitize inputs
  if (!system || !user || !userText) {
    console.error("insertAIAnswer: Invalid input");
    throw new Error("insertAIAnswer: Invalid input");
  }

  try {
    console.log(`Prompt: ${id}`);

    // Make the AI request and get the response
    const aiText: string = await aiRequest(provider, model, apiKey, system, `${user}\n${userText}`);
    console.log(`AI provider: ${provider.name} AI model: ${model.name}`);

    // Insert the AI-generated text into the email body
    if (await isOutlookClient()) {
      await setSelectedHtml(textToHtml(aiText));
    }
    return { response: aiText, error: null };
  } catch (err) {
    console.error("insertAIAnswer: " + err);
    return { response: "", error: `${ERROR_MESSAGE} ${err instanceof Error ? err.message : err}` };
  }
}

let officeLoaded = false;
if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    console.log(`Office SDK ready running ${info.host ? "in " + info.host : "outside office"}`);
    officeLoaded = true;
  });
}
/**
 * Wait for the Office.js library to load.
 * limit the waiting time to 2 seconds
 * @returns {Promise<boolean>} - A promise that resolves to true when the Office.js library is loaded or to false after 2 seconds.
 */
export async function waitForOffice(): Promise<boolean> {
  return new Promise((resolve) => {
    if (officeLoaded) {
      resolve(true);
    } else {
      setTimeout(() => {
        resolve(false);
      }, 2000);
    }
  });
}

/**
 * Detects if the current app runs in the Outlook client.
 * @returns {boolean} - True if the app runs in the Outlook client, false otherwise.
 */
export async function isOutlookClient(): Promise<boolean> {
  const isOutlookLoaded = await waitForOffice();
  return typeof Office.context.mailbox !== "undefined" && isOutlookLoaded;
}
