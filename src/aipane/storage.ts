/**
 * @file storage.ts
 * @description Safe key/value storage for the task pane.
 * Outlook: Office roaming settings (persist per mailbox, also in the new Outlook and on the web).
 * Fallback: localStorage, which can throw in a cross-site iframe when the browser blocks
 * third-party storage; last resort: in-memory for the current session.
 */

const memory: Record<string, string> = {};

function roaming(): Office.RoamingSettings | null {
  try {
    return typeof Office !== "undefined" && Office.context?.roamingSettings ? Office.context.roamingSettings : null;
  } catch {
    return null;
  }
}

export function getItem(key: string): string | null {
  const rs = roaming();
  const fromRoaming = rs ? (rs.get(key) as string | undefined) : undefined;
  if (fromRoaming) return fromRoaming;
  try {
    const fromLocal = window.localStorage.getItem(key);
    if (fromLocal !== null) return fromLocal;
  } catch {
    /* storage blocked */
  }
  return key in memory ? memory[key] : null;
}

export function setItem(key: string, value: string): void {
  memory[key] = value;
  try {
    window.localStorage.setItem(key, value);
  } catch {
    /* storage blocked */
  }
  const rs = roaming();
  if (rs) {
    rs.set(key, value);
    rs.saveAsync();
  }
}

export function removeItem(key: string): void {
  delete memory[key];
  try {
    window.localStorage.removeItem(key);
  } catch {
    /* storage blocked */
  }
  const rs = roaming();
  if (rs) {
    rs.remove(key);
    rs.saveAsync();
  }
}
