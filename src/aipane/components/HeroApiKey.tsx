/*
=========================================================
* © 2024 Ronan LE MEILLAT for SCTG Development
=========================================================
*/
import * as React from "react";
import { makeStyles, useId, Button, Input, Label } from "@fluentui/react-components";
import { AIProvider } from "../AIPrompt";

interface HeroApiKeyProps {
  provider: AIProvider;
  apiKey: string | null;
  onApiKeyChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  onApiKeySubmit: () => void;
  error?: string | null;
}

const useStyles = makeStyles({
  root: {
    // Stack the label above the field
    display: "flex",
    flexDirection: "column",
    // Use 2px gap below the label (per the design system)
    gap: "2px",
    margin: "10px",
  },
  input: {
    marginTop: "20px",
    marginBottom: "20px",
  },
});

const HeroApiKey: React.FC<HeroApiKeyProps> = ({ apiKey, onApiKeyChange, onApiKeySubmit, provider, error }) => {
  const styles = useStyles();
  const inputId = useId("input");

  return (
    <div className={styles.root}>
      <Label htmlFor={inputId} size="large">
        {provider.name} API-Key
      </Label>
      <span>
        Den Key erstellst du unter{" "}
        <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer">
          console.anthropic.com
        </a>
        . Er wird nur hier in Outlook gespeichert.
      </span>
      <Input
        className={styles.input}
        id={inputId}
        type="password"
        placeholder="sk-ant-..."
        value={apiKey || ""}
        onChange={onApiKeyChange}
        onKeyDown={(event) => {
          if (event.key === "Enter") onApiKeySubmit();
        }}
      />
      {error ? <span style={{ color: "#c50f1f", marginBottom: "10px" }}>{error}</span> : null}
      <Button appearance="primary" disabled={false} size="large" onClick={onApiKeySubmit}>
        Speichern
      </Button>
    </div>
  );
};

export default HeroApiKey;
