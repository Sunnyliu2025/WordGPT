import * as React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import Center from "./Center";

interface LoginProps {
  onSave: (token: string) => void;
}

export default function Login({ onSave }: LoginProps) {
  const [token, setToken] = React.useState<string>("");

  return (
    <div style={{ marginTop: "24px" }}>
      <TextField
        style={{
          width: "100%",
        }}
        value={token}
        onChange={(_, newValue: string) => setToken(newValue || "")}
        placeholder={"Insert your API key here"}
      />
      <Center
        style={{
          marginTop: "16px",
        }}
      >
        <DefaultButton
          iconProps={{
            iconName: "SaveAs",
          }}
          onClick={() => onSave(token)}
          styles={{
            root: {
              minWidth: "140px",
              height: "36px",
              borderRadius: "4px",
            },
          }}
        >
          Save API key
        </DefaultButton>
      </Center>
    </div>
  );
}
