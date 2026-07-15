import * as React from "react";

interface CenterProps {
  children: React.ReactNode;
}

export default function Center({ children }: CenterProps) {
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        width: "100%",
      }}
    >
      {children}
    </div>
  );
}
