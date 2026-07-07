import * as React from "react";

interface ContainerProps {
  children: React.ReactNode | React.ReactNode[];
}

export default function Container({ children }: ContainerProps) {
  return (
    <div
      style={{
        padding: "20px",
        minHeight: "100%",
        boxSizing: "border-box",
        background: "#fff",
        borderRadius: "0",
        boxShadow: "none",
      }}
    >
      {children}
    </div>
  );
}
