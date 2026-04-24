import * as React from "react";

interface ContainerProps {
  children: React.ReactNode | React.ReactNode[];
}

export default function Container({ children }: ContainerProps) {
  return (
    <div
      style={{
        padding: "16px",
        minHeight: "100%",
        boxSizing: "border-box",
      }}
    >
      {children}
    </div>
  );
}
