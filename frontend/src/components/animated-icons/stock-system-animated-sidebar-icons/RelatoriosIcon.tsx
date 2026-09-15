import * as React from "react";

export interface RelatoriosIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function RelatoriosIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Relatórios",
}: RelatoriosIconProps) {
  return (
    <svg
      aria-label={title}
      role="img"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      className={`animated-sidebar-icon animated-relatoriosicon ${className}`}
    >
      <title>{title}</title>
      <path d="M3 3v18h18"/><path d="M7 16v-5"/><path d="M12 16V7"/><path d="M17 16V4"/>
    </svg>
  );
}
