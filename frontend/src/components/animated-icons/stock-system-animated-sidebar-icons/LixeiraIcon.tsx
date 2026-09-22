export interface LixeiraIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function LixeiraIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Lixeira",
}: LixeiraIconProps) {
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
      className={`animated-sidebar-icon animated-lixeiraicon ${className}`}
    >
      <title>{title}</title>
      <path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M10 11v6"/><path d="M14 11v6"/>
    </svg>
  );
}
