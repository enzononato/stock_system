export interface PerifericosIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function PerifericosIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Periféricos",
}: PerifericosIconProps) {
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
      className={`animated-sidebar-icon animated-perifericosicon ${className}`}
    >
      <title>{title}</title>
      <rect width="14" height="18" x="5" y="3" rx="7"/><path d="M12 3v6"/><path d="M9 9h6"/>
    </svg>
  );
}
