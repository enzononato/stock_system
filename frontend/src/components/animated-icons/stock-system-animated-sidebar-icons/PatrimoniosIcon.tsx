export interface PatrimoniosIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function PatrimoniosIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Patrimônios",
}: PatrimoniosIconProps) {
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
      className={`animated-sidebar-icon animated-patrimoniosicon ${className}`}
    >
      <title>{title}</title>
      <rect width="20" height="14" x="2" y="3" rx="2"/><line x1="8" x2="16" y1="21" y2="21"/><line x1="12" x2="12" y1="17" y2="21"/>
    </svg>
  );
}
