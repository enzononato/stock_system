export interface MovimentacoesIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function MovimentacoesIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Movimentações",
}: MovimentacoesIconProps) {
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
      className={`animated-sidebar-icon animated-movimentacoesicon ${className}`}
    >
      <title>{title}</title>
      <path d="M8 3 4 7l4 4"/><path d="M4 7h16"/><path d="m16 21 4-4-4-4"/><path d="M20 17H4"/>
    </svg>
  );
}
