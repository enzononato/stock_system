export interface EmprestimosIconProps {
  size?: number | string;
  className?: string;
  strokeWidth?: number;
  title?: string;
}

export function EmprestimosIcon({
  size = 20,
  className = "",
  strokeWidth = 1.8,
  title = "Empréstimos",
}: EmprestimosIconProps) {
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
      className={`animated-sidebar-icon animated-emprestimosicon ${className}`}
    >
      <title>{title}</title>
      <path d="M18 11V6a2 2 0 0 0-4 0v5"/><path d="M14 10V4a2 2 0 0 0-4 0v6"/><path d="M10 10V6a2 2 0 0 0-4 0v8"/><path d="M6 14v-1a2 2 0 0 0-4 0c0 6 4 10 10 10h1a7 7 0 0 0 7-7v-5a2 2 0 0 0-4 0v2"/>
    </svg>
  );
}
