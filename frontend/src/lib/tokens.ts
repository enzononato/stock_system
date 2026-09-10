/**
 * Tokens semânticos monocromáticos do Design System Enterprise Editorial v5.
 * Espelho TypeScript para consumo em Recharts, SVG e elementos dinâmicos.
 */

export const TOKENS = {
  light: {
    canvas: "#F7F7F5",
    surface: "#FFFFFF",
    surfaceAlt: "#F0F0ED",
    border: "#D9D9D4",
    borderStrong: "#BDBDB7",
    textPrimary: "#111111",
    textSecondary: "#5F5F5A",
    textMuted: "#85857E",
    black: "#000000",
    white: "#FFFFFF",
    // Gráficos monocromáticos
    chartPrimary: "#111111",
    chartSecondary: "#5F5F5A",
    chartTertiary: "#85857E",
    chartGrid: "#D9D9D4",
  },
  dark: {
    canvas: "#050505",
    surface: "#0C0C0C",
    surfaceAlt: "#131313",
    border: "#272727",
    borderStrong: "#3A3A3A",
    textPrimary: "#F2F2F0",
    textSecondary: "#A4A4A0",
    textMuted: "#70706C",
    black: "#000000",
    white: "#FFFFFF",
    // Gráficos monocromáticos
    chartPrimary: "#F2F2F0",
    chartSecondary: "#A4A4A0",
    chartTertiary: "#70706C",
    chartGrid: "#272727",
  },
  fonts: {
    sans: '"Plus Jakarta Sans", ui-sans-serif, system-ui, sans-serif',
    mono: '"IBM Plex Mono", ui-monospace, SFMono-Regular, monospace',
  },
  motion: {
    easing: "cubic-bezier(0.16, 1, 0.3, 1)",
    durationMicro: "150ms",
    durationOverlay: "250ms",
  },
  radius: {
    sm: "2px",
    DEFAULT: "4px",
    md: "6px",
    max: "8px",
  },
} as const;

export type ThemeTokens = typeof TOKENS.light;
