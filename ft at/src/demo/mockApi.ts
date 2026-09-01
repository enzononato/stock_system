/** TEMPORARY FRONTEND DEMO API. Remove before real API integration. */

export type DemoUser = {
  id: number;
  name: string;
  email: string;
  role: "Gestor" | "Técnico" | "Usuário";
};

// Intentionally enabled during the frontend-only build. No environment variable is required.
export const DEMO_API_ENABLED = true;

export const DEMO_CREDENTIALS = {
  email: "demo@teste.com",
  password: "123456",
};

const DEMO_USER: DemoUser = {
  id: 1,
  name: "Usuário Demo",
  email: DEMO_CREDENTIALS.email,
  role: "Gestor",
};

const DEMO_TOKEN = "demo-frontend-only-token";

export function isDemoApiEnabled() { return DEMO_API_ENABLED; }

export function demoLogin(email: string, password: string) {
  if (!email.trim() || !password.trim()) throw new Error("Informe e-mail e senha.");
  return { access_token: DEMO_TOKEN, user: DEMO_USER };
}

export function demoMe() { return DEMO_USER; }
export function demoRefresh() { return { access_token: DEMO_TOKEN }; }
