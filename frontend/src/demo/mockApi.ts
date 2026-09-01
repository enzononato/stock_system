/** TEMPORARY FRONTEND DEMO API. Remove before real API integration. */

export type DemoUser = {
  id: number;
  name: string;
  email: string;
  role: "Gestor" | "Técnico" | "Usuário";
};

export const DEMO_API_ENABLED =
  typeof import.meta !== "undefined" && Boolean(import.meta.env?.["VITE_ENABLE_DEMO_API"] === "true");

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
