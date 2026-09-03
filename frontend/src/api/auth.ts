import api from "./client";
import { demoLogin, demoMe, demoRefresh, isDemoApiEnabled } from "@/demo/mockApi";

export async function login(username: string, password: string) {
  if (isDemoApiEnabled()) {
    const result = demoLogin(username, password);
    if (!result) throw new Error("Modo demonstração indisponível.");
    return { access_token: result.access_token };
  }

  const res = await api.post("/auth/login", { username, password });
  return res.data as { access_token: string };
}

export async function logout() {
  if (isDemoApiEnabled()) return;
  await api.post("/auth/logout");
}

export async function getMe() {
  if (isDemoApiEnabled()) {
    const user = demoMe();
    if (!user) throw new Error("Modo demonstração indisponível.");
    return { id: user.id, username: user.email, role: user.role };
  }

  const res = await api.get("/auth/me");
  return res.data as { id: number; username: string; role: string };
}

export async function refresh() {
  if (isDemoApiEnabled()) {
    const result = demoRefresh();
    if (!result) throw new Error("Modo demonstração indisponível.");
    return result;
  }

  const res = await api.post("/auth/refresh");
  return res.data as { access_token: string };
}
