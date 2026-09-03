import { isAxiosError } from "axios";

/**
 * Traduz um erro de rede/HTTP em uma mensagem apresentável ao usuário.
 * Prioriza o `detail` devolvido pelo backend (ex.: regras de negócio de
 * unidades/remoção), sem nunca expor stack traces ou objetos crus.
 */
export function getErrorMessage(
  error: unknown,
  fallback = "Não foi possível concluir a operação.",
): string {
  if (isAxiosError(error)) {
    if (error.code === "ERR_NETWORK")
      return "Não foi possível falar com o servidor. Verifique sua conexão.";
    const data = error.response?.data as unknown;
    if (typeof data === "string" && data.trim() && !data.trim().startsWith("<")) return data;
    if (data && typeof data === "object") {
      const detail =
        (data as { detail?: unknown; message?: unknown }).detail ??
        (data as { message?: unknown }).message;
      if (typeof detail === "string" && detail.trim()) return detail;
      if (Array.isArray(detail)) {
        const first = detail[0] as { msg?: string } | undefined;
        if (first?.msg) return first.msg;
      }
    }
    const status = error.response?.status;
    if (status === 401) return "Sessão inválida ou expirada.";
    if (status === 403) return "Você não tem permissão para esta ação.";
    if (status === 404) return "Registro não encontrado.";
  }
  if (error instanceof Error && error.message) return error.message;
  return fallback;
}
