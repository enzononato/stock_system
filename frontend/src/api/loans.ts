import api, { downloadAuthenticated } from "./client";

export async function initiateLoan(data: {
  item_id: number;
  usuario: string;
  cpf: string;
  center_cost: string;
  cargo: string;
  setor: string;
  revenda: string;
  date_issue: string;
  /**
   * Desmarcado por padrão: o termo assume "pessoa física" (o caso comum, já
   * que o campo ao lado no termo é CPF). Marque apenas quando o recebedor for
   * de fato representante de pessoa jurídica.
   */
  pessoa_juridica?: boolean;
}) {
  const res = await api.post("/loans", data);
  return res.data;
}

export async function confirmLoan(itemId: number, signedPdf: File) {
  const form = new FormData();
  form.append("signed_pdf", signedPdf);
  const res = await api.post(`/loans/${itemId}/confirm`, form);
  return res.data;
}

export async function initiateReturn(itemId: number) {
  const res = await api.post(`/loans/${itemId}/return/initiate`);
  return res.data as { detail: string; download_url: string; filename: string };
}

/**
 * Encapsula o fluxo de geração + download do termo de devolução: chama
 * `initiateReturn`, pega `download_url`/`filename` da resposta e baixa de forma
 * autenticada. Substitui o antigo `window.open(data.download_url, '_blank')`
 * (usado em ReturnPage.tsx), que sempre retornava 401 — `window.open` é uma
 * navegação crua do navegador e não envia o Authorization header exigido por
 * `/api/documents/files/{categoria}/{arquivo}`.
 *
 * @param filename Nome de exibição opcional; por padrão usa o `filename` devolvido por `initiateReturn`.
 */
export async function downloadReturnTerm(itemId: number, filename?: string): Promise<void> {
  const { download_url, filename: responseFilename } = await initiateReturn(itemId);
  await downloadAuthenticated(download_url, filename ?? responseFilename);
}

export async function confirmReturn(itemId: number, signedPdf: File) {
  const form = new FormData();
  form.append("signed_pdf", signedPdf);
  const res = await api.post(`/loans/${itemId}/return/confirm`, form);
  return res.data;
}

export async function downloadSignedTerm(itemId: number): Promise<Blob> {
  const urlRes = await api.get(`/loans/${itemId}/signed-term-url`);
  const fileUrl = (urlRes.data as { url: string }).url;
  // fileUrl comes as /api/documents/files/... — strip /api prefix to avoid double baseURL
  const relativePath = fileUrl.replace(/^\/api/, "");
  const fileRes = await api.get(relativePath, { responseType: "blob" });
  return fileRes.data as Blob;
}

export async function generateLoanTerm(itemId: number): Promise<Blob> {
  const res = await api.post(`/documents/loan-term/${itemId}`, {}, { responseType: "blob" });
  return res.data as Blob;
}
