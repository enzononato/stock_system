import api from "./client";

/**
 * Unidade (antiga "revenda" fixa no código): agora é uma entidade no banco,
 * com os dados jurídicos usados em termos/documentos (razão social, CNPJ,
 * endereço). `endereco`, `cep`, `cidade` e `uf` são opcionais porque nem toda
 * unidade cadastrada tem esses dados completos ainda (ex.: Petrolina, sem CEP
 * porque não constava no documento de origem que originou o cadastro).
 */
export interface Unidade {
  id: number;
  nome: string;
  razao_social: string;
  cnpj: string;
  endereco?: string;
  cep?: string;
  cidade?: string;
  uf?: string;
  is_active: boolean;
}

/** Payload de criação/edição — mesmos campos estruturados de `Unidade`, sem `id`/`is_active`. */
export interface UnidadeInput {
  nome: string;
  razao_social: string;
  cnpj: string;
  endereco?: string;
  cep?: string;
  cidade?: string;
  uf?: string;
}

/**
 * Indicadores por unidade (ver GET /api/unidades/{id}/indicadores). `itens_por_status`
 * sempre traz as 4 chaves de status possíveis (Disponível, Indisponível, Pendente,
 * Pendente Devolução) mesmo quando a contagem é zero — nenhuma delas some por falta
 * de itens naquele status.
 */
export interface IndicadoresUnidade {
  termos_emitidos: number;
  termos_confirmados: number;
  devolucoes_concluidas: number;
  itens_total: number;
  itens_por_status: Record<string, number>;
  emprestimos_ativos: number;
  perifericos_vinculados: number;
}

export interface ListUnidadesParams {
  include_inactive?: boolean;
}

/** `GET /api/unidades` — qualquer usuário autenticado. Sem `include_inactive`, só traz as ativas. */
export async function listUnidades(params?: ListUnidadesParams): Promise<Unidade[]> {
  const res = await api.get("/unidades", { params });
  return res.data as Unidade[];
}

/** `GET /api/unidades/{id}` — qualquer usuário autenticado. */
export async function getUnidade(id: number): Promise<Unidade> {
  const res = await api.get(`/unidades/${id}`);
  return res.data as Unidade;
}

/** `GET /api/unidades/{id}/indicadores` — restrito a Gestor e Técnico no backend. */
export async function getIndicadoresUnidade(id: number): Promise<IndicadoresUnidade> {
  const res = await api.get(`/unidades/${id}/indicadores`);
  return res.data as IndicadoresUnidade;
}

/** `POST /api/unidades` — restrito a Gestor. */
export async function createUnidade(data: UnidadeInput): Promise<{ detail: string }> {
  const res = await api.post("/unidades", data);
  return res.data;
}

/** `PUT /api/unidades/{id}` — restrito a Gestor. Renomear cascateia no backend (ver aviso na UnidadesPage). */
export async function updateUnidade(id: number, data: UnidadeInput): Promise<{ detail: string }> {
  const res = await api.put(`/unidades/${id}`, data);
  return res.data;
}

/**
 * `DELETE /api/unidades/{id}` — restrito a Gestor. Inativa (soft delete); o
 * backend recusa com 400 se houver itens ativos vinculados à unidade, e o
 * `detail` do erro explica o motivo (deve ser exibido tal qual ao usuário).
 */
export async function deactivateUnidade(id: number): Promise<{ detail: string }> {
  const res = await api.delete(`/unidades/${id}`);
  return res.data;
}

/**
 * `POST /api/unidades/{id}/reativar` — restrito a Gestor. Reverte a inativação
 * feita por `deactivateUnidade`; sem este endpoint não havia caminho de volta
 * para uma unidade inativada por engano (gap na especificação original,
 * corrigido nesta revisão do contrato).
 */
export async function reactivateUnidade(id: number): Promise<{ detail: string }> {
  const res = await api.post(`/unidades/${id}/reativar`);
  return res.data;
}
