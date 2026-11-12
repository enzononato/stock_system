import api, { type Paginated } from './client'

/** Registro da trilha de auditoria. `antes`/`depois` trazem só as chaves que mudaram (CPF já mascarado). */
export interface AuditEntry {
  id: number
  /** UTC em ISO, sem offset (ex.: `2026-11-10T15:30:00`). */
  ocorrido_em: string
  username?: string | null
  papel?: string | null
  ip?: string | null
  entidade: string
  entidade_id?: string | null
  acao: string
  antes?: Record<string, unknown> | null
  depois?: Record<string, unknown> | null
}

export interface ListAuditParams {
  entidade?: string
  entidade_id?: string
  username?: string
  /** `AAAA-MM-DD`. */
  de?: string
  ate?: string
  limit?: number
  offset?: number
}

/** `GET /api/audit` — restrito a Gestor; filtros e paginação aplicados no servidor. */
export async function listAudit(params: ListAuditParams = {}): Promise<Paginated<AuditEntry>> {
  const res = await api.get('/audit', { params })
  return res.data as Paginated<AuditEntry>
}
