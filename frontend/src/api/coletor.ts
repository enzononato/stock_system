import api from './client'

/** Espelho de `POST /api/coletor`: `url_arquivo` é um caminho relativo, público, que o navegador baixa direto. */
export interface ColetaCriada {
  id: number
  url_coletor_local?: string
  comando: string
  url_arquivo: string
  expira_em: string
}

/** Equipamento já cadastrado que bate com os dados coletados (`ativo=false` quando está baixado). */
export interface Correspondencia {
  id: number
  tipo: string
  brand: string | null
  model: string | null
  status: string | null
  revenda: string | null
  ativo: boolean
  /** Códigos do motivo do casamento: `serial`, `mac`, `anydesk`. */
  motivos: string[]
}

export type StatusColeta = 'aguardando' | 'recebido' | 'vencido'

export interface AndamentoColeta {
  status: StatusColeta
  expira_em: string
  campos?: Record<string, string>
  tipo_sugerido?: string | null
  correspondencias?: Correspondencia[]
}

export async function criarColeta(): Promise<ColetaCriada> {
  const res = await api.post<ColetaCriada>('/coletor')
  return res.data
}

export async function andamentoColeta(id: number): Promise<AndamentoColeta> {
  const res = await api.get<AndamentoColeta>(`/coletor/${id}/andamento`)
  return res.data
}
