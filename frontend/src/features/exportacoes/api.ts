import { isAxiosError } from 'axios'
import { downloadAuthenticated } from '@/api/client'
import type { ListManutencoesParams } from '@/api/manutencoes'
import type { EstoqueSort, SortDir } from '@/features/estoque/api'

export type FormatoExportacao = 'xlsx' | 'pdf'

type Parametros = Record<string, string | number | null | undefined>

/** Query string sem os valores vazios (`undefined`, `null`, `''`); o zero é mantido. */
export function montarQuery(params: Parametros): string {
  const query = new URLSearchParams()
  for (const [chave, valor] of Object.entries(params)) {
    if (valor === undefined || valor === null || valor === '') continue
    query.set(chave, String(valor))
  }
  return query.toString()
}

async function lerTexto(blob: Blob): Promise<string> {
  if (typeof blob.text === 'function') return blob.text()
  return new Promise((resolve, reject) => {
    const leitor = new FileReader()
    leitor.onload = () => resolve(String(leitor.result ?? ''))
    leitor.onerror = () => reject(leitor.error)
    leitor.readAsText(blob)
  })
}

/**
 * O download pede `responseType: 'blob'`, então o corpo de um erro também chega como Blob e
 * `getErrorMessage` não enxerga o `detail`. Aqui o Blob é lido como JSON e regravado em
 * `error.response.data`; corpo que não é JSON é descartado (cai na mensagem padrão do status).
 */
async function traduzirErroDeBlob(erro: unknown): Promise<void> {
  if (!isAxiosError(erro) || !erro.response) return
  const dados: unknown = erro.response.data
  if (!(dados instanceof Blob)) return
  try {
    erro.response.data = JSON.parse(await lerTexto(dados))
  } catch {
    erro.response.data = undefined
  }
}

export async function exportar(
  caminho: string,
  params: Parametros,
  nomeBase: string,
  formato: FormatoExportacao
): Promise<void> {
  try {
    await downloadAuthenticated(
      `/exportacoes/${caminho}?${montarQuery({ ...params, formato })}`,
      `${nomeBase}.${formato}`
    )
  } catch (erro) {
    await traduzirErroDeBlob(erro)
    throw erro
  }
}

export function exportarEstoque(
  p: {
    tipo?: string
    status?: string
    revenda?: string
    search?: string
    colaborador_id?: number
    sort: EstoqueSort
    dir: SortDir
  },
  formato: FormatoExportacao
) {
  return exportar('estoque', p, 'estoque', formato)
}

export function exportarHistorico(p: { search?: string }, formato: FormatoExportacao) {
  return exportar('historico', p, 'historico', formato)
}

export function exportarManutencoes(
  p: Omit<ListManutencoesParams, 'limit' | 'offset'>,
  formato: FormatoExportacao
) {
  return exportar('manutencoes', p, 'manutencoes', formato)
}

export function exportarLinhas(
  p: { busca?: string; status?: string; unidade_id?: number; operadora?: string },
  formato: FormatoExportacao
) {
  return exportar('linhas', p, 'linhas', formato)
}

export function exportarTermosPendentes(p: { busca?: string }, formato: FormatoExportacao) {
  return exportar('termos-pendentes', p, 'termos_pendentes', formato)
}

export function exportarRelatorioMensal(year: number, month: number, formato: FormatoExportacao) {
  return exportar(
    'relatorio-mensal',
    { year, month },
    `relatorio_mensal_${year}_${String(month).padStart(2, '0')}`,
    formato
  )
}
