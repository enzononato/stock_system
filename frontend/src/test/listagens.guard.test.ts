import { describe, it, expect } from 'vitest'
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { resolve, join, relative, sep } from 'node:path'

const SRC = resolve(__dirname, '..')

/** Telas de outras fases que ainda têm tamanho de página fixo (caminhos relativos a `src/`). Sai da lista quando a fase dona migrar. */
const AGUARDANDO_OUTRAS_FASES: Record<string, string> = {}

const TAMANHO_FIXO = /\b[A-Z_]*PAGE_SIZE\s*=\s*\d+/

function listarArquivos(dir: string): string[] {
  return readdirSync(dir).flatMap((nome) => {
    const caminho = join(dir, nome)
    if (statSync(caminho).isDirectory()) return listarArquivos(caminho)
    return /\.(ts|tsx)$/.test(nome) && !/\.test\.(ts|tsx)$/.test(nome) ? [caminho] : []
  })
}

const arquivos = ['pages', 'components', 'features'].flatMap((d) => listarArquivos(join(SRC, d)))
const relativo = (arquivo: string) => relative(SRC, arquivo).split(sep).join('/')
const ler = (rel: string) => readFileSync(join(SRC, rel), 'utf8')

describe('listagens da fase 1', () => {
  it('nenhuma tela usa tamanho de página fixo', () => {
    const infratores = arquivos
      .map(relativo)
      .filter((rel) => !(rel in AGUARDANDO_OUTRAS_FASES))
      .filter((rel) => TAMANHO_FIXO.test(ler(rel)))
    expect(infratores).toEqual([])
  })

  it('a lista de espera não tem entradas vencidas', () => {
    for (const rel of Object.keys(AGUARDANDO_OUTRAS_FASES)) {
      expect(
        TAMANHO_FIXO.test(ler(rel)),
        `${rel} já não usa tamanho fixo: remova-o de AGUARDANDO_OUTRAS_FASES`
      ).toBe(true)
    }
  })

  it('telas da F1 não têm selo de status próprio', () => {
    const telas = [
      'pages/StockPage.tsx',
      'pages/PeripheralsPage.tsx',
      'pages/LinkPeripheralPage.tsx',
      'pages/HistoryPage.tsx',
      'pages/ReportPage.tsx',
      'pages/ChartsPage.tsx',
      'components/equipment/ItemDetailsModal.tsx',
    ]
    for (const rel of telas) {
      const codigo = ler(rel)
      expect(codigo, `${rel}: selo local`).not.toMatch(/function (MonoBadge|StatusBadge)\b/)
      expect(codigo, `${rel}: importa badge`).not.toMatch(/from '@\/components\/ui\/badge'/)
    }
  })

  it('listagens da F1 usam Pagination e usePageSize', () => {
    for (const rel of ['pages/StockPage.tsx', 'pages/PeripheralsPage.tsx', 'pages/HistoryPage.tsx', 'pages/ReportPage.tsx']) {
      const codigo = ler(rel)
      expect(codigo, `${rel}: Pagination`).toContain("from '@/components/ui/Pagination'")
      expect(codigo, `${rel}: usePageSize`).toContain('usePageSize(')
    }
  })
})
