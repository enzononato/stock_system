/** Layout da folha de etiquetas (A4), todo em milímetros. */
export interface LayoutEtiquetas {
  colunas: number
  linhas: number
  larguraMm: number
  alturaMm: number
  margemSuperiorMm: number
  margemEsquerdaMm: number
  espacoHorizontalMm: number
  espacoVerticalMm: number
}

export const FOLHA_A4 = { larguraMm: 210, alturaMm: 297 } as const

/** 3 colunas × 8 linhas de 70 × 37 mm: 24 por folha (a margem de 0,5 mm absorve a tolerância da impressora). */
export const LAYOUT_PADRAO: LayoutEtiquetas = {
  colunas: 3,
  linhas: 8,
  larguraMm: 70,
  alturaMm: 37,
  margemSuperiorMm: 0.5,
  margemEsquerdaMm: 0,
  espacoHorizontalMm: 0,
  espacoVerticalMm: 0,
}

export const CHAVE_LAYOUT = 'revalle_etiquetas_layout'

const TOLERANCIA_MM = 0.01
const MSG_NAO_CABE = 'O layout não cabe numa folha A4 (210 × 297 mm).'

function inteiroEntre(n: number, min: number, max: number): boolean {
  return Number.isInteger(n) && n >= min && n <= max
}

function entre(n: number, min: number, max: number): boolean {
  return Number.isFinite(n) && n >= min && n <= max
}

/** Primeira falha do layout (na ordem dos campos) ou `null` quando está tudo certo. */
export function validarLayout(l: LayoutEtiquetas): string | null {
  if (!inteiroEntre(l.colunas, 1, 10)) return 'Colunas: informe um número inteiro de 1 a 10.'
  if (!inteiroEntre(l.linhas, 1, 30)) return 'Linhas: informe um número inteiro de 1 a 30.'
  if (!entre(l.larguraMm, 10, 210)) return 'Largura: informe de 10 a 210 mm.'
  if (!entre(l.alturaMm, 10, 297)) return 'Altura: informe de 10 a 297 mm.'
  const margens = [l.margemSuperiorMm, l.margemEsquerdaMm, l.espacoHorizontalMm, l.espacoVerticalMm]
  if (!margens.every((m) => entre(m, 0, 50))) return 'Margens e espaçamentos: informe de 0 a 50 mm.'
  const larguraTotal = l.margemEsquerdaMm + l.colunas * l.larguraMm + (l.colunas - 1) * l.espacoHorizontalMm
  const alturaTotal = l.margemSuperiorMm + l.linhas * l.alturaMm + (l.linhas - 1) * l.espacoVerticalMm
  if (larguraTotal > FOLHA_A4.larguraMm + TOLERANCIA_MM || alturaTotal > FOLHA_A4.alturaMm + TOLERANCIA_MM) {
    return MSG_NAO_CABE
  }
  return null
}

export function etiquetasPorFolha(l: LayoutEtiquetas): number {
  return l.colunas * l.linhas
}

export function validarPosicaoInicial(l: LayoutEtiquetas, p: number): string | null {
  const n = etiquetasPorFolha(l)
  return inteiroEntre(p, 1, n) ? null : `A posição inicial deve estar entre 1 e ${n}.`
}

export interface PosicaoEtiqueta {
  folha: number
  linha: number
  coluna: number
  esquerdaMm: number
  topoMm: number
}

/** Posição de cada etiqueta: linha a linha, da esquerda para a direita, a partir da posição inicial (1 = canto superior esquerdo). */
export function calcularPosicoes(l: LayoutEtiquetas, quantidade: number, posicaoInicial: number): PosicaoEtiqueta[] {
  const porFolha = etiquetasPorFolha(l)
  const posicoes: PosicaoEtiqueta[] = []
  for (let k = 0; k < quantidade; k++) {
    const slot = posicaoInicial - 1 + k
    const folha = Math.floor(slot / porFolha)
    const idx = slot % porFolha
    const linha = Math.floor(idx / l.colunas)
    const coluna = idx % l.colunas
    posicoes.push({
      folha,
      linha,
      coluna,
      esquerdaMm: l.margemEsquerdaMm + coluna * (l.larguraMm + l.espacoHorizontalMm),
      topoMm: l.margemSuperiorMm + linha * (l.alturaMm + l.espacoVerticalMm),
    })
  }
  return posicoes
}

export function totalFolhas(l: LayoutEtiquetas, quantidade: number, posicaoInicial: number): number {
  if (quantidade <= 0) return 0
  return Math.ceil((posicaoInicial - 1 + quantidade) / etiquetasPorFolha(l))
}

const CAMPOS: (keyof LayoutEtiquetas)[] = [
  'colunas',
  'linhas',
  'larguraMm',
  'alturaMm',
  'margemSuperiorMm',
  'margemEsquerdaMm',
  'espacoHorizontalMm',
  'espacoVerticalMm',
]

/** Layout salvo no navegador; qualquer problema (JSON ruim, campo faltando, valor inválido, storage bloqueado) devolve o padrão. */
export function carregarLayout(): LayoutEtiquetas {
  try {
    const bruto = localStorage.getItem(CHAVE_LAYOUT)
    if (bruto === null) return LAYOUT_PADRAO
    const dados = JSON.parse(bruto) as Record<string, unknown>
    if (!dados || typeof dados !== 'object') return LAYOUT_PADRAO
    if (!CAMPOS.every((c) => typeof dados[c] === 'number')) return LAYOUT_PADRAO
    const layout = Object.fromEntries(CAMPOS.map((c) => [c, dados[c] as number])) as unknown as LayoutEtiquetas
    return validarLayout(layout) === null ? layout : LAYOUT_PADRAO
  } catch {
    return LAYOUT_PADRAO
  }
}

/** Grava só o layout (a posição inicial volta a 1 a cada visita). Silencioso se o storage estiver indisponível. */
export function salvarLayout(l: LayoutEtiquetas): void {
  try {
    localStorage.setItem(CHAVE_LAYOUT, JSON.stringify(l))
  } catch {
    /* sem armazenamento: o layout vale só nesta visita */
  }
}
