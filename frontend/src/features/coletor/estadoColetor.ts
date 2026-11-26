/**
 * Estado de navegação que a página do coletor entrega ao formulário de cadastro
 * (`navigate(destino, { state })`): os campos lidos da máquina e o tipo sugerido.
 */
export type EstadoColetor = {
  origem: 'coletor'
  coletaId: number
  campos: Record<string, string>
  tipoSugerido: string | null
}

/** Valida a forma do `location.state` recebido; devolve `null` para qualquer coisa que não seja do coletor. */
export function lerEstadoColetor(state: unknown): EstadoColetor | null {
  if (typeof state !== 'object' || state === null) return null
  const s = state as Record<string, unknown>
  if (s.origem !== 'coletor') return null
  if (typeof s.coletaId !== 'number' || !Number.isFinite(s.coletaId)) return null
  if (typeof s.campos !== 'object' || s.campos === null || Array.isArray(s.campos)) return null
  const campos = s.campos as Record<string, unknown>
  if (!Object.values(campos).every((v) => typeof v === 'string')) return null
  if (s.tipoSugerido !== null && typeof s.tipoSugerido !== 'string') return null
  return {
    origem: 'coletor',
    coletaId: s.coletaId,
    campos: campos as Record<string, string>,
    tipoSugerido: s.tipoSugerido,
  }
}
