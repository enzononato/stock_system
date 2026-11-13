/**
 * Destino a abrir depois do login. O QR da etiqueta aponta para `/e/{id}`: quem
 * ainda não está logado passa pelo login e depois cai direto no equipamento.
 * Vive no `sessionStorage` (só nesta aba) e só aceita `/items/<número>`: nada de
 * endereço livre, para não virar redirecionamento aberto.
 */
export const CHAVE_DESTINO = 'revalle_destino_pos_login'

const DESTINO_VALIDO = /^\/items\/\d+$/
const VALIDADE_MS = 10 * 60 * 1000

export function guardarDestinoPosLogin(destino: string): void {
  if (!DESTINO_VALIDO.test(destino)) return
  try {
    sessionStorage.setItem(CHAVE_DESTINO, JSON.stringify({ destino, em: Date.now() }))
  } catch {
    /* sem armazenamento: o usuário cai na tela inicial depois do login */
  }
}

/** Lê e apaga o destino guardado; `null` se não há, se expirou (10 min) ou se o valor não é válido. */
export function consumirDestinoPosLogin(agora: number = Date.now()): string | null {
  try {
    const bruto = sessionStorage.getItem(CHAVE_DESTINO)
    if (bruto === null) return null
    sessionStorage.removeItem(CHAVE_DESTINO)
    const dados = JSON.parse(bruto) as { destino?: unknown; em?: unknown }
    if (typeof dados.destino !== 'string' || !DESTINO_VALIDO.test(dados.destino)) return null
    if (typeof dados.em !== 'number' || agora - dados.em > VALIDADE_MS) return null
    return dados.destino
  } catch {
    return null
  }
}
