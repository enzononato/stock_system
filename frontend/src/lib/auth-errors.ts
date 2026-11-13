import { isAxiosError } from 'axios'
import { getErrorMessage } from '@/lib/api-error'

export const MSG_MUITAS_TENTATIVAS = 'Muitas tentativas. Aguarde um minuto e tente de novo.'

/** Texto do 401 que o servidor manda a um Gestor que ainda não cadastrou o 2FA (sessão anterior ao cadastro). */
export const MSG_CADASTRO_2FA_OBRIGATORIO =
  'Cadastre a verificação em duas etapas para continuar. Entre novamente.'

/** 5 códigos errados no mesmo token da 2ª etapa: o servidor o revoga e é preciso entrar de novo. */
export const MSG_TOKEN_ESGOTADO = 'Muitas tentativas. Entre novamente.'
export const MSG_ETAPA_EXPIRADA = 'Etapa de verificação expirada. Entre novamente.'

/** Mensagem de erro das telas de login/senha/2FA. O 429 do slowapi não traz `detail`
 * (ganha o texto genérico), mas o 429 do bloqueio de 15 min da 2ª etapa traz: esse aparece. */
export function mensagemErroAuth(error: unknown, fallback: string): string {
  if (isAxiosError(error) && error.response?.status === 429) {
    const detail = (error.response.data as { detail?: unknown } | undefined)?.detail
    return typeof detail === 'string' && detail.trim() ? detail : MSG_MUITAS_TENTATIVAS
  }
  return getErrorMessage(error, fallback)
}

/** Se o erro diz que a etapa de verificação acabou (token esgotado ou expirado), devolve a mensagem; senão null. */
export function mensagemEtapaEncerrada(error: unknown): string | null {
  if (!isAxiosError(error) || error.response?.status !== 400) return null
  const detail = (error.response.data as { detail?: unknown } | undefined)?.detail
  return detail === MSG_TOKEN_ESGOTADO || detail === MSG_ETAPA_EXPIRADA ? detail : null
}
