import { isAxiosError } from 'axios'
import { getErrorMessage } from '@/lib/api-error'

export const MSG_MUITAS_TENTATIVAS = 'Muitas tentativas. Aguarde um minuto e tente de novo.'

/** Texto do 401 que o servidor manda a um Gestor que ainda não cadastrou o 2FA (sessão anterior ao cadastro). */
export const MSG_CADASTRO_2FA_OBRIGATORIO =
  'Cadastre a verificação em duas etapas para continuar. Entre novamente.'

/** Mensagem de erro das telas de login/senha/2FA. O 429 do slowapi não traz
 * `detail`, por isso ganha texto próprio; o resto segue `getErrorMessage`. */
export function mensagemErroAuth(error: unknown, fallback: string): string {
  if (isAxiosError(error) && error.response?.status === 429) return MSG_MUITAS_TENTATIVAS
  return getErrorMessage(error, fallback)
}
