import { updateColaborador, type Colaborador } from '@/api/colaboradores'
import { isValidEmail } from '@/lib/email'

/**
 * Garante que o colaborador tenha o e-mail informado gravado no cadastro (o RevSign envia o
 * link de assinatura para ele; pode ser pessoal). Vazio ou inválido levanta `Error` com a
 * mensagem para o usuário; igual ao atual (sem diferenciar caixa) não faz requisição.
 */
export async function garantirEmailDoColaborador(colaborador: Colaborador, email: string): Promise<Colaborador> {
  const valor = email.trim()
  if (!valor) throw new Error('Informe o e-mail do colaborador.')
  if (!isValidEmail(valor)) throw new Error('E-mail inválido.')
  if ((colaborador.email ?? '').trim().toLowerCase() === valor.toLowerCase()) return colaborador
  return updateColaborador(colaborador.id, { email: valor.toLowerCase() })
}
