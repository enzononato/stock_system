import api from './client'

export interface User {
  id: number
  username: string
  role: string
  email: string | null
  /** O e-mail já foi confirmado (código no login ou link de senha). */
  email_confirmado: boolean
  /** false = convite pendente: a pessoa ainda não criou a própria senha. */
  senha_definida: boolean
  totp_ativo: boolean
}

export async function listUsers() {
  const res = await api.get('/users')
  return res.data as User[]
}

/** Cria a conta sem senha: o servidor envia o convite para o e-mail. */
export async function createUser(data: { email: string; role: string }) {
  const res = await api.post('/users', data)
  return res.data as { detail: string; convite_enviado: boolean }
}

export async function updateUserEmail(id: number, email: string | null) {
  const res = await api.put(`/users/${id}`, { email })
  return res.data as User
}

export async function zerar2fa(id: number) {
  const res = await api.delete(`/users/${id}/2fa`)
  return res.data as { detail: string }
}

export async function removeUser(id: number) {
  const res = await api.delete(`/users/${id}`)
  return res.data
}

/** Reenvia o convite (conta sem senha) ou manda um link para criar nova senha. */
export async function enviarLinkSenha(id: number) {
  const res = await api.post(`/users/${id}/enviar-link-senha`)
  return res.data as { detail: string }
}
