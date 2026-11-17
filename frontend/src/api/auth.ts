import api from './client'

export type EtapaLogin = 'ok' | 'mfa' | 'cadastro_2fa'

export interface LoginResponse {
  etapa: EtapaLogin
  access_token?: string | null
  token_type?: string
  mfa_token?: string | null
  expira_em_segundos?: number | null
  codigos_recuperacao_restantes?: number | null
  codigos_recuperacao?: string[] | null
}

export interface CadastroTotp {
  segredo: string
  otpauth_uri: string
}

export async function login(username: string, password: string) {
  const res = await api.post('/auth/login', { username, password })
  return res.data as LoginResponse
}

export async function verificarMfa(mfaToken: string, codigo: string) {
  const res = await api.post('/auth/mfa/verificar', { mfa_token: mfaToken, codigo })
  return res.data as LoginResponse
}

export async function iniciarCadastroMfa(mfaToken: string) {
  const res = await api.post('/auth/mfa/cadastro/iniciar', { mfa_token: mfaToken })
  return res.data as CadastroTotp
}

export async function confirmarCadastroMfa(mfaToken: string, codigo: string) {
  const res = await api.post('/auth/mfa/cadastro/confirmar', { mfa_token: mfaToken, codigo })
  return res.data as LoginResponse
}

export async function esqueciSenha(identificador: string) {
  const res = await api.post('/auth/forgot', { identificador })
  return res.data as { detail: string }
}

export async function redefinirSenha(token: string, novaSenha: string) {
  const res = await api.post('/auth/reset', { token, nova_senha: novaSenha })
  return res.data as { detail: string }
}

export async function logout() {
  await api.post('/auth/logout')
}

export async function getMe() {
  const res = await api.get('/auth/me')
  return res.data as { id: number; username: string; role: string }
}
