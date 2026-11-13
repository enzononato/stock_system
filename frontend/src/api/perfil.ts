import api from './client'
import type { CadastroTotp } from './auth'

export interface Perfil {
  id: number
  username: string
  role: string
  email: string | null
  totp_ativo: boolean
  mfa_obrigatorio: boolean
  mfa_disponivel: boolean
  codigos_recuperacao_restantes: number
}

export async function getPerfil() {
  const res = await api.get('/perfil')
  return res.data as Perfil
}

export async function iniciarCadastro2fa() {
  const res = await api.post('/perfil/2fa/iniciar')
  return res.data as CadastroTotp
}

export async function confirmarCadastro2fa(codigo: string) {
  const res = await api.post('/perfil/2fa/confirmar', { codigo })
  return res.data as { codigos_recuperacao: string[] }
}

export async function desativar2fa(senha: string) {
  const res = await api.post('/perfil/2fa/desativar', { senha })
  return res.data as { detail: string }
}

export async function regenerarCodigos(senha: string) {
  const res = await api.post('/perfil/2fa/codigos', { senha })
  return res.data as { codigos_recuperacao: string[] }
}
