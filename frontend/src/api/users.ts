import api from './client'

export interface User {
  id: number
  username: string
  role: string
  email: string | null
  totp_ativo: boolean
}

export async function listUsers() {
  const res = await api.get('/users')
  return res.data as User[]
}

export async function createUser(data: {
  username: string
  password: string
  role: string
  email?: string | null
}) {
  const res = await api.post('/users', data)
  return res.data
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

export async function updatePassword(id: number, newPassword: string) {
  const res = await api.put(`/users/${id}/password`, { new_password: newPassword })
  return res.data
}
