import api from './client'

export interface User {
  id: number
  username: string
  role: string
}

export async function listUsers() {
  const res = await api.get('/users')
  return res.data as User[]
}

export async function createUser(data: { username: string; password: string; role: string }) {
  const res = await api.post('/users', data)
  return res.data
}

export async function removeUser(id: number) {
  const res = await api.delete(`/users/${id}`)
  return res.data
}

export async function updatePassword(id: number, newPassword: string) {
  const res = await api.put(`/users/${id}/password`, { new_password: newPassword })
  return res.data
}
