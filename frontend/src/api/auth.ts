import api from './client'

export async function login(username: string, password: string) {
  const res = await api.post('/auth/login', { username, password })
  return res.data as { access_token: string }
}

export async function logout() {
  await api.post('/auth/logout')
}

export async function getMe() {
  const res = await api.get('/auth/me')
  return res.data as { id: number; username: string; role: string }
}
