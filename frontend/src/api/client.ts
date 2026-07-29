import axios from 'axios'

let accessToken: string | null = null

export function setAccessToken(token: string | null) {
  accessToken = token
}

export function getAccessToken(): string | null {
  return accessToken
}

/** Contrato genérico de paginação usado por `listItems`/`listHistory` (ver api/items.ts e api/history.ts). */
export interface Paginated<T> {
  items: T[]
  total: number
}

const api = axios.create({
  baseURL: '/api',
  withCredentials: true, // Envia o cookie refresh_token
})

// Injeta o access token em cada request
api.interceptors.request.use((config) => {
  if (accessToken) {
    config.headers.Authorization = `Bearer ${accessToken}`
  }
  return config
})

// --- Callback de sessão expirada -------------------------------------------------
// O interceptor de 401 antigo fazia `window.location.href = '/login'` diretamente,
// o que descarta todo o estado do React (um hard reload). Em vez disso, expomos um
// callback que o AuthContext registra: ele limpa o usuário do contexto e deixa o
// AppLayout (que já faz `if (!user) return <Navigate to="/login" />`) navegar via
// React Router normalmente, sem recarregar a página.
type SessionExpiredHandler = () => void
let onSessionExpired: SessionExpiredHandler | null = null

/**
 * Registra o callback chamado quando a sessão expira de fato (refresh falhou).
 * Passe `null` para remover o callback (ex.: no cleanup do useEffect do AuthContext).
 */
export function setSessionExpiredHandler(handler: SessionExpiredHandler | null) {
  onSessionExpired = handler
}

function notifySessionExpired() {
  if (onSessionExpired) {
    onSessionExpired()
  } else {
    // Fallback de segurança caso nenhum contexto tenha registrado o callback ainda
    // (ex.: erro ocorrendo antes do AuthProvider montar).
    window.location.href = '/login'
  }
}

// Se receber 401, tenta renovar o token via refresh
let isRefreshing = false
let failedQueue: Array<{ resolve: (t: string) => void; reject: (e: unknown) => void }> = []

function processQueue(error: unknown, token: string | null) {
  failedQueue.forEach((p) => (error ? p.reject(error) : p.resolve(token!)))
  failedQueue = []
}

api.interceptors.response.use(
  (res) => res,
  async (error) => {
    const originalRequest = error.config
    const status = error.response?.status
    // Se a própria chamada de refresh voltar com 401, ou se a requisição já foi
    // reenviada uma vez (já usamos um token "novo" e mesmo assim deu 401), não
    // tentamos de novo — isso evitaria um loop infinito de refresh.
    const isRefreshCall = Boolean(originalRequest?.url && String(originalRequest.url).includes('/auth/refresh'))

    if (status !== 401 || isRefreshCall || originalRequest?._retry) {
      if (status === 401) {
        setAccessToken(null)
        notifySessionExpired()
      }
      return Promise.reject(error)
    }

    if (isRefreshing) {
      return new Promise((resolve, reject) => {
        failedQueue.push({ resolve, reject })
      }).then((token) => {
        originalRequest.headers.Authorization = `Bearer ${token}`
        return api(originalRequest)
      })
    }

    originalRequest._retry = true
    isRefreshing = true
    try {
      const res = await axios.post('/api/auth/refresh', {}, { withCredentials: true })
      const newToken = res.data.access_token
      setAccessToken(newToken)
      processQueue(null, newToken)
      originalRequest.headers.Authorization = `Bearer ${newToken}`
      return api(originalRequest)
    } catch (err) {
      processQueue(err, null)
      setAccessToken(null)
      notifySessionExpired()
      return Promise.reject(err)
    } finally {
      isRefreshing = false
    }
  }
)

/** Remove o prefixo `/api` de um caminho, já que a instância `api` usa baseURL: '/api'. */
function normalizeApiPath(path: string): string {
  let normalized = path.trim()
  if (!normalized.startsWith('/')) normalized = `/${normalized}`
  if (normalized === '/api' || normalized.startsWith('/api/')) {
    normalized = normalized.slice(4) || '/'
  }
  return normalized
}

/**
 * Baixa um recurso protegido por Bearer token e dispara o download no navegador.
 *
 * Necessário porque mecanismos nativos do navegador (`<a href download>`,
 * `window.open(url)`, navegação direta) NÃO enviam o header Authorization —
 * eles fazem uma requisição de navegação "crua", não passam pelos interceptors
 * do axios. Qualquer endpoint protegido por Bearer chamado dessa forma sempre
 * retorna 401. Por isso buscamos o arquivo via axios autenticado (responseType:
 * 'blob') e só então criamos um object URL local para o navegador baixar.
 *
 * Aceita tanto caminhos com prefixo `/api` (ex: `/api/documents/files/x/y.pdf`,
 * como os devolvidos por endpoints do backend) quanto sem esse prefixo (ex:
 * `/documents/files/x/y.pdf`), normalizando para não duplicar o baseURL.
 */
export async function downloadAuthenticated(path: string, filename: string): Promise<void> {
  const url = normalizeApiPath(path)
  const res = await api.get(url, { responseType: 'blob' })
  const blobUrl = URL.createObjectURL(res.data as Blob)

  const link = document.createElement('a')
  link.href = blobUrl
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()

  // Revoga o object URL com um pequeno atraso. Revogar de forma síncrona logo
  // após o click() pode cancelar o download em alguns navegadores (o download
  // via blob é iniciado de forma assíncrona), e nunca revogar vaza memória.
  window.setTimeout(() => URL.revokeObjectURL(blobUrl), 1000)
}

export default api
