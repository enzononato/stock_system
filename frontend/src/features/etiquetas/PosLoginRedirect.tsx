import { useEffect } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { consumirDestinoPosLogin } from '@/features/etiquetas/posLogin'

/**
 * Montado uma vez no `App`, fora das rotas: quando o login termina em `/` e há um
 * destino guardado (QR lido sem sessão), leva ao equipamento. Sem usuário não consome nada.
 */
export function PosLoginRedirect(): null {
  const { user } = useAuth()
  const { pathname } = useLocation()
  const navigate = useNavigate()

  useEffect(() => {
    if (!user || pathname !== '/') return
    const destino = consumirDestinoPosLogin()
    if (destino) navigate(destino, { replace: true })
  }, [user, pathname, navigate])

  return null
}
