import { Navigate, useParams } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { LoadingState } from '@/components/ui/StateBlocks'
import { guardarDestinoPosLogin } from '@/features/etiquetas/posLogin'

/** Destino do QR da etiqueta (`/e/{id}`): abre o equipamento, passando pelo login se preciso. */
export default function EtiquetaRedirectPage() {
  const { id } = useParams()
  const { user, isLoading } = useAuth()

  if (!id || !/^\d+$/.test(id)) return <Navigate to="/" replace />
  if (isLoading) return <LoadingState label="Abrindo equipamento..." />
  if (user) return <Navigate to={`/items/${id}`} replace />

  // Idempotente: renderizar de novo só regrava o mesmo destino.
  guardarDestinoPosLogin(`/items/${id}`)
  return <Navigate to="/login" replace />
}
