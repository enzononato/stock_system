import { useQuery } from '@tanstack/react-query'
import { getConstants, type Constants } from '@/api/constants'

// São constantes de domínio (revendas, setores, tipos de equipamento etc.) que
// raramente mudam — um staleTime alto evita refetch a cada montagem de página e,
// como o queryKey é compartilhado, elimina o problema de StockPage, LoanPage,
// RegisterItemPage, PeripheralsPage e RemovePage buscarem (e hardcodarem) as
// mesmas listas de forma independente.
const STALE_TIME_MS = 60 * 60 * 1000 // 1 hora

/**
 * Hook central de acesso às constantes de domínio expostas por `GET /api/constants`.
 * Enquanto carrega (ou se a query falhar antes de ter dados em cache), os arrays
 * derivados caem para `[]` — e `removalReasonsAttachment` para `{}` — para que as
 * páginas que os consomem não quebrem renderizando `undefined.map(...)`.
 */
export function useConstants(): {
  data: Constants | undefined
  isLoading: boolean
  centerCosts: string[]
  revendas: string[]
  setores: string[]
  equipmentTypes: string[]
  peripheralTypes: string[]
  removalReasons: string[]
  removalReasonsAttachment: Record<string, boolean>
} {
  const { data, isLoading } = useQuery({
    queryKey: ['constants'],
    queryFn: getConstants,
    staleTime: STALE_TIME_MS,
  })

  return {
    data,
    isLoading,
    centerCosts: data?.center_costs ?? [],
    revendas: data?.revendas ?? [],
    setores: data?.setores ?? [],
    equipmentTypes: data?.equipment_types ?? [],
    peripheralTypes: data?.peripheral_types ?? [],
    removalReasons: data?.removal_reasons ?? [],
    removalReasonsAttachment: data?.removal_reasons_attachment ?? {},
  }
}
