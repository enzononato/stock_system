import { useQuery } from '@tanstack/react-query'
import { getConstants, type Constants } from '@/api/constants'

// São constantes de domínio (revendas, setores, tipos de equipamento etc.) que
// em geral mudam pouco — um staleTime alto evita refetch a cada montagem de
// página e, como o queryKey é compartilhado, elimina o problema de StockPage,
// LoanPage, RegisterItemPage, PeripheralsPage e RemovePage buscarem (e
// hardcodarem) as mesmas listas de forma independente.
//
// `revendas` deixou de ser uma lista fixa: agora vem das unidades ativas no
// banco (ver UnidadesPage.tsx), editáveis pela própria tela. Com o staleTime
// de 1 hora que este hook tinha antes, criar/renomear/inativar uma unidade só
// refletiria nos seletores de Cadastrar/Emprestar depois de até uma hora — a
// UnidadesPage contorna isso invalidando explicitamente a query ['constants']
// após cada escrita, mas outra aba/sessão que não passou por essa tela (ex.:
// dois Gestores trabalhando ao mesmo tempo) só veria a mudança quando o cache
// dela expirasse. 5 minutos equilibra isso: continua evitando refetch a cada
// troca de página, mas não deixa uma unidade nova "sumida" por muito tempo em
// quem não fez a edição.
const STALE_TIME_MS = 5 * 60 * 1000 // 5 minutos

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
