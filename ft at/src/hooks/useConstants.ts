import { useQuery } from "@tanstack/react-query";

import { getConstants, type Constants } from "@/api/constants";

/**
 * Constantes de domínio (`GET /api/constants`). staleTime de 5 minutos: são
 * dados que mudam pouco, mas `revendas` vem das unidades ativas e precisa
 * refletir edições feitas por outros usuários em tempo razoável.
 */
const STALE_TIME_MS = 5 * 60 * 1000;

export function useConstants(): {
  data: Constants | undefined;
  isLoading: boolean;
  centerCosts: string[];
  revendas: string[];
  setores: string[];
  equipmentTypes: string[];
  peripheralTypes: string[];
  removalReasons: string[];
  removalReasonsAttachment: Record<string, boolean>;
} {
  const { data, isLoading } = useQuery({
    queryKey: ["constants"],
    queryFn: getConstants,
    staleTime: STALE_TIME_MS,
  });

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
  };
}
