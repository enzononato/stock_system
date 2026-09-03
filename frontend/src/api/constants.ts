import api from "./client";

export interface Constants {
  center_costs: string[];
  revendas: string[];
  setores: string[];
  equipment_types: string[];
  peripheral_types: string[];
  removal_reasons: string[];
  removal_reasons_attachment: Record<string, boolean>;
}

export async function getConstants() {
  const res = await api.get("/constants");
  return res.data as Constants;
}
