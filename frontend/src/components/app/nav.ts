import {
  ArrowRightLeft,
  BarChart3,
  Boxes,
  Building2,
  FileSignature,
  FileSpreadsheet,
  History,
  Keyboard,
  Link2,
  PackagePlus,
  ShieldCheck,
  Trash2,
  Undo2,
  UserRoundCog,
  UserX,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";

export interface NavItem {
  to: string;
  label: string;
  icon: LucideIcon;
  roles?: string[];
  group: "Operação" | "Inventário" | "Gestão";
  children?: { to: string; label: string; icon: LucideIcon; roles?: string[] }[];
}

export const NAV_ITEMS: NavItem[] = [
  { to: "/", label: "Estoque", icon: Boxes, group: "Inventário" },
  {
    to: "/register",
    label: "Cadastrar item",
    icon: PackagePlus,
    roles: ["Gestor", "Técnico"],
    group: "Inventário",
  },

  {
    to: "/loan",
    label: "Empréstimo",
    icon: ArrowRightLeft,
    roles: ["Gestor", "Técnico"],
    group: "Operação",
  },
  {
    to: "/return",
    label: "Devolução",
    icon: Undo2,
    roles: ["Gestor", "Técnico"],
    group: "Operação",
  },
  {
    to: "/terms",
    label: "Termos",
    icon: FileSignature,
    roles: ["Gestor", "Técnico"],
    group: "Operação",
  },
  { to: "/remove", label: "Remover item", icon: Trash2, roles: ["Gestor"], group: "Operação" },
  { to: "/offboarding-ti", label: "Desligamentos", icon: UserX, group: "Operação" },
  {
    to: "/peripherals",
    label: "Periféricos",
    icon: Keyboard,
    roles: ["Gestor", "Técnico"],
    group: "Operação",
    children: [
      { to: "/link", label: "Vincular periféricos", icon: Link2, roles: ["Gestor", "Técnico"] },
    ],
  },

  { to: "/charts", label: "Indicadores", icon: BarChart3, group: "Gestão" },
  {
    to: "/history",
    label: "Histórico",
    icon: History,
    roles: ["Gestor", "Técnico"],
    group: "Gestão",
  },
  {
    to: "/report",
    label: "Relatório mensal",
    icon: FileSpreadsheet,
    roles: ["Gestor", "Técnico"],
    group: "Gestão",
  },
  { to: "/unidades", label: "Unidades", icon: Building2, roles: ["Gestor"], group: "Gestão" },
  { to: "/users", label: "Usuários", icon: UserRoundCog, roles: ["Gestor"], group: "Gestão" },
];

export const ROLE_ICON = ShieldCheck;
export const NAV_GROUPS: NavItem["group"][] = ["Inventário", "Operação", "Gestão"];
