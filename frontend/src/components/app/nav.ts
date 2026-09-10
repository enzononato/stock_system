import {
  ArrowRightLeft,
  BarChart3,
  Boxes,
  Building2,
  FileSpreadsheet,
  History,
  Keyboard,
  LayoutDashboard,
  PackagePlus,
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
  group: "Principal" | "Inventário" | "Operações" | "Gestão";
}

export const NAV_GROUPS: NavItem["group"][] = ["Principal", "Inventário", "Operações", "Gestão"];

export const NAV_ITEMS: NavItem[] = [
  // Principal (Seção 21)
  { to: "/", label: "Dashboard", icon: LayoutDashboard, group: "Principal" },

  // Inventário
  { to: "/stock", label: "Estoque", icon: Boxes, group: "Inventário" },
  {
    to: "/register",
    label: "Cadastro",
    icon: PackagePlus,
    roles: ["Gestor", "Técnico"],
    group: "Inventário",
  },
  {
    to: "/history",
    label: "Histórico",
    icon: History,
    roles: ["Gestor", "Técnico"],
    group: "Inventário",
  },

  // Operações
  {
    to: "/loan",
    label: "Empréstimos",
    icon: ArrowRightLeft,
    roles: ["Gestor", "Técnico"],
    group: "Operações",
  },
  {
    to: "/return",
    label: "Devoluções",
    icon: Undo2,
    roles: ["Gestor", "Técnico"],
    group: "Operações",
  },
  {
    to: "/peripherals",
    label: "Periféricos",
    icon: Keyboard,
    roles: ["Gestor", "Técnico"],
    group: "Operações",
  },
  { to: "/offboarding-ti", label: "Offboarding", icon: UserX, group: "Operações" },

  // Gestão
  { to: "/users", label: "Usuários", icon: UserRoundCog, roles: ["Gestor"], group: "Gestão" },
  { to: "/unidades", label: "Unidades", icon: Building2, roles: ["Gestor"], group: "Gestão" },
  {
    to: "/report",
    label: "Relatórios",
    icon: FileSpreadsheet,
    roles: ["Gestor", "Técnico"],
    group: "Gestão",
  },
  { to: "/charts", label: "Indicadores", icon: BarChart3, group: "Gestão" },
];
