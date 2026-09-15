# Animated Sidebar Icons — Stock System

12 componentes React/TSX para a sidebar do Stock System.

## Instalação

Copie o conteúdo desta pasta para:

E:\STOCK SYSTEM\frontend\src\components\animated-icons\

Depois importe o CSS uma vez, por exemplo no seu entry/global CSS:

import "./components/animated-icons/animated-sidebar-icons.css";

## Uso

import { EstoqueIcon } from "@/components/animated-icons";

<EstoqueIcon size={20} />

As animações são acionadas no hover/focus. Não há dependência adicional de runtime.

## Ícones

Dashboard, Estoque, Patrimônios, Periféricos, Movimentações, Empréstimos,
Devoluções, Empresas, Usuários, Relatórios, Configurações e Lixeira.

## Observação

Os desenhos-base seguem a linguagem visual de ícones Lucide, mas os componentes
deste pacote adicionam movimento via CSS para uso na sidebar.
