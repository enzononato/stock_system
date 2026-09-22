# Ícones Animados da Sidebar — Stock System

13 componentes React/TSX para a sidebar do Stock System.

## Instalação

Esta pasta já faz parte do projeto, em:

frontend/src/components/animated-icons/stock-system-animated-sidebar-icons/

O CSS é importado uma única vez, no entry point da aplicação
(`frontend/src/main.tsx`):

```ts
import '@/components/animated-icons/stock-system-animated-sidebar-icons/animated-sidebar-icons.css'
```

## Uso

```tsx
import { EstoqueIcon } from '@/components/animated-icons/stock-system-animated-sidebar-icons'

<EstoqueIcon size={20} />
```

A animação dispara no hover/focus do item de menu inteiro (via classe
`group` no elemento pai, aplicada pela `Sidebar.tsx`), não no hover do
próprio SVG — assim passar o mouse sobre o texto do item também anima o
ícone. Não há dependência adicional de runtime.

## Ícones

Dashboard, Estoque, Patrimônios, Periféricos, Movimentações, Empréstimos,
Devoluções, Empresas, Usuários, Relatórios, Termos de Responsabilidade,
Configurações e Lixeira.

## Observação

Os desenhos-base seguem a linguagem visual de ícones Lucide, mas os
componentes deste pacote adicionam movimento via CSS para uso na sidebar.
