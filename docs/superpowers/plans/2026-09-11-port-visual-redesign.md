# Port do Redesign Visual (Enterprise Editorial) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Aplicar a identidade visual monocromática "Enterprise Editorial" da branch `redesign-frontend` sobre a estrutura atual do frontend, sem trocar stack, sem remover funcionalidade e sem quebrar deploy ou testes.

**Architecture:** A branch `redesign-frontend` é uma reescrita completa (Tailwind 4, TanStack Start/Router, Vite 8, sem vitest, sem Dockerfile/nginx) e **não será mesclada**. Ela serve só como referência de design. O port é manual: extraímos os tokens de design e a linguagem visual, e reescrevemos apenas o *interior* dos componentes atuais. **Regra central: APIs públicas de componentes não mudam** — `Badge` continua aceitando `variant="success" showDot`, `Button` continua aceitando `variant="gradient"`, etc. Assim as páginas seguem compilando, e a mudança visual chega por dentro.

**Tech Stack:** React 18, TypeScript, Vite 5, Tailwind CSS 3 (config JS, não CSS-first), react-router-dom 6, TanStack Query/Table, Radix UI, vitest + Testing Library, class-variance-authority.

**Spec:** Este documento. A referência visual é a branch `origin/redesign-frontend`, arquivos `frontend/src/styles.css` (tokens), `frontend/src/components/ui/*` (primitivos) e `frontend/src/components/layout/*` (shell). Consulte com `git show origin/redesign-frontend:<caminho>`.

## Global Constraints

- **Não mesclar `redesign-frontend`.** Nenhum `git merge`, `cherry-pick` ou `checkout` de arquivo daquela branch. Ler com `git show` apenas.
- **Não apagar nada:** `frontend/Dockerfile`, `frontend/nginx.conf.template`, `frontend/vitest.config.ts`, `frontend/src/test/**`, todos os `*.test.ts(x)`, `src/pages/**`, `src/App.tsx`, `react-router-dom`.
- **Não adicionar dependências.** Nada de `three`, `cmdk`, `@tanstack/react-router`, `tw-animate-css`, `sonner`. Tudo que o plano pede já existe em `frontend/package.json`.
- **Tailwind permanece na versão 3.** Sintaxe v4 (`@theme`, `@utility`, `@custom-variant`, `source(none)`) é proibida. Traduzir para `tailwind.config.js` + `@layer` no CSS.
- **Fora de escopo:** módulo offboarding, página Dashboard, logo 3D (`ThreeLogoCanvas`, `public/draco/**`, `.glb`), command palette, `AssetDetailsPanel`, breadcrumbs.
- **Paleta:** interface 100% monocromática (preto/branco/cinza). **Exceção única:** badges de status e mensagens de erro/ação destrutiva mantêm cor semântica dessaturada (tokens `--status-*` da Task 1).
- **Idioma:** todo texto de interface em português do Brasil. Comentários de código em português.
- **Cada task termina com a suíte inteira verde:** `npx tsc --noEmit`, `npm run test:run`, `npx vite build`.
- Comandos rodam de `C:\Users\enzo.jz\Documents\stockweb\stock_system\frontend`.

**Linha de base medida antes do port (2026-09-11):** 10 arquivos de teste, 107 testes, todos passando. `tsc --noEmit` limpo, `vite build` limpo. A paleta `brand` do `tailwind.config.js` foi verificada e **não é usada por nenhum arquivo** — removê-la é seguro. As classes `glass-card`/`glass-panel` aparecem em `ItemDetailsModal.tsx`, `DataTable.tsx`, `ChartsPage.tsx`, `StockPage.tsx` e `index.css`; `shadow-glow-*` em `Sidebar.tsx` e `FileUpload.tsx` — todas cobertas pelas tasks abaixo.

## Tabela de Mapeamento (a regra que todas as tasks de restyle aplicam)

Esta tabela é a especificação de tradução. Onde uma classe da coluna esquerda aparecer, substituir pela direita.

| Antes (atual) | Depois (monocromático) |
|---|---|
| `bg-indigo-600`, `bg-indigo-500` | `bg-primary` |
| `text-white` (sobre primary) | `text-primary-foreground` |
| `text-indigo-600`, `text-indigo-700`, `text-indigo-400` | `text-foreground` |
| `bg-gradient-to-r from-indigo-* ... to-violet-*` | `bg-primary` (gradientes somem) |
| `bg-slate-950`, `bg-slate-900` | `bg-surface` |
| `bg-slate-50`, `bg-slate-50/70`, `bg-white`, `bg-white/85` | `bg-surface` |
| `bg-slate-100`, `bg-slate-100/80`, `bg-slate-200/80` | `bg-surface-alt` |
| `text-slate-900`, `text-slate-800`, `text-slate-700` | `text-foreground` |
| `text-slate-500`, `text-slate-400`, `text-slate-300` | `text-muted-foreground` |
| `border-slate-200`, `border-slate-200/80`, `border-slate-800` | `border-border` |
| `border-indigo-*` | `border-border-strong` |
| `rounded-2xl`, `rounded-xl` | `rounded-md` |
| `rounded-lg` | `rounded` |
| `rounded-full` (em badge/pill) | `rounded-sm` |
| `rounded-full` (em avatar/dot) | mantém `rounded-full` |
| `glass-card`, `glass-panel` | `surface-panel` |
| `shadow-md`, `shadow-lg`, `shadow-2xl`, `shadow-glow-*`, `shadow-glass` | remover (design não usa sombra; só `shadow-overlay` em modal/popover) |
| `backdrop-blur-*` | remover |
| `bg-[radial-gradient(...)]` e demais gradientes decorativos | remover |
| `font-bold` em título | `font-semibold` |
| `animate-pulse` (em dot de status) | manter |
| `focus-visible:ring-2 focus-visible:ring-indigo-500/50` | `focus-visible:ring-1 focus-visible:ring-ring` |
| `active:scale-[0.98]`, `hover:-translate-y-*`, `hover:scale-*` | remover (sem movimento decorativo) |

**Cor semântica sobrevive apenas** via as classes utilitárias `badge-status-*` definidas na Task 1 e via `text-destructive` / `border-destructive` em erro e ação destrutiva. Nenhum outro `emerald-`/`rose-`/`amber-`/`sky-`/`purple-`/`indigo-`/`violet-`/`slate-` pode restar no código ao fim do plano.

### Exemplo trabalhado da tradução

Para não restar dúvida sobre o que "aplicar a tabela" significa na prática — este é um trecho real de `StockPage.tsx`:

```tsx
// ANTES
<div className="glass-card rounded-2xl p-5 shadow-lg border border-slate-200/80 hover:-translate-y-0.5 transition-all">
  <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">Total de Itens</p>
  <p className="mt-2 text-3xl font-bold text-slate-900">{total}</p>
  <span className="mt-3 inline-flex items-center gap-1 rounded-full bg-emerald-50 px-2 py-0.5 text-xs font-medium text-emerald-700">
    <TrendingUp className="h-3 w-3" /> Ativo
  </span>
</div>

// DEPOIS
<div className="surface-panel p-5">
  <p className="text-caption text-muted-foreground">Total de Itens</p>
  <p className="mt-2 text-heading-lg num text-foreground">{total}</p>
  <Badge variant="success" showDot className="mt-3">Ativo</Badge>
</div>
```

O que aconteceu: `glass-card rounded-2xl shadow-lg border border-slate-200/80` colapsou em `surface-panel` (que já traz superfície, borda e raio); o `hover:-translate-y-0.5` sumiu (sem movimento decorativo); o par rótulo/valor virou `text-caption` + `text-heading-lg num`; e a pílula manual de cor virou o `Badge` do sistema, que é o único lugar autorizado a carregar cor semântica.

---

### Task 1: Tokens de design e utilitários base

**Files:**
- Modify: `frontend/tailwind.config.js` (arquivo inteiro reescrito)
- Modify: `frontend/src/index.css` (arquivo inteiro reescrito)
- Modify: `frontend/index.html` (linha do Google Fonts)

**Interfaces:**
- Consumes: nada (primeira task)
- Produces: tokens de cor (`bg-canvas`, `bg-surface`, `bg-surface-alt`, `border-border`, `border-border-strong`, `text-foreground`, `text-muted-foreground`, `bg-primary`, `text-primary-foreground`, `bg-sidebar`, `text-destructive`), raio (`rounded-sm|DEFAULT|md|lg` = 2/4/6/8px), duração (`duration-micro` = 150ms), utilitários de tipografia (`text-caption`, `text-body-sm`, `text-body`, `text-body-lg`, `text-heading-sm`, `text-heading`, `text-heading-lg`, `num`), painéis (`surface-panel`, `surface-interactive`, `figure-ground-panel`, `page-container-dense`, `page-container-reading`) e status (`badge-status-available`, `badge-status-loaned`, `badge-status-pending`, `badge-status-return`, `badge-status-danger`, `badge-status-neutral`).

- [ ] **Step 1: Escrever o teste que trava os tokens**

Criar `frontend/src/lib/tokens.test.ts`:

```ts
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const css = readFileSync(resolve(__dirname, '../index.css'), 'utf8')
const config = readFileSync(resolve(__dirname, '../../tailwind.config.js'), 'utf8')

describe('tokens de design', () => {
  it('define a paleta monocromática no modo claro', () => {
    expect(css).toContain('--canvas: #F7F7F5')
    expect(css).toContain('--surface: #FFFFFF')
    expect(css).toContain('--surface-alt: #F0F0ED')
    expect(css).toContain('--border: #D9D9D4')
    expect(css).toContain('--border-strong: #BDBDB7')
    expect(css).toContain('--text-primary: #111111')
  })

  it('define a paleta do modo escuro', () => {
    expect(css).toContain('--canvas: #050505')
    expect(css).toContain('--surface: #0C0C0C')
    expect(css).toContain('--text-primary: #F2F2F0')
  })

  it('define os raios editoriais (2/4/6/8px)', () => {
    expect(css).toContain('--radius-sm: 2px')
    expect(css).toContain('--radius: 4px')
    expect(css).toContain('--radius-md: 6px')
    expect(css).toContain('--radius-lg: 8px')
  })

  it('mantém cor semântica apenas para status', () => {
    expect(css).toContain('--status-available')
    expect(css).toContain('--status-loaned')
    expect(css).toContain('--status-pending')
    expect(css).toContain('--status-return')
    expect(css).toContain('--status-danger')
  })

  it('expõe os utilitários de tipografia editorial', () => {
    for (const u of ['.text-caption', '.text-body-sm', '.text-body', '.text-heading', '.num']) {
      expect(css).toContain(u)
    }
  })

  it('expõe os utilitários de superfície', () => {
    for (const u of ['.surface-panel', '.surface-interactive', '.figure-ground-panel']) {
      expect(css).toContain(u)
    }
  })

  it('não usa sintaxe do Tailwind 4', () => {
    expect(css).not.toContain('@theme')
    expect(css).not.toContain('@utility')
    expect(css).not.toContain('@custom-variant')
    expect(css).not.toContain('source(none)')
  })

  it('registra os tokens novos no tailwind.config.js', () => {
    for (const t of ['canvas', 'surface', 'border-strong', 'sidebar', 'micro']) {
      expect(config).toContain(t)
    }
  })

  it('não deixa a paleta antiga no config', () => {
    expect(config).not.toContain('glow-indigo')
    expect(config).not.toContain('#6366f1')
  })
})
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

Run: `npx vitest run src/lib/tokens.test.ts`
Expected: FAIL — os tokens ainda não existem em `index.css`.

- [ ] **Step 3: Reescrever `frontend/src/index.css`**

Substituir o arquivo inteiro por:

```css
@tailwind base;
@tailwind components;
@tailwind utilities;

/*
 * Design System — Stock System / Revalle
 * Enterprise Editorial — monocromático (preto, branco, cinzas).
 * Contraste, tipografia e densidade como identidade visual.
 *
 * Exceção deliberada ao monocromático: os tokens --status-* abaixo.
 * Num sistema de inventário a cor do status é sinal funcional (varrer uma
 * tabela de 88 itens sem ela fica mais lento), então eles sobrevivem em
 * versão dessaturada, calibrada para não brigar com a base neutra.
 */

@layer base {
  :root {
    --canvas: #F7F7F5;
    --surface: #FFFFFF;
    --surface-alt: #F0F0ED;
    --border: #D9D9D4;
    --border-strong: #BDBDB7;
    --text-primary: #111111;
    --text-secondary: #5F5F5A;
    --text-muted: #85857E;

    --background: var(--canvas);
    --foreground: var(--text-primary);
    --card: var(--surface);
    --card-foreground: var(--text-primary);
    --popover: var(--surface);
    --popover-foreground: var(--text-primary);

    --primary: #111111;
    --primary-foreground: #FFFFFF;
    --secondary: var(--surface-alt);
    --secondary-foreground: var(--text-primary);
    --muted: var(--surface-alt);
    --muted-foreground: var(--text-secondary);
    --accent: var(--surface-alt);
    --accent-foreground: var(--text-primary);

    --destructive: #A32B2B;
    --destructive-foreground: #FFFFFF;

    --input: var(--border);
    --ring: #111111;

    --sidebar: var(--surface);
    --sidebar-foreground: var(--text-primary);
    --sidebar-accent: var(--surface-alt);
    --sidebar-accent-foreground: var(--text-primary);
    --sidebar-border: var(--border);

    --radius-sm: 2px;
    --radius: 4px;
    --radius-md: 6px;
    --radius-lg: 8px;

    --shadow-overlay: 0 2px 12px rgba(0, 0, 0, 0.08);

    /* Status — única cor semântica do sistema, dessaturada de propósito */
    --status-available: #1F6F4A;
    --status-available-bg: #E9F1EC;
    --status-loaned: #1F5474;
    --status-loaned-bg: #E8EFF4;
    --status-pending: #7A5E17;
    --status-pending-bg: #F4EEDD;
    --status-return: #53397E;
    --status-return-bg: #EDE9F4;
    --status-danger: #A32B2B;
    --status-danger-bg: #F6E8E8;
    --status-neutral: var(--text-secondary);
    --status-neutral-bg: var(--surface-alt);
  }

  .dark {
    --canvas: #050505;
    --surface: #0C0C0C;
    --surface-alt: #131313;
    --border: #272727;
    --border-strong: #3A3A3A;
    --text-primary: #F2F2F0;
    --text-secondary: #A4A4A0;
    --text-muted: #70706C;

    --background: var(--canvas);
    --foreground: var(--text-primary);
    --card: var(--surface);
    --card-foreground: var(--text-primary);
    --popover: var(--surface);
    --popover-foreground: var(--text-primary);

    --primary: #F2F2F0;
    --primary-foreground: #080808;
    --secondary: var(--surface-alt);
    --secondary-foreground: var(--text-primary);
    --muted: var(--surface-alt);
    --muted-foreground: var(--text-secondary);
    --accent: var(--surface-alt);
    --accent-foreground: var(--text-primary);

    --destructive: #E4A0A0;
    --destructive-foreground: #080808;

    --input: var(--border);
    --ring: #F2F2F0;

    --sidebar: var(--surface);
    --sidebar-foreground: var(--text-primary);
    --sidebar-accent: var(--surface-alt);
    --sidebar-accent-foreground: var(--text-primary);
    --sidebar-border: var(--border);

    --shadow-overlay: 0 4px 20px rgba(0, 0, 0, 0.6);

    --status-available: #6FBF93;
    --status-available-bg: #12261C;
    --status-loaned: #7FB3D5;
    --status-loaned-bg: #101F2A;
    --status-pending: #D4B468;
    --status-pending-bg: #2A2212;
    --status-return: #AE96D9;
    --status-return-bg: #1E1830;
    --status-danger: #E4A0A0;
    --status-danger-bg: #2B1414;
    --status-neutral: var(--text-secondary);
    --status-neutral-bg: var(--surface-alt);
  }

  * {
    border-color: var(--border);
  }

  html {
    -webkit-text-size-adjust: 100%;
  }

  body {
    background-color: var(--background);
    color: var(--foreground);
    font-family: theme('fontFamily.sans');
    -webkit-font-smoothing: antialiased;
    transition:
      background-color 200ms cubic-bezier(0.16, 1, 0.3, 1),
      border-color 200ms cubic-bezier(0.16, 1, 0.3, 1),
      color 200ms cubic-bezier(0.16, 1, 0.3, 1);
  }

  :focus-visible {
    outline: 2px solid var(--ring);
    outline-offset: 2px;
  }
}

@layer utilities {
  /* Escala tipográfica editorial */
  .text-caption {
    font-size: 0.75rem;
    line-height: 1rem;
    font-weight: 500;
    letter-spacing: 0.05em;
    text-transform: uppercase;
  }
  .text-body-sm {
    font-size: 0.8125rem;
    line-height: 1.25rem;
    font-weight: 400;
  }
  .text-body {
    font-size: 0.875rem;
    line-height: 1.25rem;
    font-weight: 400;
  }
  .text-body-lg {
    font-size: 1rem;
    line-height: 1.5rem;
    font-weight: 500;
  }
  .text-heading-sm {
    font-size: 1.125rem;
    line-height: 1.5rem;
    font-weight: 600;
    letter-spacing: -0.01em;
  }
  .text-heading {
    font-size: 1.375rem;
    line-height: 1.75rem;
    font-weight: 600;
    letter-spacing: -0.02em;
  }
  .text-heading-lg {
    font-size: 1.75rem;
    line-height: 2rem;
    font-weight: 600;
    letter-spacing: -0.025em;
  }

  /* Números tabulares para colunas de dados */
  .num {
    font-family: theme('fontFamily.mono');
    font-variant-numeric: tabular-nums;
  }

  /* Containers de largura máxima */
  .page-container-dense {
    max-width: 1440px;
    margin-left: auto;
    margin-right: auto;
    width: 100%;
  }
  .page-container-reading {
    max-width: 1120px;
    margin-left: auto;
    margin-right: auto;
    width: 100%;
  }

  /* Superfícies — sem sombra, separação por linha */
  .surface-panel {
    background-color: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius-md);
  }
  .surface-interactive {
    transition: background-color 150ms cubic-bezier(0.16, 1, 0.3, 1),
      border-color 150ms cubic-bezier(0.16, 1, 0.3, 1);
  }
  .surface-interactive:hover {
    background-color: var(--surface-alt);
    border-color: var(--border-strong);
  }

  /* Painel de alto contraste para decisão crítica */
  .figure-ground-panel {
    background-color: var(--surface);
    border: 1px solid var(--border-strong);
    border-radius: var(--radius-md);
    padding: 1.5rem;
  }

  /* Status — cor semântica dessaturada */
  .badge-status-available {
    color: var(--status-available);
    background-color: var(--status-available-bg);
    border-color: var(--status-available);
  }
  .badge-status-loaned {
    color: var(--status-loaned);
    background-color: var(--status-loaned-bg);
    border-color: var(--status-loaned);
  }
  .badge-status-pending {
    color: var(--status-pending);
    background-color: var(--status-pending-bg);
    border-color: var(--status-pending);
  }
  .badge-status-return {
    color: var(--status-return);
    background-color: var(--status-return-bg);
    border-color: var(--status-return);
  }
  .badge-status-danger {
    color: var(--status-danger);
    background-color: var(--status-danger-bg);
    border-color: var(--status-danger);
  }
  .badge-status-neutral {
    color: var(--status-neutral);
    background-color: var(--status-neutral-bg);
    border-color: var(--border);
  }
}

/* Barra de rolagem discreta */
::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}
::-webkit-scrollbar-track {
  background: transparent;
}
::-webkit-scrollbar-thumb {
  background: var(--border-strong);
  border-radius: var(--radius-sm);
}
::-webkit-scrollbar-thumb:hover {
  background: var(--text-muted);
}

@media (prefers-reduced-motion: reduce) {
  *,
  *::before,
  *::after {
    animation-duration: 0.001ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.001ms !important;
    scroll-behavior: auto !important;
  }
}
```

- [ ] **Step 4: Reescrever `frontend/tailwind.config.js`**

Substituir o arquivo inteiro por:

```js
/** @type {import('tailwindcss').Config} */
export default {
  darkMode: ['class'],
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    container: {
      center: true,
      padding: '2rem',
      screens: { '2xl': '1440px' },
    },
    extend: {
      fontFamily: {
        sans: ['"Plus Jakarta Sans"', 'ui-sans-serif', 'system-ui', 'sans-serif'],
        heading: ['"Plus Jakarta Sans"', 'sans-serif'],
        mono: ['"IBM Plex Mono"', 'ui-monospace', 'SFMono-Regular', 'monospace'],
      },
      // Sem sombra decorativa: o design separa superfícies por linha, não por
      // elevação. A única sombra é a de overlay (modal/popover).
      boxShadow: {
        overlay: 'var(--shadow-overlay)',
      },
      transitionDuration: {
        micro: '150ms',
      },
      colors: {
        canvas: 'var(--canvas)',
        surface: {
          DEFAULT: 'var(--surface)',
          alt: 'var(--surface-alt)',
        },
        hairline: 'var(--border)',
        'border-strong': 'var(--border-strong)',
        border: 'var(--border)',
        input: 'var(--input)',
        ring: 'var(--ring)',
        background: 'var(--background)',
        foreground: 'var(--foreground)',
        primary: {
          DEFAULT: 'var(--primary)',
          foreground: 'var(--primary-foreground)',
        },
        secondary: {
          DEFAULT: 'var(--secondary)',
          foreground: 'var(--secondary-foreground)',
        },
        destructive: {
          DEFAULT: 'var(--destructive)',
          foreground: 'var(--destructive-foreground)',
        },
        muted: {
          DEFAULT: 'var(--muted)',
          foreground: 'var(--muted-foreground)',
        },
        accent: {
          DEFAULT: 'var(--accent)',
          foreground: 'var(--accent-foreground)',
        },
        card: {
          DEFAULT: 'var(--card)',
          foreground: 'var(--card-foreground)',
        },
        popover: {
          DEFAULT: 'var(--popover)',
          foreground: 'var(--popover-foreground)',
        },
        sidebar: {
          DEFAULT: 'var(--sidebar)',
          foreground: 'var(--sidebar-foreground)',
          accent: 'var(--sidebar-accent)',
          'accent-foreground': 'var(--sidebar-accent-foreground)',
          border: 'var(--sidebar-border)',
        },
      },
      borderRadius: {
        sm: 'var(--radius-sm)',
        DEFAULT: 'var(--radius)',
        md: 'var(--radius-md)',
        lg: 'var(--radius-lg)',
      },
      keyframes: {
        'accordion-down': {
          from: { height: '0' },
          to: { height: 'var(--radix-accordion-content-height)' },
        },
        'accordion-up': {
          from: { height: 'var(--radix-accordion-content-height)' },
          to: { height: '0' },
        },
        'fade-in': {
          from: { opacity: '0' },
          to: { opacity: '1' },
        },
      },
      animation: {
        'accordion-down': 'accordion-down 0.2s ease-out',
        'accordion-up': 'accordion-up 0.2s ease-out',
        'fade-in': 'fade-in 0.2s cubic-bezier(0.16, 1, 0.3, 1)',
      },
    },
  },
  plugins: [require('tailwindcss-animate')],
}
```

- [ ] **Step 5: Acrescentar a fonte mono no `frontend/index.html`**

Trocar a linha do `<link href="https://fonts.googleapis.com/css2?...">` por:

```html
    <link
      href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap"
      rel="stylesheet"
    />
```

- [ ] **Step 6: Rodar o teste de tokens**

Run: `npx vitest run src/lib/tokens.test.ts`
Expected: PASS (9 testes)

- [ ] **Step 7: Rodar a suíte inteira e o build**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tsc sem erros; testes 100% verdes (107 anteriores + 9 novos = 116); build limpo.

Se algum teste existente quebrar aqui, é sinal de que ele dependia de classe da paleta antiga — anotar qual e corrigir o teste na task do componente correspondente, não aqui.

- [ ] **Step 8: Commit**

```bash
git add frontend/tailwind.config.js frontend/src/index.css frontend/index.html frontend/src/lib/tokens.test.ts
git commit -m "feat(ui): tokens de design monocromaticos (Enterprise Editorial)"
```

---

### Task 2: Modo escuro (provider + persistência)

**Files:**
- Create: `frontend/src/lib/theme.tsx`
- Create: `frontend/src/lib/theme.test.tsx`
- Modify: `frontend/src/main.tsx`

**Interfaces:**
- Consumes: tokens `.dark` da Task 1.
- Produces: `ThemeProvider` (componente), `useTheme(): { theme: 'light' | 'dark'; toggle: () => void; setTheme: (t: 'light' | 'dark') => void }`. A Task 7 consome `useTheme` para o botão na TopBar.

- [ ] **Step 1: Escrever o teste**

Criar `frontend/src/lib/theme.test.tsx`:

```tsx
import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen, act } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ThemeProvider, useTheme } from './theme'

function Sonda() {
  const { theme, toggle } = useTheme()
  return (
    <div>
      <span data-testid="tema">{theme}</span>
      <button onClick={toggle}>alternar</button>
    </div>
  )
}

describe('ThemeProvider', () => {
  beforeEach(() => {
    localStorage.clear()
    document.documentElement.classList.remove('dark')
  })

  it('começa no modo claro quando não há preferência salva', () => {
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('light')
    expect(document.documentElement).not.toHaveClass('dark')
  })

  it('alterna para escuro e aplica a classe no html', async () => {
    const user = userEvent.setup()
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    await user.click(screen.getByRole('button', { name: 'alternar' }))
    expect(screen.getByTestId('tema')).toHaveTextContent('dark')
    expect(document.documentElement).toHaveClass('dark')
  })

  it('persiste a escolha no localStorage', async () => {
    const user = userEvent.setup()
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    await user.click(screen.getByRole('button', { name: 'alternar' }))
    expect(localStorage.getItem('tema')).toBe('dark')
  })

  it('restaura a preferência salva ao montar', () => {
    localStorage.setItem('tema', 'dark')
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('dark')
    expect(document.documentElement).toHaveClass('dark')
  })

  it('ignora valor inválido no localStorage e cai no claro', () => {
    localStorage.setItem('tema', 'roxo')
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('light')
  })

  it('useTheme fora do provider lança erro explicativo', () => {
    expect(() => render(<Sonda />)).toThrow(/ThemeProvider/)
  })
})
```

- [ ] **Step 2: Rodar e confirmar falha**

Run: `npx vitest run src/lib/theme.test.tsx`
Expected: FAIL — módulo `./theme` não existe.

- [ ] **Step 3: Criar `frontend/src/lib/theme.tsx`**

```tsx
import { createContext, useCallback, useContext, useEffect, useState, type ReactNode } from 'react'

type Tema = 'light' | 'dark'

const CHAVE_ARMAZENAMENTO = 'tema'

interface ThemeContextValue {
  theme: Tema
  toggle: () => void
  setTheme: (tema: Tema) => void
}

const ThemeContext = createContext<ThemeContextValue | null>(null)

/**
 * Lê a preferência salva. Qualquer valor que não seja 'light' ou 'dark'
 * (lixo, versão antiga, edição manual) cai no claro em vez de quebrar.
 * localStorage pode lançar em janela anônima — por isso o try/catch.
 */
function lerTemaSalvo(): Tema {
  try {
    const salvo = localStorage.getItem(CHAVE_ARMAZENAMENTO)
    return salvo === 'dark' ? 'dark' : 'light'
  } catch {
    return 'light'
  }
}

function aplicarNoDocumento(tema: Tema) {
  document.documentElement.classList.toggle('dark', tema === 'dark')
}

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setThemeState] = useState<Tema>(() => lerTemaSalvo())

  useEffect(() => {
    aplicarNoDocumento(theme)
    try {
      localStorage.setItem(CHAVE_ARMAZENAMENTO, theme)
    } catch {
      // Sem persistência (janela anônima, storage bloqueado): o tema ainda
      // funciona na sessão atual, só não sobrevive ao reload.
    }
  }, [theme])

  const setTheme = useCallback((tema: Tema) => setThemeState(tema), [])
  const toggle = useCallback(() => setThemeState((t) => (t === 'dark' ? 'light' : 'dark')), [])

  return (
    <ThemeContext.Provider value={{ theme, toggle, setTheme }}>{children}</ThemeContext.Provider>
  )
}

export function useTheme(): ThemeContextValue {
  const ctx = useContext(ThemeContext)
  if (!ctx) throw new Error('useTheme deve ser usado dentro de ThemeProvider')
  return ctx
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `npx vitest run src/lib/theme.test.tsx`
Expected: PASS (6 testes)

- [ ] **Step 5: Envolver a aplicação com o provider**

Em `frontend/src/main.tsx`, importar `ThemeProvider` de `@/lib/theme` e envolver a árvore existente — o `ThemeProvider` fica **por fora** do `BrowserRouter` e do `QueryClientProvider`, para que o tema se aplique também às telas de erro e de carregamento.

- [ ] **Step 6: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde (122 testes).

- [ ] **Step 7: Commit**

```bash
git add frontend/src/lib/theme.tsx frontend/src/lib/theme.test.tsx frontend/src/main.tsx
git commit -m "feat(ui): modo escuro com persistencia em localStorage"
```

---

### Task 3: Primitivos — Button, Badge, StatusBadge, Input, Label

**Files:**
- Modify: `frontend/src/components/ui/button.tsx`
- Modify: `frontend/src/components/ui/badge.tsx`
- Modify: `frontend/src/components/ui/input.tsx`
- Modify: `frontend/src/components/ui/label.tsx`
- Create: `frontend/src/components/ui/badge.test.tsx`

**Interfaces:**
- Consumes: tokens da Task 1.
- Produces: mesmas APIs de hoje, **sem remoção de variante**. `Button` mantém `variant: default | gradient | destructive | outline | secondary | ghost | link` e `size: default | sm | lg | icon`. `Badge` mantém `variant: default | success | warning | danger | info | purple` e a prop `showDot`. `StatusBadge` mantém `{ status?: string }`.

**Contexto crítico:** `gradient` e `purple` continuam existindo porque páginas usam. No monocromático `gradient` passa a renderizar igual a `default` — manter o nome evita mexer em todas as chamadas.

- [ ] **Step 1: Escrever o teste do Badge**

Criar `frontend/src/components/ui/badge.test.tsx`:

```tsx
import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { Badge, StatusBadge } from './badge'

describe('Badge', () => {
  it('renderiza o conteúdo', () => {
    render(<Badge>Olá</Badge>)
    expect(screen.getByText('Olá')).toBeInTheDocument()
  })

  it('mantém as variantes usadas pelas páginas', () => {
    for (const v of ['default', 'success', 'warning', 'danger', 'info', 'purple'] as const) {
      const { unmount } = render(<Badge variant={v}>{v}</Badge>)
      expect(screen.getByText(v)).toBeInTheDocument()
      unmount()
    }
  })

  it('não usa mais a paleta antiga do Tailwind', () => {
    const { container } = render(<Badge variant="success">ok</Badge>)
    expect(container.innerHTML).not.toMatch(/emerald-|rose-|amber-|sky-|purple-|indigo-|slate-/)
  })

  it('usa token de status semântico no variant success', () => {
    const { container } = render(<Badge variant="success">ok</Badge>)
    expect(container.innerHTML).toContain('badge-status-available')
  })
})

describe('StatusBadge', () => {
  it('mostra travessão quando não há status', () => {
    render(<StatusBadge />)
    expect(screen.getByText('-')).toBeInTheDocument()
  })

  it.each([
    ['Disponível', 'badge-status-available'],
    ['Indisponível', 'badge-status-loaned'],
    ['Pendente Devolução', 'badge-status-return'],
    ['Pendente', 'badge-status-pending'],
  ])('mapeia %s para %s', (status, classe) => {
    const { container } = render(<StatusBadge status={status} />)
    expect(container.innerHTML).toContain(classe)
  })

  it('status desconhecido cai no neutro em vez de sumir', () => {
    const { container } = render(<StatusBadge status="Qualquer Coisa" />)
    expect(screen.getByText('Qualquer Coisa')).toBeInTheDocument()
    expect(container.innerHTML).toContain('badge-status-neutral')
  })
})
```

- [ ] **Step 2: Rodar e confirmar falha**

Run: `npx vitest run src/components/ui/badge.test.tsx`
Expected: FAIL — as classes `badge-status-*` ainda não são usadas pelo componente.

- [ ] **Step 3: Reescrever `frontend/src/components/ui/badge.tsx`**

```tsx
import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info' | 'purple'
  showDot?: boolean
}

/*
 * A interface é monocromática, mas o status não: num inventário a cor do
 * status é sinal funcional, não enfeite. As variantes abaixo apontam para os
 * tokens --status-* (dessaturados) em vez da paleta cheia do Tailwind.
 *
 * Os nomes das variantes são preservados (success/warning/danger/info/purple)
 * porque as páginas já os usam — trocar os nomes obrigaria a mexer em todas
 * as chamadas sem ganho visual nenhum.
 */
const variantClasses = {
  default: 'badge-status-neutral',
  success: 'badge-status-available',
  warning: 'badge-status-pending',
  danger: 'badge-status-danger',
  info: 'badge-status-loaned',
  purple: 'badge-status-return',
}

const dotClasses = {
  default: 'bg-[var(--status-neutral)]',
  success: 'bg-[var(--status-available)]',
  warning: 'bg-[var(--status-pending)]',
  danger: 'bg-[var(--status-danger)]',
  info: 'bg-[var(--status-loaned)]',
  purple: 'bg-[var(--status-return)]',
}

export function Badge({ children, className, variant = 'default', showDot = false }: BadgeProps) {
  return (
    <span
      className={cn(
        'inline-flex items-center gap-1.5 rounded-sm border px-2 py-0.5 text-xs font-semibold tracking-wide select-none transition-colors duration-micro',
        variantClasses[variant],
        className
      )}
    >
      {showDot && <span className={cn('h-1.5 w-1.5 rounded-full', dotClasses[variant])} />}
      {children}
    </span>
  )
}

/** Mapeia os status do backend (pt-BR) para a cor semântica correspondente. */
export function StatusBadge({ status }: { status?: string }) {
  if (!status) return <Badge>-</Badge>
  if (status === 'Disponível') return <Badge variant="success" showDot>{status}</Badge>
  if (status === 'Indisponível') return <Badge variant="info" showDot>{status}</Badge>
  if (status === 'Pendente Devolução') return <Badge variant="purple" showDot>{status}</Badge>
  if (status.startsWith('Pendente')) return <Badge variant="warning" showDot>{status}</Badge>
  return <Badge variant="default">{status}</Badge>
}
```

- [ ] **Step 4: Reescrever o `buttonVariants` em `frontend/src/components/ui/button.tsx`**

Trocar **apenas** a chamada `cva(...)` (o restante do arquivo permanece igual):

```tsx
const buttonVariants = cva(
  'inline-flex items-center justify-center gap-2 whitespace-nowrap rounded text-sm font-medium transition-colors duration-micro focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring disabled:pointer-events-none disabled:opacity-40 select-none [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0',
  {
    variants: {
      variant: {
        default: 'bg-primary text-primary-foreground hover:opacity-90',
        // Mantida por compatibilidade: várias páginas pedem "gradient".
        // No design monocromático não há gradiente, então renderiza como default.
        gradient: 'bg-primary text-primary-foreground hover:opacity-90',
        destructive:
          'bg-transparent text-destructive border border-destructive hover:bg-destructive hover:text-destructive-foreground',
        outline: 'border border-border bg-transparent text-foreground hover:bg-surface-alt',
        secondary: 'bg-surface-alt text-foreground border border-border hover:border-border-strong',
        ghost: 'text-foreground hover:bg-surface-alt',
        link: 'text-foreground underline-offset-4 hover:underline',
      },
      size: {
        default: 'h-9 px-4 py-2',
        sm: 'h-8 rounded px-3 text-xs',
        lg: 'h-10 rounded px-6 text-sm',
        icon: 'h-9 w-9',
      },
    },
    defaultVariants: { variant: 'default', size: 'default' },
  }
)
```

- [ ] **Step 5: Reescrever o `className` do Input em `frontend/src/components/ui/input.tsx`**

Trocar só a string dentro de `cn(...)`:

```tsx
        'flex h-9 w-full rounded border border-input bg-surface px-3 py-2 text-sm text-foreground transition-colors duration-micro file:border-0 file:bg-transparent file:text-sm file:font-medium placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40',
```

- [ ] **Step 6: Ajustar o Label em `frontend/src/components/ui/label.tsx`**

Trocar só a string do `cva`:

```tsx
const labelVariants = cva(
  'text-caption text-muted-foreground peer-disabled:cursor-not-allowed peer-disabled:opacity-70'
)
```

- [ ] **Step 7: Rodar os testes do Badge**

Run: `npx vitest run src/components/ui/badge.test.tsx`
Expected: PASS (9 testes)

- [ ] **Step 8: Confirmar que a paleta antiga sumiu destes arquivos**

Run (Bash):
```bash
cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-" src/components/ui/button.tsx src/components/ui/badge.tsx src/components/ui/input.tsx src/components/ui/label.tsx
```
Expected: nenhuma saída.

- [ ] **Step 9: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde (131 testes).

- [ ] **Step 10: Commit**

```bash
git add frontend/src/components/ui/button.tsx frontend/src/components/ui/badge.tsx frontend/src/components/ui/input.tsx frontend/src/components/ui/label.tsx frontend/src/components/ui/badge.test.tsx
git commit -m "feat(ui): primitivos monocromaticos (button, badge, input, label)"
```

---

### Task 4: Primitivos de formulário e overlay — Select, AlertDialog, Toast, FileUpload, SearchableSelect

**Files:**
- Modify: `frontend/src/components/ui/select.tsx`
- Modify: `frontend/src/components/ui/alert-dialog.tsx`
- Modify: `frontend/src/components/ui/toast.tsx`
- Modify: `frontend/src/components/ui/FileUpload.tsx`
- Modify: `frontend/src/components/ui/SearchableSelect.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `Button` da Task 3.
- Produces: mesmas APIs. `toast(mensagem, tipo?)` continua com assinatura `(message: string, type?: 'success' | 'error')`.

- [ ] **Step 1: Aplicar a tabela de mapeamento nos cinco arquivos**

Ler cada arquivo e substituir as classes conforme a **Tabela de Mapeamento** no topo deste plano. Pontos específicos:

- `select.tsx`: conteúdo do dropdown usa `bg-popover text-popover-foreground border border-border rounded-md shadow-overlay`; item ativo usa `bg-surface-alt`.
- `alert-dialog.tsx`: overlay `bg-black/50`; conteúdo `surface-panel shadow-overlay`; botão de ação destrutiva usa `Button variant="destructive"`.
- `toast.tsx`: sucesso vira `border-border bg-surface text-foreground`; erro vira `border-destructive bg-[var(--status-danger-bg)] text-destructive`. **O erro mantém cor** — é a exceção de ação destrutiva definida nas restrições globais.
- `FileUpload.tsx`: área de arrastar usa `border-2 border-dashed border-border rounded-md`; estado ativo usa `border-border-strong bg-surface-alt`.
- `SearchableSelect.tsx`: lista usa `surface-panel shadow-overlay`; opção destacada usa `bg-surface-alt`.

- [ ] **Step 2: Confirmar que a paleta antiga sumiu**

Run (Bash):
```bash
cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-card|glass-panel|backdrop-blur" src/components/ui/select.tsx src/components/ui/alert-dialog.tsx src/components/ui/toast.tsx src/components/ui/FileUpload.tsx src/components/ui/SearchableSelect.tsx
```
Expected: nenhuma saída.

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde (131 testes).

- [ ] **Step 4: Commit**

```bash
git add frontend/src/components/ui/
git commit -m "feat(ui): select, dialog, toast, upload e busca no estilo monocromatico"
```

---

### Task 5: DataTable

**Files:**
- Modify: `frontend/src/components/ui/DataTable.tsx`
- Modify: `frontend/src/components/ui/DataTable.test.tsx` (só se algum teste depender de classe antiga)

**Interfaces:**
- Consumes: tokens da Task 1; `Input` e `Button` da Task 3.
- Produces: API inalterada — `DataTable<TData>({ data, columns, searchPlaceholder?, className?, pagination? })`, com `pagination: { total, pageIndex, pageSize, onPageChange, search?, onSearchChange? }`.

**Contexto:** é o componente mais visto do sistema (aparece em 8 páginas). O design editorial pede tabela densa: linha divisória fina, cabeçalho em `text-caption`, números em `num` (tabulares), sem zebra e sem sombra.

- [ ] **Step 1: Aplicar o restyle**

Conforme a Tabela de Mapeamento, mais estas regras específicas:
- Wrapper: `surface-panel overflow-x-auto` (sem `shadow`).
- `<thead>`: `bg-surface-alt border-b border-border`; `<th>`: `text-caption text-muted-foreground px-4 py-2.5 text-left whitespace-nowrap`.
- `<tbody>` linhas: `border-b border-border last:border-0 hover:bg-surface-alt transition-colors duration-micro`.
- `<td>`: `px-4 py-2.5 text-body-sm whitespace-nowrap`.
- Rodapé de contagem/paginação: `text-caption text-muted-foreground`.
- Estado vazio: `text-body-sm text-muted-foreground py-10 text-center`.

- [ ] **Step 2: Rodar os testes do DataTable**

Run: `npx vitest run src/components/ui/DataTable.test.tsx`
Expected: PASS (8 testes). Se algum falhar por asserção de classe antiga, atualizar a asserção para a classe nova e registrar no próprio teste o motivo.

- [ ] **Step 3: Confirmar limpeza da paleta**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur|shadow-(md|lg|xl|2xl)" src/components/ui/DataTable.tsx`
Expected: nenhuma saída.

- [ ] **Step 4: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 5: Commit**

```bash
git add frontend/src/components/ui/DataTable.tsx frontend/src/components/ui/DataTable.test.tsx
git commit -m "feat(ui): DataTable denso no estilo editorial"
```

---

### Task 6: Componentes de equipamento — ItemDetailsModal, ConfirmacaoTermo, TypeSpecificFields

**Files:**
- Modify: `frontend/src/components/equipment/ItemDetailsModal.tsx`
- Modify: `frontend/src/components/equipment/ConfirmacaoTermo.tsx`
- Modify: `frontend/src/components/equipment/TypeSpecificFields.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `Badge`/`StatusBadge`/`Button` da Task 3; `FileUpload` da Task 4.
- Produces: APIs inalteradas.

**Contexto:** `ItemDetailsModal.tsx` concentra 98 ocorrências de cor antiga — é o arquivo mais pesado do port. É um modal de detalhes com seções (identificação, especificações, periféricos, histórico).

- [ ] **Step 1: Aplicar a tabela de mapeamento nos três arquivos**

Regras específicas:
- Modal: overlay `bg-black/50`; painel `surface-panel shadow-overlay`.
- Cabeçalho do modal: título em `text-heading-sm`, subtítulo em `text-body-sm text-muted-foreground`, borda inferior `border-b border-border`.
- Seções: título em `text-caption text-muted-foreground`, separadas por `border-t border-border`.
- Pares rótulo/valor: rótulo `text-caption text-muted-foreground`, valor `text-body-sm text-foreground`. Valores numéricos, seriais e datas recebem também a classe `num`.
- Ícones decorativos coloridos viram `text-muted-foreground`.

- [ ] **Step 2: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/components/equipment/`
Expected: nenhuma saída.

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 4: Commit**

```bash
git add frontend/src/components/equipment/
git commit -m "feat(ui): modal de detalhes e componentes de equipamento monocromaticos"
```

---

### Task 7: Shell — AppLayout, Sidebar, TopBar (com botão de tema)

**Files:**
- Modify: `frontend/src/components/layout/AppLayout.tsx`
- Modify: `frontend/src/components/layout/Sidebar.tsx`
- Modify: `frontend/src/components/layout/TopBar.tsx`

**Interfaces:**
- Consumes: `useTheme` da Task 2; `Button` da Task 3.
- Produces: shell visual final. Nenhuma mudança de API (`AppLayout` e `Sidebar` sem props; `TopBar` mantém `{ title?: string }`).

**Contexto:** a sidebar hoje é escura (`bg-slate-950`) com acento índigo. No design novo ela é **clara** (`bg-sidebar`, que é a mesma superfície branca), separada do conteúdo só por `border-r border-border`. Item ativo não usa preenchimento colorido: usa `bg-surface-alt` mais uma barra de 2px à esquerda (`border-l-2 border-foreground`). Os agrupamentos de navegação atuais (Visão Geral / Gestão de Itens / Relatórios & Auditoria / Administração) e todos os itens e permissões **permanecem exatamente como estão** — só muda a aparência.

- [ ] **Step 1: Restyle do AppLayout**

- Container: `flex h-screen bg-background text-foreground overflow-hidden font-sans`.
- Remover o `bg-[radial-gradient(...)]` e todo o gradiente.
- `<main>`: `flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8` com um filho `page-container-dense`.
- Tela de carregamento: fundo `bg-background`, spinner `text-muted-foreground`, texto `text-body-sm text-muted-foreground`. Sem `animate-ping`, sem painel de vidro.

- [ ] **Step 2: Restyle da Sidebar**

- `<aside>`: `flex flex-col w-64 h-full bg-sidebar border-r border-sidebar-border select-none shrink-0`.
- Cabeçalho da marca: altura `h-14`, borda inferior `border-b border-border`, sem gradiente e sem selo "PRO". Nome "Revalle" em `text-sm font-semibold tracking-tight`, subtítulo "Controle de Estoque TI" em `text-[11px] text-muted-foreground`.
- Título de grupo: `text-caption text-muted-foreground px-2 py-1`.
- Item: `flex items-center gap-2.5 px-2.5 py-1.5 rounded text-body-sm transition-colors duration-micro`.
  - Ativo: `bg-surface-alt text-foreground font-semibold border-l-2 border-foreground`.
  - Inativo: `text-muted-foreground hover:text-foreground hover:bg-surface-alt`.
  - Remover a barrinha absoluta com `shadow-glow-indigo` e o `group-hover:scale-110`.
- Rodapé do usuário: `border-t border-border p-3`, nome em `text-body-sm font-medium`, papel em `text-[11px] text-muted-foreground font-mono`.

- [ ] **Step 3: Restyle da TopBar e inclusão do botão de tema**

- `<header>`: `h-14 bg-surface border-b border-border px-4 sm:px-6 flex items-center justify-between gap-3 shrink-0 z-30` (sem `backdrop-blur`, sem sombra).
- Título: `text-body-lg font-semibold text-foreground`. Remover o indicador "Sistema Conectado" (enfeite sem informação real).
- Bloco do usuário: sem pílula colorida — `text-body-sm text-foreground` para o nome e `text-[11px] text-muted-foreground font-mono uppercase` para o papel. Remover o mapa `ROLE_BADGES`.
- Acrescentar o botão de tema **antes** do botão de sair:

```tsx
import { useTheme } from '@/lib/theme'
import { Moon, Sun } from 'lucide-react'
// ...
const { theme, toggle } = useTheme()
// ...
<Button
  variant="ghost"
  size="icon"
  onClick={toggle}
  aria-label={theme === 'dark' ? 'Ativar modo claro' : 'Ativar modo escuro'}
  title={theme === 'dark' ? 'Modo claro' : 'Modo escuro'}
  className="text-muted-foreground hover:text-foreground"
>
  {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
</Button>
```

- Botão de sair: `variant="ghost" size="icon"` com `className="text-muted-foreground hover:text-destructive"`.

- [ ] **Step 4: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur|shadow-glow" src/components/layout/`
Expected: nenhuma saída.

- [ ] **Step 5: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 6: Commit**

```bash
git add frontend/src/components/layout/
git commit -m "feat(ui): shell monocromatico com alternancia de tema"
```

---

### Task 8: LoginPage

**Files:**
- Modify: `frontend/src/pages/LoginPage.tsx`
- Modify: `frontend/src/pages/LoginPage.test.tsx` (só se depender de classe antiga)

**Interfaces:**
- Consumes: tokens da Task 1; `Button`, `Input`, `Label` da Task 3.
- Produces: nenhuma mudança de API ou de fluxo de autenticação.

**Contexto:** tela dividida em duas colunas no design novo — painel institucional à esquerda (fundo `bg-primary`, texto `text-primary-foreground`, nome do sistema e uma linha de descrição) e formulário à direita sobre `bg-canvas`. Em telas menores que `lg`, só o formulário aparece. Logo estático (`/logo-revalle.png` não existe neste repositório: usar o ícone `Boxes` do `lucide-react` dentro de um quadrado `rounded-md border border-border`, como marca provisória). **Não** adicionar three.js nem `.glb`.

- [ ] **Step 1: Aplicar o restyle**

- Raiz: `min-h-screen grid lg:grid-cols-2 bg-canvas text-foreground`.
- Coluna institucional (`hidden lg:flex`): `bg-primary text-primary-foreground flex-col justify-between p-10`.
- Coluna do formulário: `flex items-center justify-center p-6`, com `<form>` dentro de `w-full max-w-sm space-y-5`.
- Campos com `Label` + `Input` já restilizados; botão de entrar `className="w-full"`.
- Mensagem de erro: `text-body-sm text-destructive` (cor preservada — é sinal de falha).
- Remover gradientes, vidro e sombras.

- [ ] **Step 2: Rodar os testes da tela**

Run: `npx vitest run src/pages/LoginPage.test.tsx`
Expected: PASS (2 testes).

- [ ] **Step 3: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/LoginPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 4: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 5: Commit**

```bash
git add frontend/src/pages/LoginPage.tsx frontend/src/pages/LoginPage.test.tsx
git commit -m "feat(ui): tela de login em duas colunas, monocromatica"
```

---

### Task 9: StockPage

**Files:**
- Modify: `frontend/src/pages/StockPage.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `DataTable` da Task 5; `StatusBadge` da Task 3; `ItemDetailsModal` da Task 6.
- Produces: nenhuma mudança de API.

- [ ] **Step 1: Aplicar a tabela de mapeamento**

Regras específicas:
- Cabeçalho da página: título `text-heading`, descrição `text-body-sm text-muted-foreground`, separados do conteúdo por `border-b border-border pb-4`.
- Cartões de indicador (se houver): `surface-panel p-4`, rótulo `text-caption text-muted-foreground`, valor `text-heading num`.
- Barra de filtros: `surface-panel p-3 flex flex-wrap gap-3`.
- Colunas numéricas e de identificador na tabela ganham `className="num"` na célula.

- [ ] **Step 2: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/StockPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 4: Commit**

```bash
git add frontend/src/pages/StockPage.tsx
git commit -m "feat(ui): pagina de estoque monocromatica"
```

---

### Task 10: ChartsPage

**Files:**
- Modify: `frontend/src/pages/ChartsPage.tsx`

**Interfaces:**
- Consumes: tokens da Task 1.
- Produces: nenhuma mudança de API.

**Contexto:** usa `recharts`, cujas cores vêm de props (`fill`, `stroke`), não de classes Tailwind — a tabela de mapeamento não alcança. Paleta dos gráficos passa a ser uma escala de cinzas; quando um gráfico tem duas séries (empréstimos × devoluções), diferenciar por **tom** (um escuro, um claro) em vez de matiz.

- [ ] **Step 1: Definir a paleta dos gráficos**

Acrescentar no topo do arquivo:

```tsx
/*
 * Paleta dos gráficos em escala de cinza, lida dos tokens CSS para acompanhar
 * o modo claro/escuro. recharts recebe cor por prop (fill/stroke), então não dá
 * para usar classe do Tailwind aqui — é preciso resolver o valor computado.
 */
function corDoToken(nome: string, alternativa: string): string {
  if (typeof window === 'undefined') return alternativa
  const valor = getComputedStyle(document.documentElement).getPropertyValue(nome).trim()
  return valor || alternativa
}

const CORES_GRAFICO = {
  serie1: () => corDoToken('--text-primary', '#111111'),
  serie2: () => corDoToken('--text-muted', '#85857E'),
  grade: () => corDoToken('--border', '#D9D9D4'),
  eixo: () => corDoToken('--text-secondary', '#5F5F5A'),
}
```

Substituir os valores fixos (`#6366f1`, `#10b981`, etc.) por chamadas dessas funções, e aplicar `stroke={CORES_GRAFICO.grade()}` no `CartesianGrid` e `stroke={CORES_GRAFICO.eixo()}` nos eixos.

- [ ] **Step 2: Aplicar a tabela de mapeamento no restante da página**

Cartões de gráfico: `surface-panel p-5`; título do gráfico `text-body-lg font-semibold`; legenda em `text-caption text-muted-foreground`.

- [ ] **Step 3: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|#6366f1|#10b981|#f59e0b|#8b5cf6" src/pages/ChartsPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 4: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 5: Commit**

```bash
git add frontend/src/pages/ChartsPage.tsx
git commit -m "feat(ui): graficos em escala de cinza acompanhando o tema"
```

---

### Task 11: Administração — UnidadesPage e UsersPage

**Files:**
- Modify: `frontend/src/pages/UnidadesPage.tsx`
- Modify: `frontend/src/pages/UsersPage.tsx`
- Modify: `frontend/src/pages/UnidadesPage.test.tsx` (só se depender de classe antiga)

**Interfaces:**
- Consumes: tokens da Task 1; `DataTable`, `Badge`, `Button`, `Input`, `Label`, `alert-dialog`.
- Produces: nenhuma mudança de API.

**Contexto:** `UnidadesPage` tem 612 linhas — formulário de criação/edição, tabela, painel de indicadores, diálogos de confirmação de renomeação e inativação. O painel de indicadores é o lugar natural para `figure-ground-panel` (decisão crítica) e `num` nos contadores.

- [ ] **Step 1: Aplicar a tabela de mapeamento nos dois arquivos**

Regras específicas:
- Formulários: `surface-panel p-6 space-y-4`.
- Painel de indicadores: `figure-ground-panel`, cada contador com rótulo `text-caption text-muted-foreground` e valor `text-heading num`.
- Aviso de renomeação: `border border-border-strong bg-surface-alt p-3 rounded text-body-sm` (sem âmbar).
- Badge "Inativa": `<Badge variant="default">Inativa</Badge>`.

- [ ] **Step 2: Rodar os testes da tela**

Run: `npx vitest run src/pages/UnidadesPage.test.tsx`
Expected: PASS (todos os casos existentes).

- [ ] **Step 3: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/UnidadesPage.tsx src/pages/UsersPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 4: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 5: Commit**

```bash
git add frontend/src/pages/UnidadesPage.tsx frontend/src/pages/UsersPage.tsx frontend/src/pages/UnidadesPage.test.tsx
git commit -m "feat(ui): telas de unidades e usuarios monocromaticas"
```

---

### Task 12: Periféricos — PeripheralsPage e LinkPeripheralPage

**Files:**
- Modify: `frontend/src/pages/PeripheralsPage.tsx`
- Modify: `frontend/src/pages/LinkPeripheralPage.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `DataTable`, `Badge`, `Button`, `SearchableSelect`.
- Produces: nenhuma mudança de API.

- [ ] **Step 1: Aplicar a tabela de mapeamento**

Regras específicas:
- `LinkPeripheralPage` tem duas colunas (vinculados / disponíveis): cada uma vira `surface-panel p-4`, com título em `text-caption text-muted-foreground`.
- Cartão de periférico: `border border-border rounded p-3 surface-interactive`, nome em `text-body-sm font-medium`, série em `text-[11px] text-muted-foreground num`.

- [ ] **Step 2: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/PeripheralsPage.tsx src/pages/LinkPeripheralPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 4: Commit**

```bash
git add frontend/src/pages/PeripheralsPage.tsx frontend/src/pages/LinkPeripheralPage.tsx
git commit -m "feat(ui): telas de perifericos monocromaticas"
```

---

### Task 13: Empréstimo — LoanPage, ReturnPage, TermsPage

**Files:**
- Modify: `frontend/src/pages/LoanPage.tsx`
- Modify: `frontend/src/pages/ReturnPage.tsx`
- Modify: `frontend/src/pages/TermsPage.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `ConfirmacaoTermo` da Task 6; `DataTable`, `Button`, `Input`, `Label`, `SearchableSelect`.
- Produces: nenhuma mudança de API. O toggle "É pessoa jurídica" do `LoanPage` permanece com o mesmo comportamento e texto.

- [ ] **Step 1: Aplicar a tabela de mapeamento**

Regras específicas:
- Formulário de empréstimo: `surface-panel p-6 space-y-4`, usando `page-container-reading` como largura (é tela de formulário, não de dados densos).
- Painéis de confirmação (hoje âmbar/azul): `figure-ground-panel` — é decisão crítica, e o alto contraste substitui a cor.
- Checkbox "É pessoa jurídica": `rounded-sm border-border-strong`.

- [ ] **Step 2: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/LoanPage.tsx src/pages/ReturnPage.tsx src/pages/TermsPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 4: Commit**

```bash
git add frontend/src/pages/LoanPage.tsx frontend/src/pages/ReturnPage.tsx frontend/src/pages/TermsPage.tsx
git commit -m "feat(ui): telas de emprestimo, devolucao e termos monocromaticas"
```

---

### Task 14: Auditoria — HistoryPage e ReportPage

**Files:**
- Modify: `frontend/src/pages/HistoryPage.tsx`
- Modify: `frontend/src/pages/ReportPage.tsx`
- Modify: `frontend/src/pages/HistoryPage.test.tsx` (só se depender de classe antiga)

**Interfaces:**
- Consumes: tokens da Task 1; `DataTable` da Task 5; `Badge` da Task 3.
- Produces: nenhuma mudança de API. O painel de confirmação de senha do estorno mantém comportamento e mensagens.

- [ ] **Step 1: Aplicar a tabela de mapeamento**

Regras específicas:
- `OperationBadge` do `HistoryPage`: todas as operações passam a `<Badge variant="default">` (neutro), **exceto** `Exclusão` e `Estorno`, que usam `variant="danger"` — são as operações destrutivas, e a cor aqui é sinal, não enfeite.
- Painel de confirmação de senha do estorno: `figure-ground-panel`, com a mensagem de erro em `text-body-sm text-destructive`.
- Colunas de data, CPF e identificador na tabela recebem `num`.

- [ ] **Step 2: Rodar os testes da tela**

Run: `npx vitest run src/pages/HistoryPage.test.tsx`
Expected: PASS (4 testes).

- [ ] **Step 3: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|rose-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/HistoryPage.tsx src/pages/ReportPage.tsx`
Expected: nenhuma saída.

- [ ] **Step 4: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 5: Commit**

```bash
git add frontend/src/pages/HistoryPage.tsx frontend/src/pages/ReportPage.tsx frontend/src/pages/HistoryPage.test.tsx
git commit -m "feat(ui): historico e relatorio monocromaticos"
```

---

### Task 15: Itens — RegisterItemPage e RemovePage

**Files:**
- Modify: `frontend/src/pages/RegisterItemPage.tsx`
- Modify: `frontend/src/pages/RemovePage.tsx`

**Interfaces:**
- Consumes: tokens da Task 1; `TypeSpecificFields` da Task 6; `FileUpload` da Task 4; `Button`, `Input`, `Label`, `SearchableSelect`.
- Produces: nenhuma mudança de API.

- [ ] **Step 1: Aplicar a tabela de mapeamento**

Regras específicas:
- Ambas usam `page-container-reading` (são formulários).
- `RemovePage` é ação irreversível: o formulário inteiro vira `figure-ground-panel`, o título recebe `text-destructive` e o botão de confirmar usa `Button variant="destructive"`. A cor aqui é deliberada — está na exceção das restrições globais.
- `RegisterItemPage`: seção de campos específicos por tipo separada por `border-t border-border pt-4`, com título em `text-caption text-muted-foreground`.

- [ ] **Step 2: Confirmar limpeza**

Run (Bash): `cd frontend && grep -rnE "indigo-|slate-|emerald-|amber-|sky-|purple-|violet-|glass-|backdrop-blur" src/pages/RegisterItemPage.tsx src/pages/RemovePage.tsx`
Expected: nenhuma saída. (`rose-` foi omitido do padrão de propósito: nenhuma classe `rose-` deve restar, mas `text-destructive` é permitido e não casa com o padrão.)

- [ ] **Step 3: Verificação completa**

Run: `npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tudo verde.

- [ ] **Step 4: Commit**

```bash
git add frontend/src/pages/RegisterItemPage.tsx frontend/src/pages/RemovePage.tsx
git commit -m "feat(ui): cadastro e remocao de itens monocromaticos"
```

---

### Task 16: Varredura final e verificação de ponta a ponta

**Files:**
- Modify: qualquer arquivo em que a varredura ainda encontrar resíduo.

**Interfaces:**
- Consumes: tudo.
- Produces: nada novo — é o portão de qualidade final.

- [ ] **Step 1: Varredura global da paleta antiga**

Run (Bash):
```bash
cd frontend && grep -rnE "indigo-|violet-|emerald-|rose-|amber-|sky-|purple-|slate-|glass-card|glass-panel|backdrop-blur|shadow-glow" src/ --include="*.tsx" --include="*.ts" --include="*.css"
```
Expected: nenhuma saída. Qualquer resto deve ser corrigido aqui, aplicando a Tabela de Mapeamento.

- [ ] **Step 2: Confirmar que nada do que devia ficar foi apagado**

Run (Bash):
```bash
cd .. && ls frontend/Dockerfile frontend/nginx.conf.template frontend/vitest.config.ts && \
  grep -q "react-router-dom" frontend/package.json && echo "react-router presente" && \
  ls frontend/src/pages/*.tsx | grep -v "\.test\.tsx" | wc -l
```
Expected: os três arquivos existem; "react-router presente"; `14` páginas.

- [ ] **Step 3: Suíte completa e build**

Run: `cd frontend && npx tsc --noEmit && npm run test:run && npx vite build`
Expected: tsc sem erros; todos os testes verdes; build limpo.

- [ ] **Step 4: Verificação visual real, com o sistema rodando**

Subir o banco descartável e a aplicação, e conferir as telas no navegador:

```bash
cd .. && docker compose -f docker-compose.test.yml up -d
cd frontend && npm run dev
```

Conferir, no mínimo: login (duas colunas), estoque (tabela densa), alternância de tema claro/escuro na TopBar, um modal de detalhes de item, e a tela de remoção (contraste de ação destrutiva). Encerrar com `docker compose -f docker-compose.test.yml down -v`.

- [ ] **Step 5: Build das imagens Docker**

Run (Bash): `cd .. && docker compose build frontend`
Expected: build conclui. Confirma que o port não quebrou o caminho de deploy.

- [ ] **Step 6: Commit final**

```bash
git add -A
git commit -m "chore(ui): varredura final do port visual monocromatico"
```

---

## Notas para quem executa

- **Nunca** rode `git merge origin/redesign-frontend`. A branch é referência de leitura.
- Para ver como a branch resolveu algo: `git show origin/redesign-frontend:frontend/src/<caminho>`.
- Se um teste existente quebrar por causa de classe CSS, atualize a asserção do teste — mas registre no próprio teste, em comentário, qual mudança visual motivou.
- Se encontrar uma cor da paleta antiga que **não** está na Tabela de Mapeamento, escolha o token mais próximo e registre a escolha no corpo do commit.
- A exceção de cor (status, erro, ação destrutiva) é deliberada e foi decidida pelo dono do projeto. Não "corrija" isso para monocromático puro.
