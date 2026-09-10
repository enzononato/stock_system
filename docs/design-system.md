# DESIGN SYSTEM — ENTERPRISE EDITORIAL v5
## Stock System / Revalle — Gestão de Patrimônio & IT Asset Management

Documentação oficial do Design System monocromático corporativo implementado no produto real.

---

## 1. Princípios Centrais
- **Informação > Decoração**
- **Clareza > Efeitos**
- **Operação > Apresentação**
- **Composição > Componentes**
- **Identidade > Tendência**
- **Design System compartilhado ≠ Layout compartilhado** (cada rota possui composição e protagonista visual próprios).

---

## 2. Paleta Monocromática

A identidade visual é puramente construída por contraste, tipografia e densidade de dados (sem cores de sotaque como azul, navy, roxo, verde, etc.).

### Light Mode (Editorial Técnico)
| Token | Hex | Papel |
|---|---|---|
| `--canvas` | `#F7F7F5` | Fundo geral da aplicação |
| `--surface` | `#FFFFFF` | Superfícies principais (tabelas, painéis) |
| `--surface-alt` | `#F0F0ED` | Hover de linhas, cabeçalhos secundários |
| `--border` | `#D9D9D4` | Bordas hairline 1px |
| `--border-strong` | `#BDBDB7` | Bordas de foco e separação prioritária |
| `--text-primary` | `#111111` | Títulos, valores técnicos e ênfase alta |
| `--text-secondary` | `#5F5F5A` | Labels, metadados e textos de apoio |
| `--text-muted` | `#85857E` | Placeholders e indicações desabilitadas |

### Dark Mode (Preto Puro Corporativo)
| Token | Hex | Papel |
|---|---|---|
| `--canvas` | `#050505` | Fundo geral da aplicação |
| `--surface` | `#0C0C0C` | Superfícies principais |
| `--surface-alt` | `#131313` | Hover de linhas, cabeçalhos secundários |
| `--border` | `#272727` | Bordas hairline 1px |
| `--border-strong` | `#3A3A3A` | Bordas de foco e separação prioritária |
| `--text-primary` | `#F2F2F0` | Títulos, valores técnicos e ênfase alta |
| `--text-secondary` | `#A4A4A0` | Labels, metadados e textos de apoio |
| `--text-muted` | `#70706C` | Placeholders e indicações desabilitadas |

---

## 3. Botões & Ações
- **Light Primário:** `background: #111111; color: #FFFFFF;`
- **Dark Primário:** `background: #F2F2F0; color: #080808;`
- **Secundário:** Fundo transparente ou `surface-alt`, borda hairline monocromática, texto de alto contraste.
- **Destrutivo:** Contraste reforçado, texto explícito, sem depender exclusivamente de cor vermelha para comunicar consequência.

---

## 4. Tipografia
- **Interface Geral:** `Plus Jakarta Sans`
- **Dados Técnicos (Patrimônio, Serial, MAC, IDs, Datas, Números):** `IBM Plex Mono` (tabular nums)
- **Escala de Tamanhos:**
  - `caption`: 12px / 500 (uppercase, tracking 0.05em)
  - `body-sm`: 13px / 400
  - `body`: 14px / 400
  - `body-lg`: 16px / 500
  - `heading-sm`: 18px / 600
  - `heading`: 22px / 600
  - `heading-lg`: 28px / 600 (teto executivo: máx 32px)

---

## 5. Forma, Radius & Sombra
- **Bordas:** Hairline 1px em todo o produto.
- **Radius Estritos:**
  - `2px`: Badges e inputs pequenos
  - `4px`: Botões, inputs normais, linhas de tabela
  - `6px`: Painéis principais
  - `8px`: Máximo absoluto (avatares com justificativa)
  - ❌ Proibido: `rounded-xl`, `rounded-2xl`, `rounded-3xl`, `rounded-full`, pills.
- **Sombras:** Apenas para overlays flutuantes (Dialog, DropdownMenu, Toast). Zero sombra decorativa em cards estáticos.

---

## 6. Motion Tokens (Seção 44)
- **Microinterações:** 120–200ms
- **Overlays & Drawers:** 200–300ms
- **Curva de Easing Padrão:** `cubic-bezier(0.16, 1, 0.3, 1)`
- **Transição de Tema:** Suave, 200ms entre modo claro e escuro.

---

## 7. Sinalização de Status Monocromática (Seção 15)
A interface é 100% compreensível em preto e branco:
- `● ATIVO` (preenchido com texto de alto contraste)
- `○ INATIVO / DISPONÍVEL` (borda fina, vazado)
- `! ATENÇÃO` (símbolo de exclamação e contraste de superfície)
- `× ERRO / BLOQUEADO` (símbolo explícito e aviso descritivo)
- `✓ CONCLUÍDO` (símbolo de checagem e preenchimento sólido)
