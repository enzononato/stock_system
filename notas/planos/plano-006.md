# Work Plan 006 — Direção Visual Avançada (Spotify + Solvd + Elemento 3D)

## Date
2026-09-02

## Objective
Implementar as diretrizes visuais do Addendum ao `PROMPT_MASTER_ANTIGRAVITY.md`:
1. **Identidade Visual e Motion:** Integrar a paleta Azul Marinho Corporativo profunda com design tokens refinados e motion design inspirado na Solvd (transições suaves, elevações táteis, microinterações intencionais).
2. **Navegação e Experiência Unificada:** Elevar o Shell da aplicação (`AppShell.tsx`) aos padrões de produto único e contínuo inspirados no Spotify Web Player.
3. **Elemento 3D Otimizado na Tela de Login:**
   - Versionar exclusivamente `frontend/public/3d/logo.glb` (37,4 KB, compressão Draco, 21.328 faces, gerado a partir do `3dsvg.stl` original de 21,3 MB).
   - Servir os binários do Draco decoder de forma 100% self-hosted em `frontend/public/draco/` sem dependência de CDN.
   - Criar renderizador WebGL com Three.js (`ThreeLogoCanvas.tsx`) isolado via `React.lazy()` com controle de iluminação, materiais PBR acetinados e parallax suave com o mouse.
   - Criar fallback vetorial de alta performance (`LogoFallback.tsx`) para mobile (< 1024px), falha de GPU ou `prefers-reduced-motion`.
   - Redesenhar a `LoginPage.tsx` combinando área visual cinematográfica e card de acesso seguro perfeitamente integrado à autenticação real JWT.
4. **Tabelas e Responsividade:** Garantir contenção de scroll horizontal seguro em `DataTable.tsx` mantendo a padronização de 7 itens por página.
5. **Validação:** Compilar com `npm run build` (Exit Code 0), inspecionar chunks do bundle e validar comportamento no navegador.
6. **Auditoria e Registro:** Criar `notas/auditoria/auditoria-006.md` e atualizar `estado.md`.

## Scope
FRONTEND ONLY. Preservação integral do backend FastAPI, banco MySQL, autenticação JWT, contratos de API e dados reais.

## Status
Em execução.
