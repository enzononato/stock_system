import { http, HttpResponse } from 'msw'

/**
 * Handlers "de fábrica" para a suíte de frontend: cobrem só o suficiente para
 * que hooks/páginas que dependem de `GET /api/constants`, `/api/items` e
 * `/api/history` não quebrem quando um teste não se importa com o conteúdo
 * dessas respostas. Qualquer teste que precise de um corpo/status específico
 * sobrescreve com `server.use(...)` (ver `src/test/server.ts`) — nunca edite
 * estes defaults para casos de um único teste.
 */
export const handlers = [
  http.get('/api/constants', () =>
    HttpResponse.json({
      center_costs: [],
      revendas: [],
      setores: [],
      equipment_types: [],
      peripheral_types: [],
      removal_reasons: [],
      removal_reasons_attachment: {},
    })
  ),

  http.get('/api/items', () => HttpResponse.json({ items: [], total: 0 })),

  http.get('/api/history', () => HttpResponse.json({ items: [], total: 0 })),
]
