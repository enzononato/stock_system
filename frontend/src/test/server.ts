import { setupServer } from 'msw/node'
import { handlers } from './handlers'

/**
 * Servidor MSW compartilhado por toda a suíte (ver `src/test/setup.ts`, que
 * liga/reseta/desliga este servidor entre os testes). Nenhum teste de
 * frontend deve abrir uma conexão real de rede — toda chamada à API passa por
 * aqui, mesmo as que "deveriam" falhar (401, 403, 500 etc.).
 */
export const server = setupServer(...handlers)
