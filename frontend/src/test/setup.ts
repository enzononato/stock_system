import '@testing-library/jest-dom/vitest'
import { afterAll, afterEach, beforeAll } from 'vitest'
import { cleanup } from '@testing-library/react'
import { server } from './server'

// Nenhuma requisição não-mockada deve escapar para a rede de verdade: um
// handler ausente é um erro de teste (rota que o teste esqueceu de simular),
// não uma chamada real.
beforeAll(() => server.listen({ onUnhandledRequest: 'error' }))

afterEach(() => {
  cleanup()
  server.resetHandlers()
})

afterAll(() => server.close())
