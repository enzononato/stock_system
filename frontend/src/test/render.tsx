import type { ReactElement, ReactNode } from 'react'
import { render, renderHook, type RenderOptions } from '@testing-library/react'
import { QueryClient, QueryClientProvider } from '@tanstack/react-query'

/**
 * QueryClient "de teste": sem retry e sem cache entre montagens, para que uma
 * requisição que devolve erro não fique tentando de novo em loop e cada teste
 * comece do zero (mesmo reaproveitando o mesmo componente/hook).
 */
export function createTestQueryClient() {
  return new QueryClient({
    defaultOptions: {
      queries: { retry: false, gcTime: 0 },
      mutations: { retry: false },
    },
  })
}

function Wrapper({ children, client }: { children: ReactNode; client: QueryClient }) {
  return <QueryClientProvider client={client}>{children}</QueryClientProvider>
}

/** `render` com `QueryClientProvider` já configurado — use quando o componente usa `useQuery`/`useMutation`. */
export function renderWithClient(ui: ReactElement, options?: RenderOptions) {
  const client = createTestQueryClient()
  return {
    client,
    ...render(ui, { wrapper: ({ children }) => <Wrapper client={client}>{children}</Wrapper>, ...options }),
  }
}

/** `renderHook` com `QueryClientProvider` já configurado — use para hooks como `useConstants`. */
export function renderHookWithClient<TResult, TProps>(
  hook: (props: TProps) => TResult,
  client: QueryClient = createTestQueryClient()
) {
  return {
    client,
    ...renderHook(hook, { wrapper: ({ children }) => <Wrapper client={client}>{children}</Wrapper> }),
  }
}
