import { defineConfig, mergeConfig } from 'vitest/config'
import viteConfig from './vite.config'

// Reaproveita a config real do Vite (plugins, e principalmente o alias `@` que
// aponta para `./src`) em vez de duplicá-la aqui: várias importações em
// `src/**` dependem desse alias, e se ele mudar em `vite.config.ts` os testes
// devem acompanhar automaticamente, sem exigir uma segunda edição.
export default mergeConfig(
  viteConfig,
  defineConfig({
    test: {
      environment: 'jsdom',
      setupFiles: ['./src/test/setup.ts'],
      css: false,
      globals: false,
      restoreMocks: true,
      coverage: {
        provider: 'v8',
        reporter: ['text', 'html'],
        include: ['src/**/*.{ts,tsx}'],
        exclude: [
          'src/main.tsx',
          'src/test/**',
          'src/**/*.test.{ts,tsx}',
          'src/**/*.d.ts',
        ],
      },
    },
  })
)
