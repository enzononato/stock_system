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
