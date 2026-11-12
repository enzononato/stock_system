import { useCallback, useState } from 'react'

const VALID = [25, 50, 100] as const
type PageSize = (typeof VALID)[number]

function isValid(n: number): n is PageSize {
  return (VALID as readonly number[]).includes(n)
}

function read(key: string, padrao: PageSize): PageSize {
  try {
    const raw = localStorage.getItem(key)
    if (raw === null) return padrao
    const n = Number(raw)
    return isValid(n) ? n : padrao
  } catch {
    return padrao
  }
}

/** Tamanho de página por tela, lembrado entre sessões. Só aceita 25, 50 ou 100. */
export function usePageSize(tela: string, padrao: 25 | 50 | 100 = 25): [number, (n: number) => void] {
  const key = `revalle_page_size:${tela}`
  const [size, setSize] = useState<number>(() => read(key, padrao))

  const update = useCallback(
    (n: number) => {
      if (!isValid(n)) return
      setSize(n)
      try {
        localStorage.setItem(key, String(n))
      } catch {
        /* armazenamento indisponível: o valor vale só nesta sessão */
      }
    },
    [key]
  )

  return [size, update]
}
