import { useEffect, useState, type ReactNode } from 'react'
import { createPortal } from 'react-dom'

/**
 * Regras de impressão da folha de etiquetas. Ficam DENTRO do portal, não no CSS
 * global: assim o `@page { margin: 0 }` só existe enquanto a tela de etiquetas
 * está montada e não deixa em branco a impressão das outras telas.
 */
export const CSS_IMPRESSAO = `
@page { size: A4 portrait; margin: 0; }
@media screen { .etiquetas-impressao { display: none; } }
@media print {
  body.imprimindo-etiquetas { margin: 0; background: white; }
  body.imprimindo-etiquetas > *:not(.etiquetas-impressao) { display: none !important; }
  .etiquetas-impressao { display: block; }
  .etiquetas-impressao, .etiquetas-impressao * { color: black !important; background: white !important; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .etiquetas-impressao .folha { break-after: page; page-break-after: always; }
  .etiquetas-impressao .folha:last-child { break-after: auto; page-break-after: auto; }
}
`

/**
 * Camada de impressão: um `div.etiquetas-impressao` direto no `<body>` (fora do
 * contêiner do app, que o CSS de impressão esconde) e a classe
 * `imprimindo-etiquetas` no body — os dois só existem enquanto o componente está montado.
 */
export function ImpressaoEtiquetas({ children }: { children: ReactNode }) {
  const [camada] = useState(() => {
    const div = document.createElement('div')
    div.className = 'etiquetas-impressao'
    return div
  })

  useEffect(() => {
    document.body.appendChild(camada)
    document.body.classList.add('imprimindo-etiquetas')
    return () => {
      camada.remove()
      document.body.classList.remove('imprimindo-etiquetas')
    }
  }, [camada])

  return createPortal(
    <>
      <style>{CSS_IMPRESSAO}</style>
      {children}
    </>,
    camada
  )
}
