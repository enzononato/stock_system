import { describe, expect, it } from 'vitest'
import { render, screen } from '@testing-library/react'
import { ImpressaoEtiquetas } from '@/features/etiquetas/ImpressaoEtiquetas'

describe('ImpressaoEtiquetas', () => {
  it('monta a camada de impressão com @page e marca o body; ao sair, não deixa rastro', () => {
    const { unmount } = render(
      <ImpressaoEtiquetas>
        <p>conteudo da folha</p>
      </ImpressaoEtiquetas>
    )

    expect(document.body.classList.contains('imprimindo-etiquetas')).toBe(true)
    expect(document.querySelector('.etiquetas-impressao style')?.textContent).toContain('@page')
    expect(screen.getByText('conteudo da folha').closest('.etiquetas-impressao')).not.toBeNull()
    // A camada fica direto no body, fora do contêiner do React.
    expect(document.querySelector('.etiquetas-impressao')?.parentElement).toBe(document.body)

    unmount()

    expect(document.body.classList.contains('imprimindo-etiquetas')).toBe(false)
    expect(document.querySelector('.etiquetas-impressao')).toBeNull()
    expect(document.querySelector('style')?.textContent ?? '').not.toContain('@page')
  })
})
