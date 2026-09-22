import { describe, it, expect, vi } from 'vitest'
import { render } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { Checkbox } from './checkbox'

describe('Checkbox', () => {
  it('aciona onCheckedChange ao clicar no quadrado (label visível), mesmo sem a prop id', async () => {
    // Regressão: quando `id` não é passado, o `htmlFor` do label interno
    // (o quadrado visível) fica `undefined` e o clique nele não ativa o
    // input associado — só o clique no texto de um <label> externo que o
    // envolvesse funcionaria, e mesmo esse fica suprimido por aninhamento
    // de <label>. Este teste clica direto no quadrado.
    const onCheckedChange = vi.fn()
    const { container } = render(<Checkbox onCheckedChange={onCheckedChange} />)

    const square = container.querySelector('label')
    expect(square).toBeTruthy()

    await userEvent.click(square as HTMLLabelElement)

    expect(onCheckedChange).toHaveBeenCalledTimes(1)
    expect(onCheckedChange).toHaveBeenCalledWith(true)
  })
})
