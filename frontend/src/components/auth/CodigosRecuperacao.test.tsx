import { describe, expect, it, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { CodigosRecuperacao } from '@/components/auth/CodigosRecuperacao'

const CODIGOS = Array.from({ length: 10 }, (_, i) => `abcd-efgh-ijkl-mn${String(i).padStart(2, '2')}`)

describe('CodigosRecuperacao', () => {
  it('mostra os 10 códigos e só libera continuar depois de marcar a confirmação', async () => {
    const onContinuar = vi.fn()
    const user = userEvent.setup()
    render(<CodigosRecuperacao codigos={CODIGOS} onContinuar={onContinuar} />)

    CODIGOS.forEach((c) => expect(screen.getByText(c)).toBeInTheDocument())
    expect(
      screen.getByText(/Guarde estes 10 códigos em local seguro/)
    ).toBeInTheDocument()

    const continuar = screen.getByRole('button', { name: 'Continuar' })
    expect(continuar).toBeDisabled()

    await user.click(screen.getByLabelText('Guardei os códigos em local seguro'))
    expect(continuar).toBeEnabled()
    await user.click(continuar)
    expect(onContinuar).toHaveBeenCalledTimes(1)
  })

  it('usa o rótulo do botão de continuar informado', () => {
    render(<CodigosRecuperacao codigos={CODIGOS} onContinuar={() => {}} rotuloContinuar="Entrar no sistema" />)
    expect(screen.getByRole('button', { name: 'Entrar no sistema' })).toBeDisabled()
  })

  it('copia os códigos, um por linha', async () => {
    const user = userEvent.setup()
    render(<CodigosRecuperacao codigos={CODIGOS} onContinuar={() => {}} />)

    await user.click(screen.getByRole('button', { name: 'Copiar códigos' }))

    expect(await navigator.clipboard.readText()).toBe(CODIGOS.join('\n'))
  })
})

describe('CodigosRecuperacao — download', () => {
  it('baixa o .txt e só revoga a URL do Blob depois, num timeout', async () => {
    const criar = vi.fn(() => 'blob:codigos')
    const revogar = vi.fn()
    vi.stubGlobal('URL', Object.assign(URL, { createObjectURL: criar, revokeObjectURL: revogar }))
    const clique = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    try {
      const user = userEvent.setup()
      render(<CodigosRecuperacao codigos={CODIGOS} onContinuar={() => {}} />)

      await user.click(screen.getByRole('button', { name: 'Baixar .txt' }))

      expect(criar).toHaveBeenCalledTimes(1)
      expect(clique).toHaveBeenCalledTimes(1)
      expect(revogar).not.toHaveBeenCalled()
      await vi.waitFor(() => expect(revogar).toHaveBeenCalledWith('blob:codigos'), { timeout: 3000 })
    } finally {
      clique.mockRestore()
      vi.unstubAllGlobals()
    }
  })
})
