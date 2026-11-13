import { describe, expect, it, vi } from 'vitest'
import { AxiosError, type AxiosResponse, type InternalAxiosRequestConfig } from 'axios'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { CadastroAutenticador } from '@/components/auth/CadastroAutenticador'

const SEGREDO = 'JBSWY3DPEHPK3PXPJBSWY3DPEHPK3PXP'
const CADASTRO = { segredo: SEGREDO, otpauth_uri: `otpauth://totp/Revalle:x?secret=${SEGREDO}` }

function erroHttp(status: number, detail: string) {
  const response = {
    status,
    statusText: '',
    data: { detail },
    headers: {},
    config: {} as InternalAxiosRequestConfig,
  } as AxiosResponse
  return new AxiosError('falhou', 'ERR_BAD_REQUEST', undefined, undefined, response)
}

describe('CadastroAutenticador', () => {
  it('mostra o QR code e a chave em grupos de 4', async () => {
    render(
      <CadastroAutenticador
        iniciar={() => Promise.resolve(CADASTRO)}
        confirmar={() => Promise.resolve([])}
        onConcluido={() => {}}
      />
    )

    expect(screen.getByText('Gerando QR code...')).toBeInTheDocument()
    expect(await screen.findByAltText('QR code para o aplicativo autenticador')).toBeInTheDocument()
    expect(screen.getByText('JBSW Y3DP EHPK 3PXP JBSW Y3DP EHPK 3PXP')).toBeInTheDocument()
    expect(
      screen.getByText('Não consegue ler o QR code? Digite esta chave no aplicativo:')
    ).toBeInTheDocument()
  })

  it('confirma o código e entrega os códigos de recuperação', async () => {
    const codigos = ['aaaa-bbbb-cccc-dddd']
    const confirmar = vi.fn().mockResolvedValue(codigos)
    const onConcluido = vi.fn()
    const user = userEvent.setup()
    render(
      <CadastroAutenticador
        iniciar={() => Promise.resolve(CADASTRO)}
        confirmar={confirmar}
        onConcluido={onConcluido}
      />
    )

    await user.type(await screen.findByLabelText('Código de 6 dígitos'), '123456')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    await vi.waitFor(() => expect(onConcluido).toHaveBeenCalledWith(codigos))
    expect(confirmar).toHaveBeenCalledWith('123456')
  })

  it('código recusado mostra a mensagem e não conclui', async () => {
    const onConcluido = vi.fn()
    const user = userEvent.setup()
    render(
      <CadastroAutenticador
        iniciar={() => Promise.resolve(CADASTRO)}
        confirmar={() => Promise.reject(erroHttp(400, 'Código inválido.'))}
        onConcluido={onConcluido}
      />
    )

    await user.type(await screen.findByLabelText('Código de 6 dígitos'), '000000')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    expect(await screen.findByText('Código inválido.')).toBeInTheDocument()
    expect(onConcluido).not.toHaveBeenCalled()
  })

  it('falha ao iniciar mostra o detail do servidor e nenhum QR', async () => {
    const detail = 'A verificação em duas etapas não está disponível: o servidor está sem a chave.'
    render(
      <CadastroAutenticador
        iniciar={() => Promise.reject(erroHttp(503, detail))}
        confirmar={() => Promise.resolve([])}
        onConcluido={() => {}}
      />
    )

    expect(await screen.findByRole('alert')).toHaveTextContent(detail)
    expect(screen.queryByAltText('QR code para o aplicativo autenticador')).not.toBeInTheDocument()
  })
})
