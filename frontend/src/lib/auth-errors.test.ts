import { describe, expect, it } from 'vitest'
import { AxiosError, type AxiosResponse, type InternalAxiosRequestConfig } from 'axios'
import { MSG_MUITAS_TENTATIVAS, mensagemErroAuth } from '@/lib/auth-errors'

function erroHttp(status: number, data: unknown) {
  const response = {
    status,
    statusText: '',
    data,
    headers: {},
    config: {} as InternalAxiosRequestConfig,
  } as AxiosResponse
  return new AxiosError('falhou', 'ERR_BAD_REQUEST', undefined, undefined, response)
}

describe('mensagemErroAuth', () => {
  it('429 vira a mensagem de muitas tentativas (o corpo do slowapi não tem detail)', () => {
    const erro = erroHttp(429, { error: 'Rate limit exceeded: 5 per 1 minute' })
    expect(mensagemErroAuth(erro, 'Falhou.')).toBe(
      'Muitas tentativas. Aguarde um minuto e tente de novo.'
    )
    expect(MSG_MUITAS_TENTATIVAS).toBe('Muitas tentativas. Aguarde um minuto e tente de novo.')
  })

  it('outros status usam o detail do servidor', () => {
    expect(mensagemErroAuth(erroHttp(400, { detail: 'Código inválido.' }), 'Falhou.')).toBe(
      'Código inválido.'
    )
  })

  it('sem detail cai no fallback informado', () => {
    expect(mensagemErroAuth({}, 'Usuário ou senha inválidos.')).toBe('Usuário ou senha inválidos.')
  })
})
