import { describe, expect, it } from 'vitest'
import {
  cn,
  formatCpf,
  formatDate,
  formatDateTime,
  isValidCpf,
  isValidIp,
  isValidMac,
  isValidNotaFiscal,
  maskCpfInput,
  maskMacInput,
  maskNotaFiscalInput,
} from '@/lib/utils'

describe('cn', () => {
  it('resolve o alias @ e mescla classes condicionais (smoke test da infraestrutura)', () => {
    expect(cn('a', false && 'b', 'c')).toBe('a c')
  })
})

describe('isValidCpf', () => {
  it.each([
    ['111.444.777-35', true, 'CPF válido real, com máscara'],
    ['11144477735', true, 'CPF válido real, sem máscara'],
    ['529.982.247-25', true, 'segundo CPF válido real, com máscara'],
    ['52998224725', true, 'segundo CPF válido real, sem máscara'],
  ])('%s -> %s (%s)', (cpf, esperado) => {
    expect(isValidCpf(cpf)).toBe(esperado)
  })

  it('rejeita dígito verificador errado (último dígito alterado de um CPF válido)', () => {
    // 111.444.777-35 é válido; trocar o último dígito quebra o segundo DV.
    expect(isValidCpf('111.444.777-36')).toBe(false)
  })

  it('rejeita dígito verificador errado (primeiro DV alterado)', () => {
    expect(isValidCpf('111.444.777-45')).toBe(false)
  })

  it('rejeita sequência de dígitos repetidos (111.111.111-11) — regressão: a versão antiga aceitava', () => {
    expect(isValidCpf('111.111.111-11')).toBe(false)
  })

  it('rejeita todas as sequências repetidas de 0 a 9, com e sem máscara', () => {
    for (let d = 0; d <= 9; d++) {
      const digits = String(d).repeat(11)
      expect(isValidCpf(digits)).toBe(false)
      expect(isValidCpf(`${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9)}`)).toBe(false)
    }
  })

  it('rejeita string vazia', () => {
    expect(isValidCpf('')).toBe(false)
  })

  it('rejeita null e undefined', () => {
    expect(isValidCpf(null)).toBe(false)
    expect(isValidCpf(undefined)).toBe(false)
  })

  it('rejeita CPF com menos de 11 dígitos', () => {
    expect(isValidCpf('1114447773')).toBe(false)
  })

  it('rejeita CPF com mais de 11 dígitos', () => {
    expect(isValidCpf('111444777356')).toBe(false)
  })

  it('rejeita entrada só com caracteres não numéricos', () => {
    expect(isValidCpf('abc.def.ghi-jk')).toBe(false)
  })
})

describe('isValidNotaFiscal', () => {
  it('aceita exatamente 9 dígitos', () => {
    expect(isValidNotaFiscal('123456789')).toBe(true)
  })

  it('rejeita 8 dígitos', () => {
    expect(isValidNotaFiscal('12345678')).toBe(false)
  })

  it('rejeita 10 dígitos', () => {
    expect(isValidNotaFiscal('1234567890')).toBe(false)
  })

  it('rejeita string vazia, null e undefined', () => {
    expect(isValidNotaFiscal('')).toBe(false)
    expect(isValidNotaFiscal(null)).toBe(false)
    expect(isValidNotaFiscal(undefined)).toBe(false)
  })

  it('ignora caracteres não numéricos ao contar os dígitos', () => {
    expect(isValidNotaFiscal('123.456.789')).toBe(true)
    expect(isValidNotaFiscal('12.345.678')).toBe(false)
  })
})

describe('isValidMac', () => {
  it('aceita o formato com dois-pontos (AA:BB:CC:DD:EE:FF)', () => {
    expect(isValidMac('AA:BB:CC:DD:EE:FF')).toBe(true)
  })

  it('aceita o formato com hífen (AA-BB-CC-DD-EE-FF)', () => {
    expect(isValidMac('AA-BB-CC-DD-EE-FF')).toBe(true)
  })

  it('aceita o formato sem separador (AABBCCDDEEFF)', () => {
    expect(isValidMac('AABBCCDDEEFF')).toBe(true)
  })

  it('aceita minúsculas em qualquer um dos três formatos', () => {
    expect(isValidMac('aa:bb:cc:dd:ee:ff')).toBe(true)
    expect(isValidMac('aa-bb-cc-dd-ee-ff')).toBe(true)
    expect(isValidMac('aabbccddeeff')).toBe(true)
  })

  it('rejeita separadores mistos', () => {
    expect(isValidMac('AA:BB-CC:DD-EE:FF')).toBe(false)
    expect(isValidMac('AA-BB:CC-DD:EE-FF')).toBe(false)
  })

  it('rejeita caracteres fora do intervalo hexadecimal', () => {
    expect(isValidMac('GG:BB:CC:DD:EE:FF')).toBe(false)
  })

  it('rejeita quantidade errada de octetos', () => {
    expect(isValidMac('AA:BB:CC:DD:EE')).toBe(false)
    expect(isValidMac('AA:BB:CC:DD:EE:FF:00')).toBe(false)
  })

  it('rejeita string vazia, null e undefined', () => {
    expect(isValidMac('')).toBe(false)
    expect(isValidMac(null)).toBe(false)
    expect(isValidMac(undefined)).toBe(false)
  })
})

describe('isValidIp', () => {
  it('aceita um IPv4 válido', () => {
    expect(isValidIp('192.168.0.1')).toBe(true)
    expect(isValidIp('0.0.0.0')).toBe(true)
    expect(isValidIp('255.255.255.255')).toBe(true)
  })

  it('rejeita octeto acima de 255', () => {
    expect(isValidIp('192.168.0.256')).toBe(false)
    expect(isValidIp('999.1.1.1')).toBe(false)
  })

  it('rejeita zeros à esquerda em um octeto', () => {
    expect(isValidIp('192.168.00.1')).toBe(false)
    expect(isValidIp('192.168.0.01')).toBe(false)
  })

  it('rejeita número errado de octetos', () => {
    expect(isValidIp('192.168.0')).toBe(false)
    expect(isValidIp('192.168.0.1.5')).toBe(false)
  })

  it('rejeita octeto não numérico', () => {
    expect(isValidIp('192.168.a.1')).toBe(false)
  })

  it('rejeita string vazia, null e undefined', () => {
    expect(isValidIp('')).toBe(false)
    expect(isValidIp(null)).toBe(false)
    expect(isValidIp(undefined)).toBe(false)
  })
})

describe('maskCpfInput', () => {
  it('aplica a máscara progressivamente conforme o usuário digita', () => {
    expect(maskCpfInput('1')).toBe('1')
    expect(maskCpfInput('111')).toBe('111')
    expect(maskCpfInput('1114')).toBe('111.4')
    expect(maskCpfInput('111444')).toBe('111.444')
    expect(maskCpfInput('1114447')).toBe('111.444.7')
    expect(maskCpfInput('111444777')).toBe('111.444.777')
    expect(maskCpfInput('1114447773')).toBe('111.444.777-3')
    expect(maskCpfInput('11144477735')).toBe('111.444.777-35')
  })

  it('corta a digitação em 11 dígitos, ignorando o excedente', () => {
    expect(maskCpfInput('111444777356789')).toBe('111.444.777-35')
  })

  it('remove caracteres não numéricos da entrada', () => {
    expect(maskCpfInput('111.444.777-35')).toBe('111.444.777-35')
    expect(maskCpfInput('abc111444777xyz35')).toBe('111.444.777-35')
  })

  it('entrada vazia devolve string vazia', () => {
    expect(maskCpfInput('')).toBe('')
  })
})

describe('maskNotaFiscalInput', () => {
  it('mantém apenas dígitos', () => {
    expect(maskNotaFiscalInput('123abc456')).toBe('123456')
  })

  it('corta no máximo de 9 dígitos', () => {
    expect(maskNotaFiscalInput('123456789999')).toBe('123456789')
  })

  it('entrada vazia devolve string vazia', () => {
    expect(maskNotaFiscalInput('')).toBe('')
  })
})

describe('maskMacInput', () => {
  it('aplica a máscara progressivamente conforme o usuário digita', () => {
    expect(maskMacInput('A')).toBe('A')
    expect(maskMacInput('AA')).toBe('AA')
    expect(maskMacInput('AAB')).toBe('AA:B')
    expect(maskMacInput('AABB')).toBe('AA:BB')
    expect(maskMacInput('AABBCCDDEEFF')).toBe('AA:BB:CC:DD:EE:FF')
  })

  it('corta no máximo de 12 dígitos hexadecimais', () => {
    expect(maskMacInput('AABBCCDDEEFF0011')).toBe('AA:BB:CC:DD:EE:FF')
  })

  it('remove caracteres fora do intervalo hexadecimal e mistura maiúsculas/minúsculas', () => {
    expect(maskMacInput('aa:bb-cc.dd/ee\\ff')).toBe('AA:BB:CC:DD:EE:FF')
    expect(maskMacInput('zzAAbbZZ')).toBe('AA:BB')
  })

  it('entrada vazia devolve string vazia', () => {
    expect(maskMacInput('')).toBe('')
  })
})

describe('formatDate', () => {
  it('entrada nula/indefinida devolve "-"', () => {
    expect(formatDate(null)).toBe('-')
    expect(formatDate(undefined)).toBe('-')
    expect(formatDate('')).toBe('-')
  })

  it('entrada inválida (não parseável como data) devolve a string original', () => {
    expect(formatDate('não é uma data')).toBe('não é uma data')
  })

  it('entrada válida formata como data pt-BR', () => {
    expect(formatDate('2026-01-15')).toBe(new Date('2026-01-15').toLocaleDateString('pt-BR'))
  })
})

describe('formatDateTime', () => {
  it('entrada nula/indefinida devolve "-"', () => {
    expect(formatDateTime(null)).toBe('-')
    expect(formatDateTime(undefined)).toBe('-')
    expect(formatDateTime('')).toBe('-')
  })

  it('entrada inválida (não parseável como data) devolve a string original', () => {
    expect(formatDateTime('não é uma data')).toBe('não é uma data')
  })

  it('entrada válida formata como data+hora pt-BR', () => {
    const iso = '2026-01-15T10:30:00'
    expect(formatDateTime(iso)).toBe(new Date(iso).toLocaleString('pt-BR'))
  })
})

describe('formatCpf', () => {
  it('entrada nula/indefinida devolve "-"', () => {
    expect(formatCpf(null)).toBe('-')
    expect(formatCpf(undefined)).toBe('-')
    expect(formatCpf('')).toBe('-')
  })

  it('11 dígitos são pontuados, mesmo sem dígito verificador válido (formatCpf não valida)', () => {
    expect(formatCpf('11144477735')).toBe('111.444.777-35')
    expect(formatCpf('00000000000')).toBe('000.000.000-00')
  })

  it('quantidade de dígitos diferente de 11 devolve a entrada original, sem pontuar', () => {
    expect(formatCpf('123456789')).toBe('123456789')
    expect(formatCpf('123456789012')).toBe('123456789012')
  })
})
