import { describe, it, expect, beforeEach } from 'vitest'
import { renderHook, act } from '@testing-library/react'
import { usePageSize } from './usePageSize'

beforeEach(() => localStorage.clear())

describe('usePageSize', () => {
  it('usa o padrão sem valor salvo', () => {
    const { result } = renderHook(() => usePageSize('equipamentos'))
    expect(result.current[0]).toBe(25)
  })

  it('respeita o padrão informado', () => {
    const { result } = renderHook(() => usePageSize('equipamentos', 50))
    expect(result.current[0]).toBe(50)
  })

  it('valor inválido no storage volta ao padrão', () => {
    localStorage.setItem('revalle_page_size:equipamentos', '37')
    const { result } = renderHook(() => usePageSize('equipamentos'))
    expect(result.current[0]).toBe(25)
    localStorage.setItem('revalle_page_size:equipamentos', 'abc')
    const { result: r2 } = renderHook(() => usePageSize('equipamentos'))
    expect(r2.current[0]).toBe(25)
  })

  it('lê valor válido salvo', () => {
    localStorage.setItem('revalle_page_size:equipamentos', '100')
    const { result } = renderHook(() => usePageSize('equipamentos'))
    expect(result.current[0]).toBe(100)
  })

  it('persiste ao alterar', () => {
    const { result } = renderHook(() => usePageSize('equipamentos'))
    act(() => result.current[1](50))
    expect(result.current[0]).toBe(50)
    expect(localStorage.getItem('revalle_page_size:equipamentos')).toBe('50')
    const { result: r2 } = renderHook(() => usePageSize('equipamentos'))
    expect(r2.current[0]).toBe(50)
  })

  it('telas diferentes são independentes', () => {
    const { result } = renderHook(() => usePageSize('a'))
    act(() => result.current[1](100))
    const { result: r2 } = renderHook(() => usePageSize('b'))
    expect(r2.current[0]).toBe(25)
  })

  it('ignora tamanho inválido ao alterar', () => {
    const { result } = renderHook(() => usePageSize('a'))
    act(() => result.current[1](33))
    expect(result.current[0]).toBe(25)
  })
})
