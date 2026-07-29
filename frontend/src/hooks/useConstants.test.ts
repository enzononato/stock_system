import { describe, expect, it } from 'vitest'
import { http, HttpResponse, delay } from 'msw'
import { waitFor } from '@testing-library/react'
import { server } from '@/test/server'
import { createTestQueryClient, renderHookWithClient } from '@/test/render'
import { useConstants } from '@/hooks/useConstants'

describe('useConstants', () => {
  it('devolve arrays vazios enquanto carrega, e preenchidos depois que a query resolve', async () => {
    server.use(
      http.get('/api/constants', async () => {
        await delay(30)
        return HttpResponse.json({
          center_costs: ['101 - Puxada'],
          revendas: ['Revalle Juazeiro'],
          setores: ['TI'],
          equipment_types: ['Notebook'],
          peripheral_types: ['Mouse'],
          removal_reasons: ['Descarte'],
          removal_reasons_attachment: { Descarte: true },
        })
      })
    )

    const { result } = renderHookWithClient(() => useConstants())

    // Enquanto carrega: arrays vazios, nunca undefined.map(...)
    expect(result.current.isLoading).toBe(true)
    expect(result.current.centerCosts).toEqual([])
    expect(result.current.revendas).toEqual([])
    expect(result.current.setores).toEqual([])
    expect(result.current.equipmentTypes).toEqual([])
    expect(result.current.peripheralTypes).toEqual([])
    expect(result.current.removalReasons).toEqual([])
    expect(result.current.removalReasonsAttachment).toEqual({})

    await waitFor(() => expect(result.current.isLoading).toBe(false))

    expect(result.current.centerCosts).toEqual(['101 - Puxada'])
    expect(result.current.revendas).toEqual(['Revalle Juazeiro'])
    expect(result.current.setores).toEqual(['TI'])
    expect(result.current.equipmentTypes).toEqual(['Notebook'])
    expect(result.current.peripheralTypes).toEqual(['Mouse'])
    expect(result.current.removalReasons).toEqual(['Descarte'])
    expect(result.current.removalReasonsAttachment).toEqual({ Descarte: true })
  })

  it('compartilha o cache entre chamadas: a segunda montagem não dispara uma nova requisição', async () => {
    let chamadasAoServidor = 0
    server.use(
      http.get('/api/constants', () => {
        chamadasAoServidor += 1
        return HttpResponse.json({
          center_costs: [],
          revendas: ['Revalle Juazeiro'],
          setores: [],
          equipment_types: [],
          peripheral_types: [],
          removal_reasons: [],
          removal_reasons_attachment: {},
        })
      })
    )

    // Mesmo QueryClient para as duas montagens: é o que faz o cache (queryKey
    // compartilhado) valer a pena entre StockPage/LoanPage/RegisterItemPage/etc.
    const client = createTestQueryClient()

    const primeira = renderHookWithClient(() => useConstants(), client)
    await waitFor(() => expect(primeira.result.current.isLoading).toBe(false))
    expect(chamadasAoServidor).toBe(1)

    const segunda = renderHookWithClient(() => useConstants(), client)
    // Já nasce com os dados do cache, sem precisar esperar isLoading virar false.
    expect(segunda.result.current.isLoading).toBe(false)
    expect(segunda.result.current.revendas).toEqual(['Revalle Juazeiro'])
    expect(chamadasAoServidor).toBe(1)
  })
})
