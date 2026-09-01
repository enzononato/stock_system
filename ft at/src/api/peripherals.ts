import api from './client'

export interface Peripheral {
  id: number
  tipo: string
  brand?: string
  model?: string
  identificador?: string
  status?: string
  motivo_substituicao?: string
  date_registered?: string
  link_id?: number
}

export async function listPeripherals(params?: {
  status?: string
  tipo?: string
  include_inactive?: boolean
}) {
  const res = await api.get('/peripherals', { params })
  return res.data as Peripheral[]
}

export async function listItemPeripherals(itemId: number) {
  const res = await api.get(`/items/${itemId}/peripherals`)
  return res.data as Peripheral[]
}

export async function createPeripheral(data: { tipo: string; brand?: string; model?: string; identificador?: string }) {
  const res = await api.post('/peripherals', data)
  return res.data
}

export async function linkPeripheral(itemId: number, peripheralId: number) {
  const res = await api.post(`/items/${itemId}/peripherals/${peripheralId}`)
  return res.data
}

export async function unlinkPeripheral(linkId: number) {
  const res = await api.delete(`/peripherals/links/${linkId}`)
  return res.data
}

export async function deletePeripheral(peripheralId: number): Promise<{ detail: string }> {
  const res = await api.delete(`/peripherals/${peripheralId}`)
  return res.data as { detail: string }
}

export async function replacePeripheral(
  itemId: number,
  oldPeripheralId: number,
  newPeripheralId: number,
  reason: string,
  attachment?: File
) {
  const form = new FormData()
  form.append('new_peripheral_id', String(newPeripheralId))
  form.append('reason', reason)
  if (attachment) form.append('attachment', attachment)
  const res = await api.post(`/items/${itemId}/peripherals/${oldPeripheralId}/replace`, form)
  return res.data
}
