import api from './client'

export async function initiateLoan(data: {
  item_id: number
  usuario: string
  cpf: string
  center_cost: string
  cargo: string
  setor: string
  revenda: string
  date_issue: string
}) {
  const res = await api.post('/loans', data)
  return res.data
}

export async function confirmLoan(itemId: number, signedPdf: File) {
  const form = new FormData()
  form.append('signed_pdf', signedPdf)
  const res = await api.post(`/loans/${itemId}/confirm`, form)
  return res.data
}

export async function initiateReturn(itemId: number) {
  const res = await api.post(`/loans/${itemId}/return/initiate`)
  return res.data as { detail: string; download_url: string; filename: string }
}

export async function confirmReturn(itemId: number, signedPdf: File) {
  const form = new FormData()
  form.append('signed_pdf', signedPdf)
  const res = await api.post(`/loans/${itemId}/return/confirm`, form)
  return res.data
}

export async function downloadSignedTerm(itemId: number): Promise<Blob> {
  const urlRes = await api.get(`/loans/${itemId}/signed-term-url`)
  const fileUrl = (urlRes.data as { url: string }).url
  // fileUrl comes as /api/documents/files/... — strip /api prefix to avoid double baseURL
  const relativePath = fileUrl.replace(/^\/api/, '')
  const fileRes = await api.get(relativePath, { responseType: 'blob' })
  return fileRes.data as Blob
}

export async function generateLoanTerm(itemId: number): Promise<Blob> {
  const res = await api.post(`/documents/loan-term/${itemId}`, {}, { responseType: 'blob' })
  return res.data as Blob
}
