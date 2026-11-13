export type StatusTone =
  | 'available'
  | 'loaned'
  | 'pending'
  | 'return'
  | 'maintenance'
  | 'retired'
  | 'danger'
  | 'neutral'

/** Sem acento, sem caixa e sem espaços nas pontas — a comparação não deve depender de como o backend grafou o status. */
function normalize(s: string): string {
  return s
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase()
}

const TONE_BY_STATUS: Record<Exclude<StatusTone, 'neutral'>, string[]> = {
  available: ['Disponível', 'Ativo', 'Ativa', 'Assinado', 'OK'],
  loaned: ['Indisponível', 'Em uso', 'Em Uso', 'Enviado'],
  pending: ['Pendente', 'Agendado', 'Aberta', 'Aberto'],
  return: ['Pendente Devolução'],
  maintenance: ['Em Manutenção', 'Enviada', 'Em reparo'],
  retired: ['Baixado', 'Descartado', 'Substituido', 'Desativado', 'Cancelada', 'Cancelado', 'Desligado', 'Encerrado'],
  danger: ['Erro', 'Recusado', 'Expirado', 'Sem conserto'],
}

const LOOKUP = new Map<string, StatusTone>(
  (Object.entries(TONE_BY_STATUS) as [StatusTone, string[]][]).flatMap(([tone, names]) =>
    names.map((n) => [normalize(n), tone] as const)
  )
)

/**
 * Mapeia o status (pt-BR, como vem do backend) para o tom semântico usado nos tokens `--status-<tom>`.
 * Casamento exato (sem acento, caixa e espaços nas pontas); "Pendente …" desconhecido não herda o tom
 * de "Pendente" e fica neutro — status novo entra na tabela `TONE_BY_STATUS`.
 */
export function statusTone(status: string | null | undefined): StatusTone {
  if (!status) return 'neutral'
  return LOOKUP.get(normalize(status)) ?? 'neutral'
}
