import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'

interface FieldConfig {
  key: string
  label: string
  type?: 'text' | 'select'
  options?: string[]
  placeholder?: string
}

const TYPE_FIELDS: Record<string, FieldConfig[]> = {
  Celular: [
    { key: 'identificador', label: 'IMEI / Nº de Série' },
    { key: 'storage', label: 'Armazenamento (GB)' },
  ],
  Notebook: [
    { key: 'dominio', label: 'Domínio', type: 'select', options: ['Sim', 'Não'] },
    { key: 'host', label: 'Host' },
    { key: 'endereco_fisico', label: 'Endereço Físico (MAC)' },
    { key: 'cpu', label: 'Processador (CPU)' },
    { key: 'ram', label: 'Memória RAM' },
    { key: 'storage', label: 'Armazenamento (GB)' },
    { key: 'sistema', label: 'Sistema Operacional' },
    { key: 'licenca', label: 'Licença Windows' },
    { key: 'anydesk', label: 'AnyDesk' },
  ],
  Desktop: [
    { key: 'dominio', label: 'Domínio', type: 'select', options: ['Sim', 'Não'] },
    { key: 'host', label: 'Host' },
    { key: 'endereco_fisico', label: 'Endereço Físico (MAC)' },
    { key: 'cpu', label: 'Processador (CPU)' },
    { key: 'ram', label: 'Memória RAM' },
    { key: 'storage', label: 'Armazenamento (GB)' },
    { key: 'sistema', label: 'Sistema Operacional' },
    { key: 'licenca', label: 'Licença Windows' },
    { key: 'anydesk', label: 'AnyDesk' },
  ],
  Impressora: [
    { key: 'setor', label: 'Setor' },
    { key: 'ip', label: 'IP' },
    { key: 'mac', label: 'MAC' },
  ],
  Tablet: [
    { key: 'identificador', label: 'Nº de Série' },
    { key: 'storage', label: 'Armazenamento (GB)' },
  ],
  Switch: [
    { key: 'poe', label: 'POE', type: 'select', options: ['Sim', 'Não'] },
    { key: 'quantidade_portas', label: 'Qtd. de Portas' },
  ],
  HD: [{ key: 'storage', label: 'Armazenamento (GB)' }],
  Nobreak: [
    { key: 'identificador', label: 'Nº de Série' },
    { key: 'codigo_patrimonial', label: 'Código Patrimonial' },
    { key: 'responsavel', label: 'Responsável' },
    { key: 'potencia_nominal', label: 'Potência Nominal (VA/W)' },
    { key: 'autonomia_estimada', label: 'Autonomia Estimada (min)' },
    { key: 'ip_snmp', label: 'IP Placa SNMP' },
  ],
  'Access Point': [
    { key: 'identificador', label: 'Nº de Série' },
    { key: 'codigo_patrimonial', label: 'Código Patrimonial' },
    { key: 'local_instalacao', label: 'Local de Instalação' },
    { key: 'setor', label: 'Setor' },
    { key: 'ip', label: 'IP' },
    { key: 'mac', label: 'MAC' },
  ],
}

interface TypeSpecificFieldsProps {
  tipo: string
  values: Record<string, string>
  onChange: (key: string, value: string) => void
  disabled?: boolean
}

export function TypeSpecificFields({ tipo, values, onChange, disabled }: TypeSpecificFieldsProps) {
  const fields = TYPE_FIELDS[tipo] ?? []
  if (fields.length === 0) return null

  return (
    <>
      {fields.map((field) => (
        <div key={field.key} className="flex flex-col gap-1.5">
          <Label htmlFor={field.key}>{field.label}</Label>
          {field.type === 'select' ? (
            <Select
              value={values[field.key] ?? ''}
              onValueChange={(v) => onChange(field.key, v)}
              disabled={disabled}
            >
              <SelectTrigger id={field.key}>
                <SelectValue placeholder={`Selecione ${field.label}`} />
              </SelectTrigger>
              <SelectContent>
                {field.options!.map((opt) => (
                  <SelectItem key={opt} value={opt}>{opt}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          ) : (
            <Input
              id={field.key}
              value={values[field.key] ?? ''}
              onChange={(e) => onChange(field.key, e.target.value)}
              placeholder={field.placeholder ?? field.label}
              disabled={disabled}
            />
          )}
        </div>
      ))}
    </>
  )
}
