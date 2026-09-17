import { useState, useEffect, type FormEvent } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { createItem, updateItem, getItem } from '@/api/items'
import { TypeSpecificFields, validateTypeSpecificFields } from '@/components/equipment/TypeSpecificFields'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { toast } from '@/components/ui/toast'
import { useConstants } from '@/hooks/useConstants'
import { getErrorMessage } from '@/lib/api-error'
import { isValidNotaFiscal, maskNotaFiscalInput } from '@/lib/utils'
import { ArrowLeft, Save, Loader2, PackagePlus } from 'lucide-react'

const SPECIFIC_KEYS = [
  'identificador',
  'dominio',
  'host',
  'endereco_fisico',
  'cpu',
  'ram',
  'storage',
  'sistema',
  'licenca',
  'anydesk',
  'setor',
  'ip',
  'mac',
  'potencia_nominal',
  'autonomia_estimada',
  'ip_snmp',
  'codigo_patrimonial',
  'responsavel',
  'local_instalacao',
  'poe',
  'quantidade_portas',
] as const

function todayBr(): string {
  const now = new Date()
  return `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`
}

interface Props {
  mode: 'create' | 'edit'
}

export default function RegisterItemPage({ mode }: Props) {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const { equipmentTypes, revendas, isLoading: constantsLoading } = useConstants()

  const { data: existingItem, isLoading: loadingItem } = useQuery({
    queryKey: ['item', id],
    queryFn: () => getItem(Number(id)),
    enabled: mode === 'edit' && !!id,
  })

  const [tipo, setTipo] = useState('')
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [revenda, setRevenda] = useState('')
  const [notaFiscal, setNotaFiscal] = useState('')
  const [fornecedor, setFornecedor] = useState('')
  const [dateRegistered, setDateRegistered] = useState(todayBr())
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({})
  const [hydrated, setHydrated] = useState(mode === 'create')

  useEffect(() => {
    if (mode === 'edit' && existingItem && !hydrated) {
      setTipo(existingItem.tipo ?? '')
      setBrand(existingItem.brand ?? '')
      setModel(existingItem.model ?? '')
      setRevenda(existingItem.revenda ?? '')
      setNotaFiscal(existingItem.nota_fiscal ?? '')
      setFornecedor(existingItem.fornecedor ?? '')
      if (existingItem.date_registered) {
        const d = new Date(existingItem.date_registered)
        if (!Number.isNaN(d.getTime())) {
          setDateRegistered(
            `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`
          )
        }
      }
      const fields: Record<string, string> = {}
      const record = existingItem as unknown as Record<string, unknown>
      for (const k of SPECIFIC_KEYS) {
        const val = record[k]
        if (val) fields[k] = String(val)
      }
      setSpecificFields(fields)
      setHydrated(true)
    }
  }, [mode, existingItem, hydrated])

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) =>
      mode === 'create' ? createItem(data) : updateItem(Number(id), data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      toast.success(mode === 'create' ? 'Item cadastrado com sucesso!' : 'Item atualizado com sucesso!')
      void navigate('/stock')
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao salvar item.'))
    },
  })

  function handleSubmit(e: FormEvent) {
    e.preventDefault()

    // O backend rejeita nota fiscal fora de 9 dígitos e campos de MAC/IP malformados
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error('Nota fiscal inválida. Informe 9 dígitos.')
      return
    }
    const specificFieldsError = validateTypeSpecificFields(tipo, specificFields)
    if (specificFieldsError) {
      toast.error(specificFieldsError)
      return
    }

    const data: Record<string, unknown> = {
      tipo,
      brand,
      model,
      revenda,
      nota_fiscal: notaFiscal,
      fornecedor,
      ...specificFields,
    }
    if (mode === 'create') data['date_registered'] = dateRegistered
    mutation.mutate(data)
  }

  if (mode === 'edit' && loadingItem) {
    return (
      <div className="py-20 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
        <Loader2 className="h-5 w-5 animate-spin" />
        Carregando dados do equipamento...
      </div>
    )
  }

  return (
    <div className="mx-auto max-w-4xl space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Inventário
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">
              {mode === 'create' ? 'Novo Cadastro' : 'Edição de Item'}
            </span>
          </div>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            {mode === 'create' ? 'Cadastrar Equipamento' : 'Editar Equipamento'}
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            {mode === 'create'
              ? 'Registre um novo ativo de TI no estoque patrimonial.'
              : 'Atualize as especificações e vínculos do equipamento selecionado.'}
          </p>
        </div>

        <Button
          variant="outline"
          size="sm"
          onClick={() => void navigate('/stock')}
          className="self-start"
        >
          <ArrowLeft className="mr-1.5 size-3.5" />
          Voltar ao Estoque
        </Button>
      </div>

      <form
        onSubmit={handleSubmit}
        className="rounded border border-border bg-surface p-6 space-y-6"
      >
        {/* Bloco 1: Informações Gerais */}
        <div className="space-y-4">
          <div className="border-b border-border pb-2">
            <h2 className="text-body font-semibold text-foreground">1. Identificação Geral</h2>
            <p className="text-caption text-muted-foreground">Tipo de ativo, fabricante e modelo comercial.</p>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Tipo *</Label>
              <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                <SelectTrigger className="h-9">
                  <SelectValue placeholder="Selecione o tipo" />
                </SelectTrigger>
                <SelectContent>
                  {equipmentTypes.map((t) => (
                    <SelectItem key={t} value={t}>
                      {t}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="brand" className="text-caption font-medium">Marca *</Label>
              <Input
                id="brand"
                value={brand}
                onChange={(e) => setBrand(e.target.value)}
                placeholder="Ex: Dell, Lenovo, HP"
                required
                className="h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="model" className="text-caption font-medium">Modelo *</Label>
              <Input
                id="model"
                value={model}
                onChange={(e) => setModel(e.target.value)}
                placeholder="Ex: Latitude 3420"
                required
                className="h-9"
              />
            </div>
          </div>
        </div>

        {/* Bloco 2: Lotação & Origem */}
        <div className="space-y-4 pt-2">
          <div className="border-b border-border pb-2">
            <h2 className="text-body font-semibold text-foreground">2. Lotação e Dados de Origem</h2>
            <p className="text-caption text-muted-foreground">Revenda responsável, nota fiscal e fornecedor.</p>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Revenda de Origem *</Label>
              <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
                <SelectTrigger className="h-9">
                  <SelectValue placeholder="Selecione a revenda" />
                </SelectTrigger>
                <SelectContent>
                  {revendas.map((r) => (
                    <SelectItem key={r} value={r}>
                      {r}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="nota_fiscal" className="text-caption font-medium">Nota Fiscal</Label>
              <Input
                id="nota_fiscal"
                value={notaFiscal}
                onChange={(e) => setNotaFiscal(maskNotaFiscalInput(e.target.value))}
                placeholder="000000000 (9 dígitos)"
                inputMode="numeric"
                className="font-mono h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="fornecedor" className="text-caption font-medium">Fornecedor</Label>
              <Input
                id="fornecedor"
                value={fornecedor}
                onChange={(e) => setFornecedor(e.target.value)}
                placeholder="Nome da fornecedora"
                className="h-9"
              />
            </div>
          </div>

          {mode === 'create' && (
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-3 pt-2">
              <div className="flex flex-col gap-1.5">
                <Label htmlFor="date_registered" className="text-caption font-medium">Data de Cadastro *</Label>
                <Input
                  id="date_registered"
                  value={dateRegistered}
                  onChange={(e) => setDateRegistered(e.target.value)}
                  placeholder="DD/MM/AAAA"
                  required
                  className="font-mono h-9"
                />
              </div>
            </div>
          )}
        </div>

        {/* Bloco 3: Especificações Técnicas Específicas */}
        {tipo && (
          <div className="space-y-4 border-t border-border pt-4">
            <div className="border-b border-border pb-2">
              <h2 className="text-body font-semibold text-foreground">
                3. Informações Específicas — {tipo}
              </h2>
              <p className="text-caption text-muted-foreground">
                Parâmetros técnicos específicos para ativos do tipo {tipo}.
              </p>
            </div>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
              <TypeSpecificFields
                tipo={tipo}
                values={specificFields}
                onChange={(k, v) => setSpecificFields((prev) => ({ ...prev, [k]: v }))}
              />
            </div>
          </div>
        )}

        <div className="flex items-center justify-between border-t border-border pt-4">
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={() => void navigate('/stock')}
          >
            Cancelar
          </Button>

          <Button type="submit" disabled={mutation.isPending} size="sm">
            {mode === 'create' ? <PackagePlus className="mr-1.5 size-3.5" /> : <Save className="mr-1.5 size-3.5" />}
            {mutation.isPending
              ? 'Salvando…'
              : mode === 'create'
                ? 'Cadastrar Equipamento'
                : 'Salvar Alterações'}
          </Button>
        </div>
      </form>
    </div>
  )
}
