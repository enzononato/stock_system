import { useState } from 'react'
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
import { isValidNotaFiscal, maskNotaFiscalInput } from '@/lib/utils'
import { ArrowLeft, Save } from 'lucide-react'

interface Props { mode: 'create' | 'edit' }

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

  const [tipo, setTipo] = useState(existingItem?.tipo ?? '')
  const [brand, setBrand] = useState(existingItem?.brand ?? '')
  const [model, setModel] = useState(existingItem?.model ?? '')
  const [revenda, setRevenda] = useState(existingItem?.revenda ?? '')
  const [notaFiscal, setNotaFiscal] = useState(existingItem?.nota_fiscal ?? '')
  const [fornecedor, setFornecedor] = useState(existingItem?.fornecedor ?? '')
  const [dateRegistered, setDateRegistered] = useState(() => {
    if (existingItem?.date_registered) {
      const d = new Date(existingItem.date_registered)
      return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
    }
    const now = new Date()
    return `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`
  })
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({})

  // Quando o item existente carregar, preenche os campos
  if (existingItem && !brand && existingItem.brand) {
    setTipo(existingItem.tipo ?? '')
    setBrand(existingItem.brand ?? '')
    setModel(existingItem.model ?? '')
    setRevenda(existingItem.revenda ?? '')
    setNotaFiscal(existingItem.nota_fiscal ?? '')
    setFornecedor(existingItem.fornecedor ?? '')
    const fields: Record<string, string> = {}
    const specificKeys = ['identificador','dominio','host','endereco_fisico','cpu','ram','storage',
      'sistema','licenca','anydesk','setor','ip','mac','potencia_nominal','autonomia_estimada',
      'ip_snmp','codigo_patrimonial','responsavel','local_instalacao','poe','quantidade_portas']
    specificKeys.forEach(k => {
      const val = (existingItem as unknown as Record<string, unknown>)[k]
      if (val) fields[k] = String(val)
    })
    setSpecificFields(fields)
  }

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) =>
      mode === 'create' ? createItem(data) : updateItem(Number(id), data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      toast(mode === 'create' ? 'Item cadastrado com sucesso!' : 'Item atualizado com sucesso!')
      navigate('/')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao salvar item.'
      toast(msg, 'error')
    },
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()

    // T3: o backend agora rejeita nota fiscal fora de 9 dígitos e campos de
    // MAC/IP malformados — validamos aqui antes de enviar. Nota fiscal é
    // opcional (sem *): só valida se algo foi preenchido.
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast('Nota fiscal inválida. Informe 9 dígitos.', 'error')
      return
    }
    const specificFieldsError = validateTypeSpecificFields(tipo, specificFields)
    if (specificFieldsError) {
      toast(specificFieldsError, 'error')
      return
    }

    const data: Record<string, unknown> = {
      tipo, brand, model, revenda,
      nota_fiscal: notaFiscal,
      fornecedor,
      ...specificFields,
    }
    if (mode === 'create') data.date_registered = dateRegistered
    mutation.mutate(data)
  }

  if (mode === 'edit' && loadingItem) {
    return <div className="py-12 text-center text-muted-foreground">Carregando...</div>
  }

  return (
    <div className="max-w-4xl space-y-6">
      <div className="flex items-center gap-3">
        <Button variant="ghost" size="sm" onClick={() => navigate(-1)}>
          <ArrowLeft size={16} />
        </Button>
        <h2 className="text-h1">{mode === 'create' ? 'Cadastrar Item' : 'Editar Item'}</h2>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[1fr_260px] gap-5 items-start">
      <form onSubmit={handleSubmit} className="bg-card rounded-xl border p-6 space-y-4">
        {/* Tipo */}
        <div className="flex flex-col gap-1.5">
          <Label>Tipo *</Label>
          <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
            <SelectTrigger><SelectValue placeholder="Selecione o tipo" /></SelectTrigger>
            <SelectContent>
              {equipmentTypes.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Marca *</Label>
            <Input value={brand} onChange={e => setBrand(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Modelo *</Label>
            <Input value={model} onChange={e => setModel(e.target.value)} required />
          </div>
        </div>

        <div className="flex flex-col gap-1.5">
          <Label>Revenda *</Label>
          <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
            <SelectTrigger><SelectValue placeholder="Selecione a revenda" /></SelectTrigger>
            <SelectContent>
              {revendas.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Nota Fiscal</Label>
            <Input
              value={notaFiscal}
              onChange={e => setNotaFiscal(maskNotaFiscalInput(e.target.value))}
              placeholder="9 dígitos"
              inputMode="numeric"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Fornecedor</Label>
            <Input value={fornecedor} onChange={e => setFornecedor(e.target.value)} />
          </div>
        </div>

        {mode === 'create' && (
          <div className="flex flex-col gap-1.5">
            <Label>Data de Cadastro *</Label>
            <Input
              value={dateRegistered}
              onChange={e => setDateRegistered(e.target.value)}
              placeholder="dd/mm/aaaa"
              required
            />
          </div>
        )}

        {/* Campos específicos do tipo */}
        {tipo && (
          <div className="border-t pt-4 space-y-4">
            <p className="text-sm font-medium text-muted-foreground">Informações específicas — {tipo}</p>
            <TypeSpecificFields
              tipo={tipo}
              values={specificFields}
              onChange={(k, v) => setSpecificFields(prev => ({ ...prev, [k]: v }))}
            />
          </div>
        )}

        <div className="flex justify-end pt-2">
          <Button type="submit" disabled={mutation.isPending}>
            <Save size={14} className="mr-2" />
            {mutation.isPending ? 'Salvando...' : (mode === 'create' ? 'Cadastrar' : 'Salvar Alterações')}
          </Button>
        </div>
      </form>

      {/* Painel de contexto — orienta o preenchimento e evita a área vazia ao lado do formulário */}
      <div className="surface rounded-lg p-4 space-y-3 sticky top-4">
        <h3 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">Checklist</h3>
        <ul className="space-y-2 text-xs">
          <li className={`flex items-center gap-2 ${tipo ? 'text-success' : 'text-muted-foreground'}`}>
            <span className={`h-1.5 w-1.5 rounded-full ${tipo ? 'bg-success' : 'bg-border'}`} /> Tipo selecionado
          </li>
          <li className={`flex items-center gap-2 ${brand && model ? 'text-success' : 'text-muted-foreground'}`}>
            <span className={`h-1.5 w-1.5 rounded-full ${brand && model ? 'bg-success' : 'bg-border'}`} /> Marca e modelo
          </li>
          <li className={`flex items-center gap-2 ${revenda ? 'text-success' : 'text-muted-foreground'}`}>
            <span className={`h-1.5 w-1.5 rounded-full ${revenda ? 'bg-success' : 'bg-border'}`} /> Revenda definida
          </li>
        </ul>
        <div className="pt-3 border-t border-border">
          <p className="text-[11px] text-muted-foreground leading-relaxed">
            Campos com <span className="text-foreground font-medium">*</span> são obrigatórios. Após o tipo ser selecionado, campos específicos do equipamento aparecem automaticamente no formulário.
          </p>
        </div>
      </div>
      </div>
    </div>
  )
}
