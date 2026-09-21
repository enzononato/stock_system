import { useEffect, useState } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { createItem, updateItem, getItem } from '@/api/items'
import { TypeSpecificFields, validateTypeSpecificFields } from '@/components/equipment/TypeSpecificFields'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { useConstants } from '@/hooks/useConstants'
import { isValidNotaFiscal, maskNotaFiscalInput } from '@/lib/utils'
import { ArrowLeft, Save, Loader2 } from 'lucide-react'

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
  // Hidratação uma-única-vez: a guarda antiga (`!brand`) recarregava o
  // formulário inteiro do servidor sempre que o usuário limpasse o campo
  // Marca, apagando qualquer edição em andamento. Com a flag, a hidratação
  // roda só na primeira vez que `existingItem` chega.
  const [hydrated, setHydrated] = useState(false)

  useEffect(() => {
    if (existingItem && !hydrated) {
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
      setHydrated(true)
    }
  }, [existingItem, hydrated])

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) =>
      mode === 'create' ? createItem(data) : updateItem(Number(id), data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      toast(mode === 'create' ? 'Item cadastrado com sucesso!' : 'Item atualizado com sucesso!')
      // Antes da Task 7 deste plano, "/" era a lista de estoque; agora é o
      // Dashboard. Quem cadastra ou edita um item quer voltar para a lista,
      // não para o painel de indicadores — por isso o destino é "/stock".
      navigate('/stock')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao salvar item.'), 'error')
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
    return (
      <div className="py-12 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" />
        Carregando...
      </div>
    )
  }

  return (
    <div className="page-container-reading space-y-6">
      <Button variant="ghost" size="sm" onClick={() => navigate(-1)}>
        <ArrowLeft size={16} />
      </Button>

      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail={mode === 'create' ? 'Novo Cadastro' : 'Edição de Item'}
        title={mode === 'create' ? 'Cadastrar Item' : 'Editar Item'}
        description={
          mode === 'create'
            ? 'Cadastre um novo equipamento no estoque.'
            : 'Atualize os dados do equipamento selecionado.'
        }
      />

      <form onSubmit={handleSubmit} className="surface-panel p-6 space-y-6">
        {/* Seção 1: Identificação Geral */}
        <div className="space-y-4">
          <PanelHeader
            title="1. Identificação Geral"
            description="Tipo de ativo, fabricante e modelo comercial."
          />
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
        </div>

        {/* Seção 2: Lotação e Dados de Origem */}
        <div className="space-y-4 border-t border-border pt-4">
          <PanelHeader
            title="2. Lotação e Dados de Origem"
            description="Revenda responsável, nota fiscal e fornecedor."
          />
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
        </div>

        {/* Seção 3: Campos específicos do tipo */}
        {tipo && (
          <div className="space-y-4 border-t border-border pt-4">
            <PanelHeader
              title={`3. Campos específicos — ${tipo}`}
              description={`Parâmetros técnicos aplicáveis a ativos do tipo ${tipo}.`}
            />
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <TypeSpecificFields
                tipo={tipo}
                values={specificFields}
                onChange={(k, v) => setSpecificFields(prev => ({ ...prev, [k]: v }))}
              />
            </div>
          </div>
        )}

        <div className="flex justify-end pt-2">
          <Button type="submit" disabled={mutation.isPending}>
            <Save size={14} />
            {mutation.isPending ? 'Salvando...' : (mode === 'create' ? 'Cadastrar' : 'Salvar Alterações')}
          </Button>
        </div>
      </form>
    </div>
  )
}
