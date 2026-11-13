import { useState, useEffect, type FormEvent } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { createItem, updateItem, getItem, uploadNotaFiscal, baixarNotaFiscal } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { LocalSelect } from '@/components/locais/LocalSelect'
import { TypeSpecificFields, validateTypeSpecificFields } from '@/components/equipment/TypeSpecificFields'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { TransferirItemDialog } from '@/features/ciclo-vida/TransferirItemDialog'
import {
  anexoGrandeDemais,
  formatMoedaInput,
  isDataBR,
  isoParaBR,
  maskDataBR,
  maskMoedaInput,
  parseMoedaBR,
} from '@/features/ciclo-vida/format'
import { useConstants } from '@/hooks/useConstants'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { isValidNotaFiscal, maskNotaFiscalInput } from '@/lib/utils'
import { ArrowLeft, Save, Loader2, PackagePlus, ArrowLeftRight, Download } from 'lucide-react'

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

const NF_ACCEPT = { 'application/pdf': ['.pdf'], 'image/png': ['.png'], 'image/jpeg': ['.jpg', '.jpeg'] }

function todayBr(): string {
  const now = new Date()
  return `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`
}

interface Props {
  mode: 'create' | 'edit'
}

/** Resultado do salvamento: o equipamento já está gravado mesmo quando o anexo da NF falha (`erroNf`). */
interface SalvarResultado {
  id: number
  erroNf?: string
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

  // Unidades ativas, para achar a unidade (por nome) cujos locais o LocalSelect oferece.
  const { data: unidades, isSuccess: unidadesCarregadas } = useQuery({
    queryKey: ['unidades', 'ativas'],
    queryFn: () => listUnidades(),
  })

  const [tipo, setTipo] = useState('')
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [revenda, setRevenda] = useState('')
  const [localId, setLocalId] = useState<number | null>(null)
  const [notaFiscal, setNotaFiscal] = useState('')
  const [fornecedor, setFornecedor] = useState('')
  const [dateRegistered, setDateRegistered] = useState(todayBr())
  const [valorCompra, setValorCompra] = useState('')
  const [dataCompra, setDataCompra] = useState('')
  const [garantiaAte, setGarantiaAte] = useState('')
  const [nfFile, setNfFile] = useState<File | null>(null)
  // Troca a chave do FileUpload para limpar o arquivo escolhido (ele guarda o próprio estado).
  const [nfKey, setNfKey] = useState(0)
  const [transferirAberto, setTransferirAberto] = useState(false)
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({})
  // Hidratação uma única vez: a flag impede que o formulário seja recarregado
  // do servidor (apagando edições em andamento) depois da primeira carga.
  const [hydrated, setHydrated] = useState(mode === 'create')

  // Cadastro que falhou só no anexo da NF navega para a edição do mesmo componente:
  // volta a hidratar do servidor e descarta o arquivo escolhido.
  useEffect(() => {
    setHydrated(mode === 'create')
    setNfFile(null)
    setNfKey((k) => k + 1)
  }, [mode])

  useEffect(() => {
    if (mode === 'edit' && existingItem && !hydrated) {
      setTipo(existingItem.tipo ?? '')
      setBrand(existingItem.brand ?? '')
      setModel(existingItem.model ?? '')
      setRevenda(existingItem.revenda ?? '')
      setLocalId(existingItem.local_id ?? null)
      setNotaFiscal(existingItem.nota_fiscal ?? '')
      setFornecedor(existingItem.fornecedor ?? '')
      // Valor e datas só-dia: convertidos por texto (nunca por `new Date`, que desloca o dia).
      setValorCompra(formatMoedaInput(existingItem.valor_compra))
      setDataCompra(isoParaBR(existingItem.data_compra))
      setGarantiaAte(isoParaBR(existingItem.garantia_ate))
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

  // Comparação sem diferenciar maiúsculas: item legado com a unidade grafada de outro jeito
  // não pode perder o local por causa da caixa.
  const unidadeDoItem = unidades?.find((u) => u.nome.trim().toLocaleLowerCase('pt-BR') === revenda.trim().toLocaleLowerCase('pt-BR'))
  const unidadeId = unidadeDoItem ? unidadeDoItem.id : null

  const mutation = useMutation({
    mutationFn: async ({
      data,
      arquivo,
    }: {
      data: Record<string, unknown>
      arquivo: File | null
    }): Promise<SalvarResultado> => {
      let itemId: number
      if (mode === 'create') {
        itemId = (await createItem(data)).id
      } else {
        await updateItem(Number(id), data)
        itemId = Number(id)
      }
      if (arquivo) {
        // O equipamento já está gravado: falha no anexo não pode virar um segundo cadastro.
        try {
          await uploadNotaFiscal(itemId, arquivo)
        } catch (err) {
          return { id: itemId, erroNf: getErrorMessage(err, 'Erro ao anexar a nota fiscal.') }
        }
      }
      return { id: itemId }
    },
    onSuccess: (resultado) => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      void queryClient.invalidateQueries({ queryKey: ['item'] })
      if (resultado.erroNf) {
        toast(`Equipamento salvo, mas a nota fiscal não foi anexada: ${resultado.erroNf}`, 'error')
        // Na criação, vai para a edição do item já gravado (anexar a NF de novo lá); na edição, fica na tela.
        if (mode === 'create') void navigate(`/edit/${resultado.id}`)
        return
      }
      toast(mode === 'create' ? 'Item cadastrado com sucesso!' : 'Item atualizado com sucesso!')
      // "/" é o Dashboard; quem cadastra ou edita um item quer voltar para a
      // lista de estoque, não para o painel de indicadores.
      void navigate('/stock')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao salvar item.'), 'error')
    },
  })

  function selecionarNf(file: File | null) {
    if (file && anexoGrandeDemais(file)) {
      toast('A nota fiscal deve ter no máximo 10 MB.', 'error')
      setNfFile(null)
      setNfKey((k) => k + 1)
      return
    }
    setNfFile(file)
  }

  function handleSubmit(e: FormEvent) {
    e.preventDefault()

    // O backend rejeita nota fiscal fora de 9 dígitos e campos de MAC/IP
    // malformados — validamos aqui antes de enviar. Nota fiscal é opcional
    // (sem *): só valida se algo foi preenchido.
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast('Nota fiscal inválida. Informe 9 dígitos.', 'error')
      return
    }
    const valor = valorCompra.trim() ? parseMoedaBR(valorCompra) : null
    if (valorCompra.trim() && valor === null) {
      toast('Valor de compra inválido.', 'error')
      return
    }
    if (dataCompra.trim() && !isDataBR(dataCompra)) {
      toast('Data da compra inválida. Use dd/mm/aaaa.', 'error')
      return
    }
    if (garantiaAte.trim() && !isDataBR(garantiaAte)) {
      toast('Fim da garantia inválido. Use dd/mm/aaaa.', 'error')
      return
    }
    if (nfFile && anexoGrandeDemais(nfFile)) {
      toast('A nota fiscal deve ter no máximo 10 MB.', 'error')
      return
    }
    const specificFieldsError = validateTypeSpecificFields(tipo, specificFields)
    if (specificFieldsError) {
      toast(specificFieldsError, 'error')
      return
    }

    const data: Record<string, unknown> = {
      tipo,
      brand,
      model,
      nota_fiscal: notaFiscal,
      fornecedor,
      ...specificFields,
    }
    if (mode === 'create') {
      // Na edição a unidade está travada (só muda pela transferência): não é enviada.
      data['revenda'] = revenda
      data['date_registered'] = dateRegistered
      // Na criação só vão os campos preenchidos.
      if (valor !== null) data['valor_compra'] = valor
      if (dataCompra.trim()) data['data_compra'] = dataCompra.trim()
      if (garantiaAte.trim()) data['garantia_ate'] = garantiaAte.trim()
      if (localId !== null) data['local_id'] = localId
    } else {
      // Na edição, `null` limpa o campo no backend (chave ausente = não muda).
      data['valor_compra'] = valor
      data['data_compra'] = dataCompra.trim() || null
      data['garantia_ate'] = garantiaAte.trim() || null
      data['local_id'] = localId
    }
    mutation.mutate({ data, arquivo: nfFile })
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
            {mode === 'edit' && existingItem?.status && <StatusPill status={existingItem.status} />}
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
              <Label htmlFor="tipo" className="text-caption font-medium">Tipo *</Label>
              <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                <SelectTrigger id="tipo" className="h-9">
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
            <p className="text-caption text-muted-foreground">Revenda responsável, local, nota fiscal e fornecedor.</p>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="revenda" className="text-caption font-medium">Revenda de Origem *</Label>
              <div className="flex items-center gap-2">
                <Select
                  value={revenda}
                  // Na edição a unidade é travada: o Radix chama onValueChange('') quando o valor
                  // gravado não está na lista (ex.: grafia antiga), e isso apagaria a unidade do estado.
                  onValueChange={(v) => {
                    if (mode === 'create') setRevenda(v)
                  }}
                  required={mode === 'create'}
                  disabled={constantsLoading || mode === 'edit'}
                >
                  <SelectTrigger id="revenda" className="h-9">
                    <SelectValue placeholder="Selecione a revenda" />
                  </SelectTrigger>
                  <SelectContent>
                    {(revenda && !revendas.some((r) => r === revenda) ? [revenda, ...revendas] : revendas).map((r) => (
                      <SelectItem key={r} value={r}>
                        {upper(r)}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {mode === 'edit' && existingItem && (
                  <Button
                    type="button"
                    variant="outline"
                    size="sm"
                    onClick={() => setTransferirAberto(true)}
                  >
                    <ArrowLeftRight className="mr-1.5 size-3.5" />
                    Transferir
                  </Button>
                )}
              </div>
              {mode === 'edit' && (
                <p className="text-caption text-muted-foreground">
                  Para mudar a unidade, use Transferir: a mudança fica registrada na linha do tempo.
                </p>
              )}
            </div>
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Local</Label>
              {unidadesCarregadas ? (
                <LocalSelect unidadeId={unidadeId} value={localId} onChange={setLocalId} />
              ) : (
                <LocalSelect unidadeId={null} value={null} onChange={() => {}} disabled />
              )}
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

        {/* Bloco 3: Compra e garantia */}
        <div className="space-y-4 border-t border-border pt-4">
          <div className="border-b border-border pb-2">
            <h2 className="text-body font-semibold text-foreground">3. Compra e garantia</h2>
            <p className="text-caption text-muted-foreground">
              Valor, data da compra, fim da garantia e a nota fiscal digitalizada (opcionais).
            </p>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="valor_compra" className="text-caption font-medium">Valor de compra (R$)</Label>
              <Input
                id="valor_compra"
                value={valorCompra}
                onChange={(e) => setValorCompra(maskMoedaInput(e.target.value))}
                placeholder="0,00"
                inputMode="decimal"
                className="font-mono h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="data_compra" className="text-caption font-medium">Data da compra</Label>
              <Input
                id="data_compra"
                value={dataCompra}
                onChange={(e) => setDataCompra(maskDataBR(e.target.value))}
                placeholder="DD/MM/AAAA"
                inputMode="numeric"
                className="font-mono h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="garantia_ate" className="text-caption font-medium">Fim da garantia</Label>
              <Input
                id="garantia_ate"
                value={garantiaAte}
                onChange={(e) => setGarantiaAte(maskDataBR(e.target.value))}
                placeholder="DD/MM/AAAA"
                inputMode="numeric"
                className="font-mono h-9"
              />
            </div>
          </div>

          <div className="space-y-2">
            {mode === 'edit' && existingItem?.nf_anexo_key && (
              <div className="flex flex-wrap items-center gap-3 text-body-sm text-foreground">
                <span>Nota fiscal anexada</span>
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() =>
                    void baixarNotaFiscal(existingItem).catch((err: unknown) =>
                      toast(getErrorMessage(err, 'Não foi possível baixar a nota fiscal.'), 'error')
                    )
                  }
                >
                  <Download className="mr-1.5 size-3.5" />
                  Baixar nota fiscal
                </Button>
              </div>
            )}
            <FileUpload
              key={nfKey}
              accept={NF_ACCEPT}
              onFile={selecionarNf}
              label="Arraste ou clique para anexar a nota fiscal (PDF ou imagem, até 10 MB)"
            />
          </div>
        </div>

        {/* Bloco 4: Especificações Técnicas Específicas */}
        {tipo && (
          <div className="space-y-4 border-t border-border pt-4">
            <div className="border-b border-border pb-2">
              <h2 className="text-body font-semibold text-foreground">
                4. Informações Específicas — {tipo}
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

          <Button type="submit" loading={mutation.isPending} size="sm">
            {mode === 'create' ? <PackagePlus className="mr-1.5 size-3.5" /> : <Save className="mr-1.5 size-3.5" />}
            {mode === 'create' ? 'Cadastrar Equipamento' : 'Salvar Alterações'}
          </Button>
        </div>
      </form>

      {mode === 'edit' && existingItem && (
        <TransferirItemDialog
          item={{ id: existingItem.id, revenda: revenda, local_id: localId }}
          open={transferirAberto}
          onOpenChange={setTransferirAberto}
          onTransferido={({ para_unidade, para_local_id }) => {
            setRevenda(para_unidade)
            setLocalId(para_local_id)
          }}
        />
      )}
    </div>
  )
}
