import { useState } from 'react'
import type { ApagamentoInput } from '@/api/items'
import { FileUpload } from '@/components/ui/FileUpload'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { anexoGrandeDemais, hojeBR, isDataBR, maskDataBR } from './format'

export interface ApagamentoState {
  dadosApagados: '' | 'sim' | 'nao_aplica'
  metodo: string
  apagadoPor: string
  apagadoEm: string
  comprovante: File | null
}

export function apagamentoInicial(): ApagamentoState {
  return { dadosApagados: '', metodo: '', apagadoPor: '', apagadoEm: hojeBR(), comprovante: null }
}

/** Todos os dados exigidos pelo backend estão preenchidos (a data só vale se existir no calendário). */
export function apagamentoCompleto(s: ApagamentoState, permiteNaoAplica: boolean): boolean {
  if (s.dadosApagados === 'nao_aplica') return permiteNaoAplica
  if (s.dadosApagados === 'sim') {
    return s.metodo !== '' && s.apagadoPor.trim() !== '' && isDataBR(s.apagadoEm)
  }
  return false
}

export function paraApagamentoInput(s: ApagamentoState): ApagamentoInput {
  if (s.dadosApagados === 'nao_aplica') return { dados_apagados: 'nao_aplica' }
  return {
    dados_apagados: 'sim',
    metodo_apagamento: s.metodo,
    apagado_por: s.apagadoPor.trim(),
    apagado_em: s.apagadoEm,
    comprovante: s.comprovante,
  }
}

const ACCEPT = { 'application/pdf': ['.pdf'], 'image/png': ['.png'], 'image/jpeg': ['.jpg', '.jpeg'] }

const radioClass = 'size-4 accent-[var(--ring)]'

interface Props {
  tipo: string
  value: ApagamentoState
  onChange: (s: ApagamentoState) => void
  /** Chave → rótulo (de `/api/constants`). */
  metodos: Record<string, string>
  /** Só equipamento fora de posse (roubo ou perda) pode registrar "não se aplica". */
  permiteNaoAplica: boolean
}

/**
 * Registro de apagamento de dados (LGPD) da baixa: obrigatório para os tipos com
 * armazenamento. O comprovante é opcional (PDF ou imagem, até 10 MB).
 */
export function ApagamentoDadosFields({ tipo, value, onChange, metodos, permiteNaoAplica }: Props) {
  // O FileUpload guarda o próprio arquivo; trocar a chave o esvazia quando o arquivo é descartado.
  const [uploadKey, setUploadKey] = useState(0)
  const set = (parcial: Partial<ApagamentoState>) => onChange({ ...value, ...parcial })

  function selecionarComprovante(file: File | null) {
    if (file && anexoGrandeDemais(file)) {
      toast('O comprovante deve ter no máximo 10 MB.', 'error')
      setUploadKey((k) => k + 1)
      set({ comprovante: null })
      return
    }
    set({ comprovante: file })
  }

  return (
    <fieldset className="space-y-4 rounded border border-border bg-surface-alt p-4">
      <legend className="px-1 text-caption font-semibold text-foreground">
        {`Apagamento de dados (LGPD) — obrigatório para ${tipo}`}
      </legend>

      <div className="flex flex-col gap-2">
        <label className="flex items-center gap-2 text-body-sm text-foreground">
          <input
            type="radio"
            name="dados_apagados"
            className={radioClass}
            checked={value.dadosApagados === 'sim'}
            onChange={() => set({ dadosApagados: 'sim' })}
          />
          Sim, os dados foram apagados
        </label>
        {permiteNaoAplica && (
          <label className="flex items-center gap-2 text-body-sm text-foreground">
            <input
              type="radio"
              name="dados_apagados"
              className={radioClass}
              checked={value.dadosApagados === 'nao_aplica'}
              onChange={() => set({ dadosApagados: 'nao_aplica' })}
            />
            Não se aplica — equipamento fora de posse (roubo ou perda)
          </label>
        )}
      </div>

      {value.dadosApagados === 'sim' && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="apagamento-metodo" className="text-caption font-medium">
                Método de apagamento
              </Label>
              <select
                id="apagamento-metodo"
                aria-label="Método de apagamento"
                className="flex h-9 w-full rounded border border-input bg-surface px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring"
                value={value.metodo}
                onChange={(e) => set({ metodo: e.target.value })}
              >
                <option value="">Selecione o método</option>
                {Object.entries(metodos).map(([chave, rotulo]) => (
                  <option key={chave} value={chave}>
                    {rotulo}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="apagamento-por" className="text-caption font-medium">
                Quem apagou os dados
              </Label>
              <Input
                id="apagamento-por"
                value={value.apagadoPor}
                maxLength={100}
                onChange={(e) => set({ apagadoPor: e.target.value })}
                className="h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="apagamento-em" className="text-caption font-medium">
                Data do apagamento
              </Label>
              <Input
                id="apagamento-em"
                value={value.apagadoEm}
                onChange={(e) => set({ apagadoEm: maskDataBR(e.target.value) })}
                placeholder="DD/MM/AAAA"
                inputMode="numeric"
                className="h-9 font-mono"
              />
            </div>
          </div>

          <FileUpload
            key={uploadKey}
            accept={ACCEPT}
            onFile={selecionarComprovante}
            label="Comprovante do apagamento (opcional)"
          />
        </div>
      )}
    </fieldset>
  )
}
