import { useState } from 'react'
import { Link } from 'react-router-dom'
import type { AlertasPainel, TipoAlerta } from '@/api/painel'
import { Button } from '@/components/ui/button'
import { StatusPill } from '@/components/ui/StatusPill'
import {
  ROTULO_SEVERIDADE,
  ROTULO_TIPO_ALERTA,
  TOM_SEVERIDADE,
  agruparAlertas,
} from './alertas'

const POR_GRUPO = 5

/** Alertas calculados na hora pelo servidor, agrupados por tipo; cada um leva ao registro. */
export function ListaAlertas({ dados }: { dados: AlertasPainel }) {
  const [expandidos, setExpandidos] = useState<Partial<Record<TipoAlerta, boolean>>>({})
  const grupos = agruparAlertas(dados.alertas)
  const { limites } = dados

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between border-b border-border pb-3">
        <h2 className="text-body font-semibold text-foreground">Alertas</h2>
        <span className="rounded-sm border border-border px-2 py-0.5 font-mono text-caption font-bold text-foreground">
          {dados.total}
        </span>
      </div>

      {grupos.length === 0 ? (
        <p className="py-6 text-center text-body-sm text-muted-foreground">Nenhum alerta no momento.</p>
      ) : (
        <div className="space-y-5">
          {grupos.map(({ tipo, alertas }) => {
            const expandido = expandidos[tipo] === true
            const visiveis = expandido ? alertas : alertas.slice(0, POR_GRUPO)
            return (
              <section key={tipo} aria-label={ROTULO_TIPO_ALERTA[tipo]} className="space-y-1">
                <h3 className="flex items-baseline gap-2 text-caption font-semibold text-muted-foreground">
                  <span>{ROTULO_TIPO_ALERTA[tipo]}</span>
                  <span className="font-mono">{alertas.length}</span>
                </h3>
                <ul className="divide-y divide-border">
                  {visiveis.map((a) => (
                    <li key={a.chave}>
                      <Link
                        to={a.link}
                        className="flex items-start justify-between gap-3 rounded-sm py-2.5 hover:bg-surface-alt focus-visible:ring-1 focus-visible:ring-ring"
                      >
                        <span className="min-w-0 space-y-0.5">
                          <span className="block text-sm font-bold text-foreground">{a.titulo}</span>
                          <span className="block truncate text-xs text-muted-foreground">{a.descricao}</span>
                        </span>
                        <StatusPill
                          status={ROTULO_SEVERIDADE[a.severidade]}
                          tone={TOM_SEVERIDADE[a.severidade]}
                          className="shrink-0"
                        />
                      </Link>
                    </li>
                  ))}
                </ul>
                {alertas.length > POR_GRUPO && (
                  <Button
                    type="button"
                    variant="ghost"
                    size="sm"
                    onClick={() => setExpandidos((atual) => ({ ...atual, [tipo]: !expandido }))}
                  >
                    {expandido ? 'Mostrar menos' : `Mostrar todos (${alertas.length})`}
                  </Button>
                )}
              </section>
            )
          })}
        </div>
      )}

      <p className="border-t border-border pt-3 text-xs text-muted-foreground">
        {`Limites: garantia em até ${limites.garantia_dias} dias, termo há mais de ${limites.termo_pendente_dias} dias, devolução há mais de ${limites.devolucao_pendente_dias} dias, manutenção parada há mais de ${limites.manutencao_dias} dias.`}
      </p>
    </div>
  )
}
