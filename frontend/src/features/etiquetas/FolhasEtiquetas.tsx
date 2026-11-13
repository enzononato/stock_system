import type { Etiqueta } from '@/api/etiquetas'
import { calcularPosicoes, type LayoutEtiquetas } from '@/features/etiquetas/layout'

interface FolhasEtiquetasProps {
  etiquetas: Etiqueta[]
  /** SVG do QR de cada etiqueta, pelo id do equipamento. */
  qrs: Record<number, string>
  layout: LayoutEtiquetas
  /** 1 = primeira célula da primeira folha (aproveita folhas já usadas pela metade). */
  posicaoInicial: number
}

/** Milímetros sem ruído de ponto flutuante (0.1 + 0.2). */
function mm(n: number): string {
  return `${Number(n.toFixed(2))}mm`
}

/**
 * As folhas A4 prontas para a tela (prévia) e para a impressora: cada etiqueta é
 * posicionada em milímetros pelo layout. Papel é branco com tinta preta nos dois
 * temas, por isso aqui as cores são as palavras-chave `black`/`white` e não os
 * tokens do tema (exceção consciente).
 */
export function FolhasEtiquetas({ etiquetas, qrs, layout, posicaoInicial }: FolhasEtiquetasProps) {
  const posicoes = calcularPosicoes(layout, etiquetas.length, posicaoInicial)
  const totalFolhas = posicoes.length === 0 ? 0 : posicoes[posicoes.length - 1].folha + 1
  const lado = Math.min(layout.alturaMm - 4, layout.larguraMm * 0.45)

  return (
    <>
      {Array.from({ length: totalFolhas }, (_, folha) => (
        <div
          key={folha}
          className="folha"
          data-testid="folha"
          style={{
            width: '210mm',
            height: '297mm',
            position: 'relative',
            background: 'white',
            color: 'black',
            overflow: 'hidden',
          }}
        >
          {etiquetas.map((etiqueta, i) => {
            const pos = posicoes[i]
            if (pos.folha !== folha) return null
            const equipamento = [etiqueta.brand, etiqueta.model].filter(Boolean).join(' ')
            return (
              <div
                key={etiqueta.id}
                data-testid="etiqueta"
                style={{
                  position: 'absolute',
                  left: mm(pos.esquerdaMm),
                  top: mm(pos.topoMm),
                  width: mm(layout.larguraMm),
                  height: mm(layout.alturaMm),
                  boxSizing: 'border-box',
                  padding: '2mm',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '2mm',
                  overflow: 'hidden',
                }}
              >
                <div
                  role="img"
                  aria-label={`QR do patrimônio #${etiqueta.id}`}
                  style={{ width: mm(lado), height: mm(lado), flex: 'none' }}
                  className="[&>svg]:h-full [&>svg]:w-full"
                  dangerouslySetInnerHTML={{ __html: qrs[etiqueta.id] ?? '' }}
                />
                <div style={{ minWidth: 0, lineHeight: 1.15 }}>
                  <div style={{ fontSize: '12pt', fontWeight: 700, fontFamily: 'monospace' }}>#{etiqueta.id}</div>
                  <div style={{ fontSize: '8pt' }}>{etiqueta.tipo}</div>
                  <div
                    style={{
                      fontSize: '7pt',
                      display: '-webkit-box',
                      WebkitLineClamp: 2,
                      WebkitBoxOrient: 'vertical',
                      overflow: 'hidden',
                    }}
                  >
                    {equipamento}
                  </div>
                </div>
              </div>
            )
          })}
        </div>
      ))}
    </>
  )
}
