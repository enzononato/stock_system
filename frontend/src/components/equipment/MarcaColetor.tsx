/**
 * Chave do campo -> valor que ele tinha antes do coletor. A chave presente marca o campo como
 * "do coletor"; um texto (vazio = o campo estava vazio) acrescenta "antes: ...", usado na
 * edição, quando o valor mudou. `undefined` marca sem mostrar o valor antigo.
 */
export type CamposMarcados = Record<string, string | undefined>

/** Marca ao lado do rótulo de um campo preenchido pelo coletor (e o valor anterior, se mudou). */
export function MarcaColetor({ marcados, campo }: { marcados?: CamposMarcados; campo: string }) {
  if (!marcados || !(campo in marcados)) return null
  const antes = marcados[campo]
  return (
    <span className="flex flex-wrap items-center gap-x-2 text-[11px] leading-tight">
      <span className="rounded border border-border bg-brand-soft px-1.5 py-0.5 font-semibold text-brand-soft-foreground">
        do coletor
      </span>
      {antes !== undefined && (
        <span className="text-muted-foreground">antes: {antes === '' ? 'vazio' : antes}</span>
      )}
    </span>
  )
}
