import { useQuery } from "@tanstack/react-query";

import { listItemsPaginated, type Item } from "@/api/items";

export interface Collaborator {
  nome: string;
  cpf?: string | undefined;
  setor?: string | undefined;
  revenda?: string | undefined;
  items: Item[];
}

/**
 * O backend não expõe um cadastro de colaboradores. A única fonte real de
 * colaboradores hoje é o próprio patrimônio: itens emprestados carregam
 * `assigned_to`, `cpf`, `setor` e `revenda`. Derivamos a lista daí — sem
 * inventar dados — e reaproveitamos a mesma consulta para saber quais itens
 * estão vinculados a cada pessoa.
 */
export function useCollaborators() {
  const query = useQuery({
    queryKey: ["items", "assigned", { limit: 500 }],
    queryFn: () => listItemsPaginated({ status: "emprestado", limit: 500, offset: 0 }),
  });

  const byName = new Map<string, Collaborator>();
  for (const item of query.data?.items ?? []) {
    const nome = (item.assigned_to ?? "").trim();
    if (!nome) continue;
    const current = byName.get(nome);
    if (current) {
      current.items.push(item);
    } else {
      byName.set(nome, {
        nome,
        cpf: item.cpf,
        setor: item.setor,
        revenda: item.revenda,
        items: [item],
      });
    }
  }

  return {
    ...query,
    collaborators: [...byName.values()].sort((a, b) => a.nome.localeCompare(b.nome)),
  };
}

export function useCollaboratorItems(nome: string | undefined) {
  const { collaborators, isLoading, error, refetch } = useCollaborators();
  const found = nome ? collaborators.find((c) => c.nome === nome) : undefined;
  return { items: found?.items ?? [], isLoading, error, refetch };
}
