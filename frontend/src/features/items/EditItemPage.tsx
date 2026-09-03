import { useParams } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";

import { getItem } from "@/api/items";
import { ItemForm } from "@/features/items/ItemForm";
import { ErrorState, LoadingState } from "@/components/app/StateBlocks";

export function EditItemPage() {
  const { id } = useParams({ from: "/_shell/edit/$id" });
  const itemId = Number(id);

  const {
    data: item,
    isLoading,
    isError,
    error,
    refetch,
  } = useQuery({
    queryKey: ["item", itemId],
    queryFn: () => getItem(itemId),
    enabled: Number.isFinite(itemId),
  });

  if (isLoading) return <LoadingState label="Carregando item…" />;
  if (isError || !item)
    return (
      <ErrorState
        error={error}
        onRetry={() => void refetch()}
        title="Não foi possível carregar o item"
      />
    );

  return <ItemForm mode="edit" itemId={itemId} existingItem={item} />;
}
