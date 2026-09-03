import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { isValidMac, isValidIp, maskMacInput } from "@/lib/utils";

interface FieldConfig {
  key: string;
  label: string;
  type?: "text" | "select";
  options?: string[];
  placeholder?: string;
  /** Aplicada a cada tecla digitada (ex.: formata MAC como AA:BB:CC:DD:EE:FF). */
  mask?: (value: string) => string;
  /** Validador usado por `validateTypeSpecificFields`. Campo vazio é sempre considerado válido. */
  validate?: (value: string) => boolean;
  /** Mensagem de erro exibida quando o campo foi preenchido e `validate` retorna false. */
  invalidMessage?: string;
}

// Os campos de MAC (endereco_fisico, mac) e IP (ip, ip_snmp) declaram aqui sua
// máscara/validador — este é o lugar natural, já que é aqui que cada tipo de
// equipamento define quais campos específicos existem. Nota fiscal não entra
// aqui: é um campo comum a todos os tipos, validado na tela que consome isto.
const TYPE_FIELDS: Record<string, FieldConfig[]> = {
  Celular: [
    { key: "identificador", label: "IMEI / Nº de Série" },
    { key: "storage", label: "Armazenamento (GB)" },
  ],
  Notebook: [
    { key: "dominio", label: "Domínio", type: "select", options: ["Sim", "Não"] },
    { key: "host", label: "Host" },
    {
      key: "endereco_fisico",
      label: "Endereço Físico (MAC)",
      mask: maskMacInput,
      validate: isValidMac,
      invalidMessage: "MAC inválido.",
    },
    { key: "cpu", label: "Processador (CPU)" },
    { key: "ram", label: "Memória RAM" },
    { key: "storage", label: "Armazenamento (GB)" },
    { key: "sistema", label: "Sistema Operacional" },
    { key: "licenca", label: "Licença Windows" },
    { key: "anydesk", label: "AnyDesk" },
  ],
  Desktop: [
    { key: "dominio", label: "Domínio", type: "select", options: ["Sim", "Não"] },
    { key: "host", label: "Host" },
    {
      key: "endereco_fisico",
      label: "Endereço Físico (MAC)",
      mask: maskMacInput,
      validate: isValidMac,
      invalidMessage: "MAC inválido.",
    },
    { key: "cpu", label: "Processador (CPU)" },
    { key: "ram", label: "Memória RAM" },
    { key: "storage", label: "Armazenamento (GB)" },
    { key: "sistema", label: "Sistema Operacional" },
    { key: "licenca", label: "Licença Windows" },
    { key: "anydesk", label: "AnyDesk" },
  ],
  Impressora: [
    { key: "setor", label: "Setor" },
    { key: "ip", label: "IP", validate: isValidIp, invalidMessage: "IP inválido." },
    {
      key: "mac",
      label: "MAC",
      mask: maskMacInput,
      validate: isValidMac,
      invalidMessage: "MAC inválido.",
    },
  ],
  Tablet: [
    { key: "identificador", label: "Nº de Série" },
    { key: "storage", label: "Armazenamento (GB)" },
  ],
  Switch: [
    { key: "poe", label: "POE", type: "select", options: ["Sim", "Não"] },
    { key: "quantidade_portas", label: "Qtd. de Portas" },
  ],
  HD: [{ key: "storage", label: "Armazenamento (GB)" }],
  Nobreak: [
    { key: "identificador", label: "Nº de Série" },
    { key: "codigo_patrimonial", label: "Código Patrimonial" },
    { key: "responsavel", label: "Responsável" },
    { key: "potencia_nominal", label: "Potência Nominal (VA/W)" },
    { key: "autonomia_estimada", label: "Autonomia Estimada (min)" },
    { key: "ip_snmp", label: "IP Placa SNMP", validate: isValidIp, invalidMessage: "IP inválido." },
  ],
  "Access Point": [
    { key: "identificador", label: "Nº de Série" },
    { key: "codigo_patrimonial", label: "Código Patrimonial" },
    { key: "local_instalacao", label: "Local de Instalação" },
    { key: "setor", label: "Setor" },
    { key: "ip", label: "IP", validate: isValidIp, invalidMessage: "IP inválido." },
    {
      key: "mac",
      label: "MAC",
      mask: maskMacInput,
      validate: isValidMac,
      invalidMessage: "MAC inválido.",
    },
  ],
};

/**
 * Valida os campos específicos preenchidos para o tipo informado. Campos
 * vazios são sempre válidos — quem decide se um campo é obrigatório é a tela
 * que consome este componente. Retorna a primeira mensagem de erro
 * encontrada, ou `null` se tudo estiver ok.
 */
export function validateTypeSpecificFields(
  tipo: string,
  values: Record<string, string>,
): string | null {
  const fields = TYPE_FIELDS[tipo] ?? [];
  for (const field of fields) {
    const value = values[field.key];
    if (value && field.validate && !field.validate(value)) {
      return field.invalidMessage ?? `${field.label} inválido.`;
    }
  }
  return null;
}

interface TypeSpecificFieldsProps {
  tipo: string;
  values: Record<string, string>;
  onChange: (key: string, value: string) => void;
  disabled?: boolean | undefined;
}

export function TypeSpecificFields({ tipo, values, onChange, disabled }: TypeSpecificFieldsProps) {
  const fields = TYPE_FIELDS[tipo] ?? [];
  if (fields.length === 0) return null;

  return (
    <>
      {fields.map((field) => (
        <div key={field.key} className="flex flex-col gap-1.5">
          <Label htmlFor={field.key}>{field.label}</Label>
          {field.type === "select" ? (
            <Select
              value={values[field.key] ?? ""}
              onValueChange={(v) => onChange(field.key, v)}
              disabled={disabled ?? false}
            >
              <SelectTrigger id={field.key}>
                <SelectValue placeholder={`Selecione ${field.label}`} />
              </SelectTrigger>
              <SelectContent>
                {field.options?.map((opt) => (
                  <SelectItem key={opt} value={opt}>
                    {opt}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          ) : (
            <Input
              id={field.key}
              value={values[field.key] ?? ""}
              onChange={(e) =>
                onChange(field.key, field.mask ? field.mask(e.target.value) : e.target.value)
              }
              placeholder={field.placeholder ?? field.label}
              disabled={disabled}
            />
          )}
        </div>
      ))}
    </>
  );
}
