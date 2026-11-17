/** Mesmo padrão simples do backend: algo@dominio.ext, sem espaços. */
export function isValidEmail(v: string): boolean {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v.trim())
}
