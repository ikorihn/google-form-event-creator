export function formatLiteral(text: string): string {
  return text.replace(/\n +\|/, "\n")
}

export function addBr(text: string): string {
  return text.replace(/\n/, "<br/>\n");
}