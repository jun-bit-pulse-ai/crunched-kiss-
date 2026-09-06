/** Turn a selected cell's formula into the prompt sent through the normal chat pipeline. */
export function buildExplainFormulaPrompt(formula: string): string {
  return `Explain this Excel formula in plain English (what it does, step by step): ${formula}`;
}
