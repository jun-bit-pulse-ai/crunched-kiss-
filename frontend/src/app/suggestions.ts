/**
 * Parses a 💡 suggestion line from assistant text.
 * Format: 💡 "Suggestion 1" · "Suggestion 2" · "Suggestion 3"
 * Returns null if no suggestion line is found.
 * Also returns the text with the suggestion line removed.
 */
export function parseSuggestions(text: string): { text: string; suggestions: string[] } | null {
  const lines = text.split("\n");
  const lastLine = lines[lines.length - 1].trim();

  if (!lastLine.startsWith("💡")) {
    return null;
  }

  // Extract quoted strings: "text" · "text" · "text"
  const suggestions: string[] = [];
  const quoteRegex = /"([^"]+)"/g;
  let match: RegExpExecArray | null;
  while ((match = quoteRegex.exec(lastLine)) !== null) {
    suggestions.push(match[1].trim());
  }

  if (suggestions.length === 0) {
    return null;
  }

  // Remove the suggestion line from text
  const textWithoutSuggestions = lines.slice(0, -1).join("\n").trim();
  return { text: textWithoutSuggestions, suggestions };
}
