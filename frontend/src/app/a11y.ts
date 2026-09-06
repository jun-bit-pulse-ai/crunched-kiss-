/** Accessible names for the task pane. Pure helpers so they can be unit-tested. */

export function toolCardLabel(name: string, summary: string, error = false): string {
  const prefix = error ? "Tool error" : "Tool";
  const body = summary.trim();
  return body ? `${prefix} ${name}: ${body}` : `${prefix} ${name}`;
}

export function speakerLabel(role: "user" | "assistant" | "system"): string {
  if (role === "user") {
    return "You";
  }
  if (role === "system") {
    return "System";
  }
  return "Crunched";
}
