export const WRITE_DECLINED = "User declined this write.";

export type WritePreview = {
  sheet: string;
  address: string;
  values: unknown;
};

export function writePreview(input: Record<string, unknown>): WritePreview {
  return {
    sheet: String(input.sheet ?? ""),
    address: String(input.address ?? ""),
    values: input.values,
  };
}

export function writePreviewSummary(preview: WritePreview): string {
  return `${preview.sheet}!${preview.address}`;
}
