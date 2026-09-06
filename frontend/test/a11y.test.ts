import assert from "assert";
import { speakerLabel, toolCardLabel } from "../src/app/a11y";

describe("toolCardLabel", () => {
  it("includes the tool name and summary", () => {
    assert.strictEqual(
      toolCardLabel("list_workbook_meta", "Data 5000×200 · Budget 6×4"),
      "Tool list_workbook_meta: Data 5000×200 · Budget 6×4"
    );
  });

  it("marks errors so color is not the only signal", () => {
    assert.strictEqual(
      toolCardLabel("read_range", "Worksheet Missing not found", true),
      "Tool error read_range: Worksheet Missing not found"
    );
  });

  it("still names the tool when the summary is empty", () => {
    assert.strictEqual(toolCardLabel("find", "  "), "Tool find");
  });
});

describe("speakerLabel", () => {
  it("names the speaker for each chat role", () => {
    assert.strictEqual(speakerLabel("user"), "You");
    assert.strictEqual(speakerLabel("assistant"), "Crunched");
    assert.strictEqual(speakerLabel("system"), "System");
  });
});
