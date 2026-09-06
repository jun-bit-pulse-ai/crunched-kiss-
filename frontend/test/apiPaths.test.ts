import assert from "assert";
import { CHAT_PATH, HEALTH_PATH } from "../src/app/apiPaths";

describe("api paths", () => {
  it("uses same-origin /api routes so webpack can proxy", () => {
    assert.strictEqual(CHAT_PATH, "/api/chat");
    assert.strictEqual(HEALTH_PATH, "/api/health");
  });
});
