import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { htmlToText } from "../../src/mail/sanitizer.js";

describe("mail sanitizer", () => {
  it("removes script/style tags and decodes common entities", () => {
    assert.equal(
      htmlToText("<style>x</style><p>Hello&nbsp;Contoso &amp; Fabrikam</p><script>x</script>"),
      "Hello Contoso & Fabrikam"
    );
  });
});
