import Transformer from "./transform.js";

class PTransformer extends Transformer {
  async transform() {
    if (this.getType() === "br") {
      return [{ type: "break", run: [] }];
    }

    const output = [];

    for (const node of this.getChildren()) {
      const tagName = node?.tagName?.toLowerCase();

      if (!this.supports().includes(tagName)) {
        const forward = await this.forward(node);
        if (forward) {
          for (const item of forward) {
            output.push(...item.run);
          }
        }
        continue;
      }

      if (tagName === "br") {
        if (output.length === 0) {
          output.push({ content: "", style: { break: 1 } });
        } else {
          const lastRun = output[output.length - 1];
          if (lastRun?.type === "text") {
            lastRun.style = lastRun.style || {};
            lastRun.style.break = 1;
          }
        }
      }
    }

    return [{ run: output }];
  }

  supports() {
    return ["p", "h1", "h2", "h3", "h4", "h5", "h6", "br"];
  }
}

export default PTransformer;
