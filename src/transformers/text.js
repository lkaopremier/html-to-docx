import Transformer from "./transform.js";

class TextTransformer extends Transformer {
    async transform() {
        return [{ run: [{ type: 'text', content: this.content()}] }]
    }

    content(){
        return this.node.rawText;
    }

    supports() {
        return ["text"];
    }
}

export default TextTransformer;