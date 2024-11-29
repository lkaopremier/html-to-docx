import Transformer from "./transform.js";

class SpanTransformer extends Transformer {
    transform() {
        return []
    }

    supports() {
        return ["span"];
    }
}

export default SpanTransformer;