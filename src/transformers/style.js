import Transformer from "./transform.js";

class StyleTransformer extends Transformer {
    transform() {
        return []
    }

    supports() {
        return ['strong', 'a', 'i', 's', 'u', 'b', 'em'];
    }
}

export default StyleTransformer;