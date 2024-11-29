import Transformer from "./transform.js";

class ListTransformer extends Transformer {
    transform() {
        return []
    }

    supports() {
        return ["ul", "ol"];
    }
}

export default ListTransformer;