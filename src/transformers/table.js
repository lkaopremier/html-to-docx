import Transformer from "./transform.js";

class TableTransformer extends Transformer {
    transform() {
        return []
    }

    supports() {
        return ["text", "span"];
    }
}

export default TableTransformer;