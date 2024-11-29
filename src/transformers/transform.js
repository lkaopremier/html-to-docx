import transformers from './index.js'

class Transformer {
    constructor(node) {
        this.node = node;
    }

    getChildren() {
        return this.node.childNodes ?? []
    }

    getType(){
        return this.node?.tagName?.toLowerCase() || "text";
    }

    async forward(node){
        const transformer =  transformers.getTransformer(node)
        const instance = transformer ? new transformer(node) : null
        return await instance?.transform()
    }
}

export default Transformer;