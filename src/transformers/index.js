import text from './text.js';
import img from './img.js';
import list from './list.js';
import p from './p.js';
import style from './style.js';
import table from './table.js';
import span from './span.js';

const transformers = [text, img, list, p, style, span, table];

export default {
  transformers,
  getTransformer(node) {
    if(!node){
        return null
    }

    const type = node?.tagName?.toLowerCase() || "text";

    return transformers
        .find(transformer => (new transformer(node)).supports().includes(type)) || null;
  }
};