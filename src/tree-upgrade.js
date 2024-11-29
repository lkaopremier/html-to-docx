import { parse as htmlParser, TextNode } from 'node-html-parser'
import {
  pageNodes,
  normalizeHtml,
  trim,
  normalizeMeasure,
  capitalize,
  colorToHex,
  imageBase64ToBuffer,
  splitMeasure,
  convertCssToDocxMeasurement,
  parseCssMargin,
} from './helper.js'
import lodash from 'lodash'
import { AlignmentType, HeadingLevel } from 'docx'
import transformers from './transformers/index.js'

async function parseTreeNode(mainNode) {
  const nodes = mainNode instanceof TextNode ? [mainNode] : mainNode.childNodes;

  const sheet = await Promise.all(
    nodes.flatMap(node => {
      const transformer =  transformers.getTransformer(node)
      const instance = transformer ? new transformer(node) : null
      return instance?.transform()
    })
  );

  return sheet.flat();
}

export async function nodeTree(content) {
  const sheets = {}

  const pages = pageNodes(
    htmlParser(normalizeHtml(content)).querySelector('body'),
  )

  for (const index in pages) {
    sheets[index] = await parseTreeNode(pages[index])
  }

/*
  if(existsSync('./demo/tree.json')){
    unlinkSync('./demo/tree.json')
  }
  writeFileSync('./demo/tree.json', JSON.stringify(sheets, null, 2))
return;*/

  return sheets
}
