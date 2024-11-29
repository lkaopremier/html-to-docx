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
import { decode } from 'html-entities'

function getStyle(node) {
  const style = {}

  switch (node?.tagName?.toLowerCase()) {
    case 'strong':
    case 'b':
      style.bold = true
      break

    case 'em':
    case 'i':
      style.italics = true
      break

    case 'u':
      style.underline = {}
      break

    case 's':
      style.strike = {}
      break

    case 'h1':
      style.heading = HeadingLevel.HEADING_1
      break

    case 'h2':
      style.heading = HeadingLevel.HEADING_2
      break

    case 'h3':
      style.heading = HeadingLevel.HEADING_3
      break

    case 'h4':
      style.heading = HeadingLevel.HEADING_4
      break

    case 'h5':
      style.heading = HeadingLevel.HEADING_5
      break

    case 'h6':
      style.heading = HeadingLevel.HEADING_6
      break
  }

  lodash.merge(style, getCurrentNodeStyle(node))

  if (node.parentNode) {
    return lodash.merge(getStyle(node.parentNode), style)
  }

  return style
}

function getCurrentNodeStyle(node) {
  const style = {}
  if (node?.attributes?.style) {
    const styles = node.attributes?.style?.split(';') ?? []

    styles.forEach(item => {
      const [key, value] = item.split(':').map(s => s.trim())
      switch (key) {
        case 'font-weight':
          if (value === 'bold') {
            style.bold = true
          }
          break

        case 'font-style':
          if (value === 'italic') {
            style.italics = true
          }
          break

        case 'text-decoration':
          if (value === 'underline') {
            style.underline = {}
          }
          break

        case 'text-transform':
          if (value === 'uppercase') {
            style.uppercase = true
          } else if (value === 'capitalize') {
            style.capitalize = true
          } else if (value === 'lowercase') {
            style.lowercase = true
          } else if (value === 'invertcase') {
            style.invertcase = true
          } else if (value === 'uppercasesentence') {
            style.uppercasesentence = true
          }
          break

        case 'font-family':
          style.font = trim(trim(value, "'"), '"')
          break

        case 'font-size':
          style.size = normalizeMeasure(value)
          break

        case 'color':
          style.color = colorToHex(value)
          break

        case 'margin':
        case 'margin-top':
        case 'margin-bottom':
          const { top, bottom } = parseCssMargin(value)

          if (key === 'margin') {
            style.spacing = {
              before: top,
              after: bottom,
            }
          } else if (key === 'margin-top') {
            style.spacing = {
              before: top,
            }
          } else if (key === 'margin-bottom') {
            style.spacing = {
              after: bottom,
            }
          }
          break

        case 'text-align':
          switch (value) {
            case 'left':
              style.alignment = AlignmentType.LEFT
              break

            case 'right':
              style.alignment = AlignmentType.RIGHT
              break

            case 'center':
              style.alignment = AlignmentType.CENTER
              break

            case 'justify':
              style.alignment = AlignmentType.JUSTIFIED
              break
          }
          break

        case 'text-indent':
          style.indent = { left: normalizeMeasure(value) }
          break

        case 'width':
          if (
            ['table', 'tr', 'th', 'td'].includes(node?.tagName.toLowerCase())
          ) {
            style.width = convertCssToDocxMeasurement(value)
          } else {
            style.width = normalizeMeasure(value)
          }
          break

        case 'height':
          style.height = normalizeMeasure(value)
          break
      }
    })
  }

  return style
}

async function paragraphNodeParse(
  sheet,
  node,
  paragraph = null,
  newParagraphAfterBr = false,
) {
  if (node instanceof TextNode) {
    const style = getStyle(node?.parentNode) ?? {}
    const textRun = {
      type: 'text',
      style,
      content: node.rawText,
    }

    if (!paragraph || newParagraphAfterBr) {
      paragraph = { run: [textRun] }
    } else {
      paragraph.run.push(textRun)
    }
  } else if (node.tagName?.toLowerCase() === 'img') {
    const image =
      node.attributes.src && (await imageBase64ToBuffer(node.attributes.src))

    if (image) {
      const style = getCurrentNodeStyle(node)

      if (style.width && !style.height) {
        const [numericValue] = splitMeasure(style.width)
        style.width = numericValue
        style.height = numericValue / image.ratio
      } else if (!style.width && style.height) {
        const [numericValue] = splitMeasure(style.height)
        style.height = numericValue
        style.width = numericValue * image.ratio
      }

      const imageRun = {
        type: 'image',
        style: {
          type: image.extension,
          ...getStyle(node),
          transformation: {
            width: image.width,
            height: image.height,
            ...style,
          },
        },
        data: image.buffer,
      }

      if (!paragraph || newParagraphAfterBr) {
        paragraph = { run: [imageRun] }
      } else {
        paragraph.run.push(imageRun)
      }
    }
  } else if (node.tagName?.toLowerCase() === 'br') {
    if (paragraph) {
      sheet.push(normalizeParagraph(paragraph))
      paragraph = null
    }

    if (
      ['br', 'table', 'p'].includes(
        node.previousElementSibling?.tagName?.toLowerCase(),
      )
    ) {
      sheet.push({ type: 'break', run: [] })
    }
  } else if (node.childNodes) {
    for (const child of node.childNodes) {
      paragraph = await paragraphNodeParse(
        sheet,
        child,
        paragraph,
        newParagraphAfterBr,
      )
      newParagraphAfterBr = false
    }
  }

  return paragraph
}

function reflectStyleToParagraph(paragraph) {
  if (!paragraph || (paragraph.run ?? []).length === 0) return paragraph
  let style = {}

  for (const text of paragraph.run) {
    style = { ...style, ...(text.style ?? {}) }
  }

  const { alignment, indent, heading, spacing } = style

  if (!paragraph.style) paragraph.style = {}

  if (alignment) {
    paragraph.style.alignment = alignment
  }

  if (indent) {
    paragraph.style.indent = indent
  }

  if (heading) {
    paragraph.style.heading = heading
  }

  if (spacing) {
    paragraph.style.spacing = spacing
  }

  return paragraph
}

function applyStyle(paragraph) {
  if (!paragraph || (paragraph.run ?? []).length === 0) return paragraph
  paragraph.run = paragraph.run.map(item => {
    if (['table', 'row', 'cell', 'image'].includes(item.type)) return item

    const style = item.style ?? {}
    let content = item?.content?.length > 0 ? item.content : null

    if (content) {
      if (style.uppercase) {
        content = content.toUpperCase()
      } else if (style.capitalize) {
        content = capitalize(content, true)
      } else if (style.lowercase) {
        content = content.toLowerCase()
      } else if (style.invertcase) {
        content = [...content]
          .map(char =>
            char === char.toUpperCase()
              ? char.toLowerCase()
              : char.toUpperCase(),
          )
          .join('')
      } else if (style.uppercasesentence) {
        content = content
          .toLowerCase()
          .split('.')
          .map(sentence => {
            const firstAlphaIndex = sentence.search(/[a-zA-Z]/)
            if (firstAlphaIndex !== -1) {
              return (
                sentence.slice(0, firstAlphaIndex) +
                sentence[firstAlphaIndex].toUpperCase() +
                sentence.slice(firstAlphaIndex + 1)
              )
            }
            return sentence
          })
          .join('.')
      }
    }

    return { ...item, content }
  })

  return paragraph
}

function trimContent(content, isFirst, isLast, isSingle) {
  if (isSingle) return decode(content.trim())
  if (isFirst) return decode(content.trimStart())
  if (isLast) return decode(content.trimEnd())
  return decode(content)
}

function trimParagraph(paragraph) {
  if (!paragraph || (paragraph.run ?? []).length === 0) return paragraph

  const run = []
  let foundNonEmptyText = false

  for (const line of paragraph.run) {
    if (paragraph?.type === 'list' && line.content.trim().length === 0) {
      continue
    } else if (
      !foundNonEmptyText &&
      line.type === 'text' &&
      line.content.trim().length === 0
    ) {
      continue
    }

    foundNonEmptyText = true

    if (paragraph?.type === 'list') {
      run.push({ ...line, content: line.content.trim() })
    } else {
      run.push(line)
    }
  }

  while (
    run.length > 0 &&
    run[run.length - 1].type === 'text' &&
    run[run.length - 1].content.trim().length === 0
  ) {
    run.pop()
  }

  return {
    ...paragraph,
    run: run.map((item, index) => {
      if (item.content === undefined) return item

      const isFirst = index === 0
      const isLast = index === run.length - 1
      const isSingle = run.length === 1

      return {
        ...item,
        content: trimContent(item.content, isFirst, isLast, isSingle),
      }
    }),
  }
}

function normalizeParagraph(paragraph) {
  return applyStyle(reflectStyleToParagraph(trimParagraph(paragraph)))
}

function getListFormat(node) {
  const items = (node.attributes?.style?.split(';') ?? [])
    .map(item => item.trim())
    .filter(item => item.startsWith('list-style'))
  for (const item of items) {
    const [key, value] = item.split(':').map(s => s.trim())

    if (key === 'list-style-type' || key === 'list-style') {
      return value
    }
  }

  switch (node?.tagName?.toLowerCase()) {
    case 'ol':
      return 'decimal'

    default:
      return 'bullet'
  }
}

function getListLevel(node, level = 0) {
  level = level ?? 0

  if (typeof node?.parentNode !== 'undefined') {
    if (['ul', 'ol'].includes(node?.parentNode?.tagName?.toLowerCase())) {
      return getListLevel(node.parentNode, level + 1)
    } else {
      return getListLevel(node.parentNode, level)
    }
  }

  return level
}

function normalizeSheet(sheet) {
  const items = []

  for (const index in sheet) {
    const item = sheet[index]
    switch (item.type) {
      case 'list':
        let i = 0

        const bullet = item.bullet ?? {}

        const style = {
          ...(item.style ?? {}),
          bullet,
        }

        const start =
          bullet.start && bullet.start > 0 ? bullet.start : undefined

        for (const listItem of item.run) {
          i = start ?? i + 1

          items.push({
            type: 'list',
            style: lodash.merge(lodash.cloneDeep(style), listItem.style ?? {}, {
              bullet: { start: i },
            }),
            run: [
              {
                type: 'text',
                content: listItem.content,
                style: {},
              },
            ],
          })
        }
        break

      default:
        items.push(item)
        break
    }
  }

  return items
}

async function parseTreeNode(mainNode) {
  const sheet = []

  let paragraph = null

  for (const node of mainNode.childNodes) {
    if (node instanceof TextNode) {
      paragraph = await paragraphNodeParse(sheet, node, paragraph)
    } else {
      switch (node.tagName?.toLowerCase()) {
        case 'span':
        case 'strong':
        case 'a':
        case 'i':
        case 's':
        case 'u':
        case 'b':
        case 'em':
        case 'h1':
        case 'h2':
        case 'h3':
        case 'h4':
        case 'h5':
        case 'h6':
        case 'p':
        case 'ul':
        case 'ol':
        case 'li':
          if (
            ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol'].includes(
              node.tagName?.toLowerCase(),
            ) &&
            paragraph !== null
          ) {
            sheet.push(normalizeParagraph(paragraph))
            paragraph = null
          }

          if (
            ['ul', 'ol'].includes(node.tagName?.toLowerCase()) &&
            paragraph === null
          ) {
            paragraph = {
              type: 'list',
              bullet: {
                level: getListLevel(node),
                format: getListFormat(node),
                start:
                  node.attributes?.start && parseInt(node.attributes?.start),
              },
              run: [],
            }
          }

          for (const child of node.childNodes) {
            paragraph = await paragraphNodeParse(sheet, child, paragraph)
          }
          break

        case 'br':
          if (paragraph !== null) {
            sheet.push(normalizeParagraph(paragraph))
            paragraph = null
          }

          if (
            ['br', 'table', 'p'].includes(
              node.previousElementSibling?.tagName?.toLowerCase(),
            )
          ) {
            sheet.push({ type: 'break', run: [] })
          }
          break

        case 'img':
          paragraph = await paragraphNodeParse(sheet, node, paragraph)
          break

        case 'table':
          if (paragraph !== null) {
            sheet.push(normalizeParagraph(paragraph))
            paragraph = null
          }

          paragraph = {
            type: 'table',
            style: getCurrentNodeStyle(node),
            run: await Promise.all(
              Array.from(node.querySelectorAll('tbody > tr')).map(
                async trNode => {
                  const cells = await Promise.all(
                    Array.from(trNode.querySelectorAll('td')).map(
                      async tdNode => {
                        return {
                          type: 'cell',
                          style: getCurrentNodeStyle(tdNode),
                          run: await parseTreeNode(tdNode),
                        }
                      },
                    ),
                  )

                  return {
                    type: 'row',
                    style: getCurrentNodeStyle(trNode),
                    run: cells,
                  }
                },
              ),
            ),
          }
          break

        default:
          break
      }
    }
  }

  if (paragraph !== null) {
    sheet.push(normalizeParagraph(paragraph))
  }

  return sheet
    .map(item => {
      if (item.type !== 'break' && item?.run?.length === 0) {
        return undefined
      }

      return item
    })
    .filter(Boolean)
}

export async function nodeTree(content) {
  const sheets = {}

  const pages = pageNodes(
    htmlParser(normalizeHtml(content)).querySelector('body'),
  )

  for (const index in pages) {
    sheets[index] = normalizeSheet(await parseTreeNode(pages[index]))
  }

  return sheets
}
