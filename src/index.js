import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  BorderStyle,
  LevelFormat,
  AlignmentType,
} from 'docx'
import { nodeTree } from './tree.js'
import lodash from 'lodash'
import { getListItemNumber } from './helper.js'
import { decode } from 'html-entities'

async function transform(lines, option) {
  if (!option) option = {}

  return await Promise.all(
    lines
      .map(async line => {
        if (line.type === 'list') {
          if (!option.numbering) option.numbering = []

          const { bullet, ...style } = line.style ?? {}

          const reference = 'list-custom-numbering' + option.numbering.length

          option.numbering.push({
            reference,
            levels: [
              {
                level: line.style.bullet.level,
                format: LevelFormat.BULLET,
                text: `${getListItemNumber(
                  line.style.bullet.format,
                  line.style.bullet.start,
                )}.`,
                alignment: AlignmentType.END,
                style: {
                  run: { ...style, underline: undefined },
                },
              },
            ],
          })

          return new Paragraph({
            ...style,
            children: [
              new TextRun({
                text: decode(line.run[0].content),
                ...style,
              }),
            ],
            bullet,
            numbering: {
              reference,
              level: bullet.level,
            },
          })
        } else if (line.type === 'table') {
          return new Table({
            ...(line.style ?? {}),
            borders: {
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              top: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: await Promise.all(
              line.run.map(async row => {
                const children = await Promise.all(
                  row.run.map(async cell => {
                    return new TableCell({
                      ...(cell.style ?? {}),
                      children: await transform(cell.run),
                    })
                  }),
                )

                return new TableRow({
                  ...(row.style ?? {}),
                  children,
                })
              }),
            ),
          })
        } else {
          return new Paragraph({
            ...lodash.merge(
              {
                spacing: {
                  after: 200,
                },
              },
              line.style ?? {},
            ),
            children: (line.run ?? []).filter(Boolean).map(item => {
              switch (item.type) {
                case 'text':
                  return new TextRun({
                    ...(item.style ?? {}),
                    text: decode(item.content),
                  })

                case 'image':
                  return new ImageRun({
                    ...(item.style ?? {}),
                    data: item.data,
                  })
                default:
                  return undefined
              }
            }),
          })
        }
      })
      .filter(Boolean),
  )
}

async function htmlToDocx(html, options) {
  const { top, bottom, left, right } = (options ?? {}).margins ?? {}

  return new Promise(async (resolve, reject) => {
    try {
      const tree = await nodeTree(html)

      const option = { numbering: [] }

      const sections = await Promise.all(
        Object.values(tree).map(async sheet => {
          return {
            properties: {
              page: {
                margin: {
                  top: top, 
                  bottom: bottom, 
                  left: left, 
                  right: right, 
                },
              }
            },
            children: await transform(sheet, option),
          }
        }),
      )

      const config = {
        numbering: {
          config: option.numbering ?? [],
        },
      }

      const doc = new Document({
        ...config,
        sections,
      })

      Packer.toBuffer(doc).then(buffer => resolve(buffer))
    } catch (error) {
      reject(error)
    }
  })
}

export default htmlToDocx
