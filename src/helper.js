import { parse as htmlParser } from 'node-html-parser'
import { minify } from 'html-minifier'
import sharp from 'sharp'

export function normalizeHtml(content) {
  const elm = htmlParser(content)

  const head =
    elm.querySelector('head')?.outerHTML ??
    `<head><meta charset="UTF-8"><title>Document</title></head>`

  const body = cleanHtmlContent(
    elm.querySelector('body')?.outerHTML?.trim() ??
      `<body>${cleanHtmlContent(elm.outerHTML).trim()}</body>`,
  )

  const doc = minify(`<!DOCTYPE html><html>${head}${body}</html>`, {
    removeAttributeQuotes: true,
    removeEmptyAttributes: true,
    removeComments: true,
  })

  return cleanHtmlContent(doc)
}

export function cleanHtmlContent(content) {
  return content.replace(/(\s{2,}|\n)/g, match => {
    if (/<(pre|code)>/.test(match)) {
      return match
    }

    return ' '
  })
}

export function pageNodes(bodyElm) {
  const pages = {}
  let children = []
  bodyElm.childNodes.forEach(child => {
    if (child && child.classList?.contains('page-break')) {
      pages[Object.keys(pages).length + 1] = arrayToNodeHTMLElement(children)
      children = []
    } else {
      children.push(child)
    }
  })

  if (children.length > 0) {
    pages[Object.keys(pages).length + 1] = arrayToNodeHTMLElement(children)
  }

  return Object.values(pages)
}

export function arrayToNodeHTMLElement(array) {
  const fragment = htmlParser('')
  array.forEach(item => fragment.appendChild(item))
  return fragment
}

export function trim(str, ch) {
  var start = 0,
    end = str.length

  while (start < end && str[start] === ch) ++start

  while (end > start && str[end - 1] === ch) --end

  return start > 0 || end < str.length ? str.substring(start, end) : str
}

export function splitMeasure(value) {
  const match = value?.toString().match(/^(\d+(\.\d+)?)(px|pt|cm|in|mm|pc|pi)$/)

  if (!match) {
    return []
  }

  return [parseFloat(match[1]), match[3]]
}

export function normalizeMeasure(value) {
  const [numericValue, unit] = splitMeasure(value)

  if (!numericValue || !unit) {
    return ''
  }

  if (unit === 'px') {
    return `${numericValue * 0.75}pt`
  } else if (['pt', 'in', 'cm', 'mm', 'pc', 'pi'].includes(unit)) {
    return `${numericValue}${unit}`
  }
}

/**
 * Capitalizes first letters of words in string.
 * @param {string} str String to be modified
 * @param {boolean=false} lower Whether all other letters should be lowercased
 * @return {string}
 * @usage
 *   capitalize('fix this string');     // -> 'Fix This String'
 *   capitalize('javaSCrIPT');          // -> 'JavaSCrIPT'
 *   capitalize('javaSCrIPT', true);    // -> 'Javascript'
 */
export const capitalize = (str, lower = false) => {
  return (lower ? str.toLowerCase() : str).replace(
    /(?:^|\s|["'([{])+\S/g,
    match => match.toUpperCase(),
  )
}

export function colorToHex(color) {
  const namedColors = {
    red: '#FF0000',
    blue: '#0000FF',
    green: '#008000',
    black: '#000000',
    white: '#FFFFFF',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    pink: '#FFC0CB',
    gray: '#808080',
    silver: '#C0C0C0',
    maroon: '#800000',
    olive: '#808000',
    lime: '#00FF00',
    teal: '#008080',
    navy: '#000080',
    aqua: '#00FFFF',
    fuchsia: '#FF00FF',
    cyan: '#00FFFF',
    brown: '#A52A2A',
    gold: '#FFD700',
    coral: '#FF7F50',
    violet: '#EE82EE',
    indigo: '#4B0082',
    khaki: '#F0E68C',
    salmon: '#FA8072',
    chocolate: '#D2691E',
    tan: '#D2B48C',
    azure: '#F0FFFF',
    beige: '#F5F5DC',
    lavender: '#E6E6FA',
    crimson: '#DC143C',
    turquoise: '#40E0D0',
    ivory: '#FFFFF0',
    orchid: '#DA70D6',
    plum: '#DDA0DD',
    sienna: '#A0522D',
    midnightblue: '#191970',
    seashell: '#FFF5EE',
    tomato: '#FF6347',
    snow: '#FFFAFA',
    mintcream: '#F5FFFA',
    wheat: '#F5DEB3',
    moccasin: '#FFE4B5',
    hotpink: '#FF69B4',
    skyblue: '#87CEEB',
    slategray: '#708090',
    darkblue: '#00008B',
    darkgreen: '#006400',
    darkred: '#8B0000',
    lightblue: '#ADD8E6',
    lightgreen: '#90EE90',
    lightpink: '#FFB6C1',
    lightgray: '#D3D3D3',
  }

  if (namedColors[color.toLowerCase()]) {
    return namedColors[color.toLowerCase()]
  }

  const hexRegex = /^#([0-9A-F]{3}|[0-9A-F]{4}|[0-9A-F]{6}|[0-9A-F]{8})$/i
  if (hexRegex.test(color)) {
    return color.toUpperCase()
  }

  const rgbRegex =
    /^rgba?\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})(?:,\s*(\d?\.?\d+))?\)$/i
  const rgbMatch = color.match(rgbRegex)
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0')
    const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0')
    const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0')
    const a =
      rgbMatch[4] !== undefined
        ? Math.round(parseFloat(rgbMatch[4]) * 255)
            .toString(16)
            .padStart(2, '0')
        : null

    return `#${r}${g}${b}${a || ''}`.toUpperCase()
  }

  const hslRegex =
    /^hsla?\((\d{1,3}),\s*([\d.]+%)?,\s*([\d.]+%)?(?:,\s*(\d?\.?\d+))?\)$/i
  const hslMatch = color.match(hslRegex)
  if (hslMatch) {
    const h = parseInt(hslMatch[1]) % 360
    const s = parseFloat(hslMatch[2]) / 100
    const l = parseFloat(hslMatch[3]) / 100
    const a = hslMatch[4] !== undefined ? parseFloat(hslMatch[4]) : 1

    const rgb = hslToRgb(h, s, l)
    const r = rgb.r.toString(16).padStart(2, '0')
    const g = rgb.g.toString(16).padStart(2, '0')
    const b = rgb.b.toString(16).padStart(2, '0')
    const alpha =
      a < 1
        ? Math.round(a * 255)
            .toString(16)
            .padStart(2, '0')
        : null

    return `#${r}${g}${b}${alpha || ''}`.toUpperCase()
  }

  return '#00000'
}

export function hslToRgb(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1))
  const m = l - c / 2

  let r = 0,
    g = 0,
    b = 0
  if (h < 60) {
    r = c
    g = x
    b = 0
  } else if (h < 120) {
    r = x
    g = c
    b = 0
  } else if (h < 180) {
    r = 0
    g = c
    b = x
  } else if (h < 240) {
    r = 0
    g = x
    b = c
  } else if (h < 300) {
    r = x
    g = 0
    b = c
  } else {
    r = c
    g = 0
    b = x
  }

  return {
    r: Math.round((r + m) * 255),
    g: Math.round((g + m) * 255),
    b: Math.round((b + m) * 255),
  }
}

export function calculateRatio(w, h) {
  if (typeof w !== 'number' || typeof h !== 'number' || h === 0) {
    throw new Error(
      'Les paramètres doivent être des nombres, et h ne doit pas être 0',
    )
  }

  return (w / h).toFixed(2)
}

export async function imageBase64ToBuffer(base64Content) {
  if (!base64Content.startsWith('data:')) {
    return undefined
  }

  const mimeTypeMatch = base64Content.match(/^data:(.+);base64,/)
  if (!mimeTypeMatch) {
    return undefined
  }

  const mimeType = mimeTypeMatch[1]

  const base64Data = base64Content.replace(/^data:.+;base64,/, '')

  const buffer = Buffer.from(base64Data, 'base64')

  const mimeToExtension = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/svg+xml': 'svg',
  }

  const extension = mimeToExtension[mimeType]
  if (!extension) {
    return undefined
  }

  const metadata = await sharp(buffer).metadata()
  const { width, height } = metadata

  if (!width || !height) {
    return undefined
  }

  const ratio = calculateRatio(width, height)

  return { mimeType, extension, buffer, width, height, ratio }
}

export function random(length) {
  return Math.random().toString(36).substring(length)
}

export function convertCssToDocxMeasurement(cssValue) {
  const auto = { size: 0, type: 'nil' }

  if (typeof cssValue !== 'string') {
    return auto
  }

  if (cssValue.endsWith('%')) {
    const percentage = parseFloat(cssValue)

    if (isNaN(percentage)) {
      return auto
    }

    return { size: percentage, type: 'pct' }
  }

  if (cssValue === 'auto' || cssValue === '0' || cssValue === 'none') {
    return auto
  }

  if (cssValue.endsWith('px')) {
    const pixels = parseFloat(cssValue)
    if (isNaN(pixels)) {
      return auto
    }

    return { size: Math.round(pixels * 15), type: 'dxa' }
  }

  if (cssValue.endsWith('pt')) {
    const points = parseFloat(cssValue)
    if (isNaN(points)) {
      return auto
    }

    return { size: Math.round(points * 20), type: 'dxa' }
  }

  return auto
}

/**
 * Convertit une valeur CSS de marge (ex: "15px", "10px 20px", "5px 10px 15px 20px")
 * en un objet avec `top`, `left`, `bottom`, et `right` (espacement en twips).
 *
 * @param cssMargin - Une chaîne représentant les marges CSS.
 * @returns Un objet contenant les marges `top`, `left`, `bottom`, et `right` en twips.
 */
export function parseCssMargin(cssMargin) {
  const toTwips = value => {
    if (value.endsWith('px')) {
      return Math.round(parseFloat(value) * 15)
    }
    if (value.endsWith('pt')) {
      return Math.round(parseFloat(value) * 20)
    }
    return 0
  }

  const parts = cssMargin.split(' ').map(part => part.trim())

  let top, right, bottom, left

  if (parts.length === 1) {
    top = right = bottom = left = toTwips(parts[0])
  } else if (parts.length === 2) {
    top = bottom = toTwips(parts[0])
    right = left = toTwips(parts[1])
  } else if (parts.length === 3) {
    top = toTwips(parts[0])
    right = left = toTwips(parts[1])
    bottom = toTwips(parts[2])
  } else if (parts.length === 4) {
    top = toTwips(parts[0])
    right = toTwips(parts[1])
    bottom = toTwips(parts[2])
    left = toTwips(parts[3])
  } else {
    top = right = bottom = left = 0
  }

  return { top, right, bottom, left }
}

export function getListItemNumber(styleType, start) {
  function toRoman(num, isUpper = false) {
    const romanNumerals = [
      { value: 1000, numeral: 'M' },
      { value: 900, numeral: 'CM' },
      { value: 500, numeral: 'D' },
      { value: 400, numeral: 'CD' },
      { value: 100, numeral: 'C' },
      { value: 90, numeral: 'XC' },
      { value: 50, numeral: 'L' },
      { value: 40, numeral: 'XL' },
      { value: 10, numeral: 'X' },
      { value: 9, numeral: 'IX' },
      { value: 5, numeral: 'V' },
      { value: 4, numeral: 'IV' },
      { value: 1, numeral: 'I' },
    ]

    let result = ''
    let tempNum = num

    for (let i = 0; i < romanNumerals.length; i++) {
      while (tempNum >= romanNumerals[i].value) {
        result += romanNumerals[i].numeral
        tempNum -= romanNumerals[i].value
      }
    }

    return isUpper ? result.toUpperCase() : result.toLowerCase()
  }

  function toAlpha(num, isUpper = false) {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz'
    let result = ''
    while (num > 0) {
      num--
      result = alphabet[num % 26] + result
      num = Math.floor(num / 26)
    }
    return isUpper ? result.toUpperCase() : result.toLowerCase()
  }

  function toGreek(num, isUpper = false) {
    const greekAlphabet = [
      'α',
      'β',
      'γ',
      'δ',
      'ε',
      'ζ',
      'η',
      'θ',
      'ι',
      'κ',
      'λ',
      'μ',
      'ν',
      'ξ',
      'ο',
      'π',
      'ρ',
      'σ',
      'τ',
      'υ',
      'φ',
      'χ',
      'ψ',
      'ω',
    ]
    let result = greekAlphabet[(num - 1) % 24]
    return isUpper ? result.toUpperCase() : result.toLowerCase()
  }

  function toLetter(num, isUpper = false) {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz'
    let letter = alphabet[(num - 1) % 26]
    return isUpper ? letter.toUpperCase() : letter.toLowerCase()
  }

  switch (styleType) {
    case 'decimal':
      return start
    case 'upper-roman':
      return toRoman(start, true)
    case 'lower-roman':
      return toRoman(start, false)
    case 'upper-alpha':
      return toAlpha(start, true)
    case 'lower-alpha':
      return toAlpha(start, false)
    case 'upper-letter':
      return toLetter(start, true)
    case 'lower-letter':
      return toLetter(start, false)
    case 'upper-greek':
      return toGreek(start, true)
    case 'lower-greek':
      return toGreek(start, false)
    case 'circle':
      return '○'
    case 'disc':
      return '•'
    case 'square':
      return '▪'
    case 'none':
      return ''
    default:
      return start
  }
}

export function getWordIndent(level) {
  const baseIndent = 720 // 720 TWIP = 0.5 pouce
  const hangingIndent = 360 // 360 TWIP = 0.25 pouce

  const left = level * baseIndent

  return {
    left: left,
    hanging: hangingIndent,
  }
}

export default {
  normalizeHtml,
  pageNodes,
  trim,
  colorToHex,
  hslToRgb,
  normalizeMeasure,
  splitMeasure,
  capitalize,
  calculateRatio,
  arrayToNodeHTMLElement,
  imageBase64ToBuffer,
  convertCssToDocxMeasurement,
  random,
  parseCssMargin,
  getListItemNumber,
}
