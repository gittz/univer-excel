// Excel 文件里，2010 以后新增的函数必须带 _xlfn. 前缀（FILTER/SORT 还要 _xlws.），
// LET / LAMBDA 的参数名要带 _xlpm. 前缀，否则 Excel 打开会显示 #NAME?。
// 这里负责在 Univer 公式和 Excel 文件公式之间加 / 去这些前缀。

const XLWS = new Set(['FILTER', 'SORT'])

const XLFN = new Set(`
ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT
BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL
DROP ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML
FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR
FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH
IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP
MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV
NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST
QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P
SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT
TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS
XLOOKUP XMATCH XOR Z.TEST GROUPBY PIVOTBY PERCENTOF REGEXTEST REGEXEXTRACT REGEXREPLACE TRIMRANGE
`.trim().split(/\s+/))

// 把公式切成：字符串字面量 / 带引号的工作表名 / 其他片段，只在“其他片段”里做替换
function splitProtected(formula) {
  const parts = []
  let i = 0
  let buf = ''
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"' || ch === "'") {
      if (buf) { parts.push({ code: true, text: buf }); buf = '' }
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) { j += 2; continue }
          break
        }
        j++
      }
      parts.push({ code: false, text: formula.slice(i, j + 1) })
      i = j + 1
      continue
    }
    buf += ch
    i++
  }
  if (buf) parts.push({ code: true, text: buf })
  return parts
}

const IDENT = /(^|[^A-Za-z0-9_.\\$!:])([A-Za-z_\\][A-Za-z0-9_.]*)(?=\s*(\()?)/g

// 找出 LET / LAMBDA 声明的参数名
function collectLambdaParams(code) {
  const params = new Set()
  const re = /\b(LET|LAMBDA)\s*\(/gi
  let m
  while ((m = re.exec(code))) {
    const fn = m[1].toUpperCase()
    // 解析这一层括号内的顶层参数
    let depth = 0
    let start = m.index + m[0].length
    const args = []
    for (let i = start; i < code.length; i++) {
      const ch = code[i]
      if (ch === '(' || ch === '{') depth++
      else if (ch === ')' || ch === '}') {
        if (depth === 0) { args.push(code.slice(start, i)); break }
        depth--
      } else if (ch === ',' && depth === 0) { args.push(code.slice(start, i)); start = i + 1 }
    }
    const names = fn === 'LET' ? args.filter((_, idx) => idx % 2 === 0 && idx < args.length - 1) : args.slice(0, -1)
    for (const n of names) {
      const name = n.trim().replace(/^_xlpm\./i, '')
      if (/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) params.add(name.toUpperCase())
    }
  }
  return params
}

// Univer 公式（'=XLOOKUP(...)'）-> Excel 文件公式（'_xlfn.XLOOKUP(...)'，不带等号）
export function toExcelFormula(formula) {
  const body = String(formula).replace(/^=/, '')
  const parts = splitProtected(body)
  const params = collectLambdaParams(parts.filter((p) => p.code).map((p) => p.text).join(' '))
  return parts.map((p) => {
    if (!p.code) return p.text
    return p.text.replace(IDENT, (all, pre, name, paren) => {
      const upper = name.toUpperCase()
      if (params.has(upper)) return `${pre}_xlpm.${name}`
      if (!paren) return all
      if (XLWS.has(upper)) return `${pre}_xlfn._xlws.${name}`
      if (XLFN.has(upper)) return `${pre}_xlfn.${name}`
      return all
    })
  }).join('')
}

// Excel 文件公式 -> Univer 公式（带等号，去掉所有 _xlfn./_xlws./_xlpm. 前缀）
export function fromExcelFormula(formula) {
  const parts = splitProtected(String(formula).replace(/^=/, ''))
  return '=' + parts.map((p) => (p.code ? p.text.replace(/_xlfn\.|_xlws\.|_xlpm\./gi, '') : p.text)).join('')
}
