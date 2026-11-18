import JSZip from 'jszip'

const WIDTHS = [36, 18, 10, 16, 22, 54]
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

function escapeText(value) {
  const text = String(value ?? '').toWellFormed()
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
  if (text.length > 32767) throw new RangeError('An Excel cell cannot exceed 32,767 characters')
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

// The product only writes six-column summaries. Inline strings keep user data
// out of formulas, relationships, filenames, and XML attribute contexts.
export async function createSummaryXlsx(rows) {
  if (!Array.isArray(rows) || rows.length > 1048576) {
    throw new RangeError('An Excel worksheet cannot exceed 1,048,576 rows')
  }
  const sheetRows = rows.map((row, index) => {
    if (!Array.isArray(row) || row.length > WIDTHS.length) {
      throw new RangeError('A summary row cannot exceed six columns')
    }
    const number = index + 1
    const cells = row.map((value, column) => {
      const version = String(value ?? '').toLowerCase()
      const style = column === 2 && ['3', 'v3', 'snmpv3'].includes(version) ? 2
        : column === 2 && ['2', '2c'].includes(version) ? 3 : 1
      const ref = `${String.fromCharCode(65 + column)}${number}`
      return `<c r="${ref}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${escapeText(value)}</t></is></c>`
    }).join('')
    return `<row r="${number}" ht="14" customHeight="1">${cells}</row>`
  }).join('')
  const columns = WIDTHS.map((width, index) =>
    `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join('')
  const zip = new JSZip()
  zip.file('[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
  zip.file('_rels/.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
  zip.file('xl/workbook.xml', XML + `<workbook xmlns="${MAIN}" xmlns:r="${REL}">`
    + '<sheets><sheet name="SNMP Summary" sheetId="1" r:id="rId1"/></sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
    + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`)
  zip.file('xl/styles.xml', XML + `<styleSheet xmlns="${MAIN}">`
    + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FF34EBC0"/><bgColor indexed="64"/></patternFill></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFEBA134"/><bgColor indexed="64"/></patternFill></fill></fills>'
    + '<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="4">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + [0, 2, 3].map(fill => `<xf numFmtId="0" fontId="0" fillId="${fill}" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="0"/></xf>`).join('')
    + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
  zip.file('xl/worksheets/sheet1.xml', XML + `<worksheet xmlns="${MAIN}"><sheetFormatPr defaultRowHeight="14"/>`
    + `<cols>${columns}</cols><sheetData>${sheetRows}</sheetData></worksheet>`)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
