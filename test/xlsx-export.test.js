import assert from 'node:assert/strict'
import { test } from 'node:test'
import JSZip from 'jszip'
import { createSummaryXlsx } from '../src/xlsx-export.js'

test('exports styled SNMP summaries with fixed worksheet and package paths', async () => {
  const bytes = await createSummaryXlsx([
    ['Customer: Example'], [],
    ['host-1', '192.0.2.1', '3', 'AuthPriv', 'user', '--snmp.version 3'],
    ['host-2', '192.0.2.2', '2c'],
  ])
  const zip = await JSZip.loadAsync(bytes)
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string')
  assert.match(sheet, /r="C3" t="inlineStr" s="2"/)
  assert.match(sheet, /r="C4" t="inlineStr" s="3"/)
  assert.match(sheet, /width="54"/)
  assert.match(sheet, /<t xml:space="preserve">AuthPriv<\/t>/)
  assert.ok(zip.file('[Content_Types].xml'))
  assert.ok(zip.file('xl/_rels/workbook.xml.rels'))
})

test('untrusted labels remain text, including formulas and XML injection attempts', async () => {
  const zip = await JSZip.loadAsync(await createSummaryXlsx([
    ['=HYPERLINK("https://example.invalid", "click")', '</t></is></c><f>1+1</f>', '3'],
    ['A&B <customer>', 'bad\u0000control\ud800'],
  ]))
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string')
  assert.doesNotMatch(sheet, /<f(?:\s|>)/)
  assert.match(sheet, /&lt;\/t&gt;&lt;\/is&gt;&lt;\/c&gt;&lt;f&gt;1\+1&lt;\/f&gt;/)
  assert.match(sheet, /A&amp;B &lt;customer&gt;/)
  assert.doesNotMatch(sheet, /\u0000/)
  assert.match(sheet, /badcontrol\uFFFD/)
  assert.deepEqual(Object.keys(zip.files).filter(name => !zip.files[name].dir).sort(), [
    '[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels',
    'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml',
  ].sort())
})

test('rejects values outside Excel cell, row, and column limits', async () => {
  await assert.rejects(createSummaryXlsx([['x'.repeat(32768)]]), /32,767/)
  await assert.rejects(createSummaryXlsx([Array(7).fill('x')]), /six columns/)
  await assert.rejects(createSummaryXlsx(new Array(1048577)), /1,048,576/)
})
