import React, { useState } from 'react'
import YAML from 'js-yaml'
import JSZip from 'jszip'
import { saveAs } from 'file-saver'
import ExcelJS from 'exceljs'

const SNMP_EXPORTERS = ['exporter_cms','exporter_acm','exporter_aes','exporter_gateway','exporter_aaep','exporter_avayasbc',
  'exporter_audiocodesbc']

function normalizeVersion(v) {
  if (!v) return ''
  const s = String(v).trim().toLowerCase()
  if (s === '2' || s === 'v2' || s === 'snmpv2') return '2c'
  if (s === '2c' || s === '3' || s === 'v3' || s === 'snmpv3') return (s === 'v3' || s === 'snmpv3') ? '3' : s
  const m = s.match(/(2c|3|2)\b/)
  if (!m) return ''
  return m[1] === '2' ? '2c' : m[1]
}

function parseTrapVersion(trapArgs) {
  if (!trapArgs || typeof trapArgs !== 'string') return ''
  const m = trapArgs.match(/--snmp\.version\s+(\S+)/)
  return m ? normalizeVersion(m[1]) : ''
}

function parseV3Security(args) {
  if (!args) return ''
  const s = String(args)
  const authProto = (s.match(/--snmp\.auth-protocol\s+(\S+)/)?.[1] || '').toLowerCase()
  const privProto = (s.match(/--snmp\.privacy-protocol\s+(\S+)/)?.[1] || '').toLowerCase()
  const hasSecret = !!s.match(/--snmp\.(auth|privacy)-(password|passphrase)\s+(\S+)/)
  const hasUser = !!s.match(/--snmp\.username\s+(\S+)/)
  const isNone = (v) => !v || v === 'none'
  if (isNone(authProto) && isNone(privProto)) return 'NoAuthNoPriv'
  if (!isNone(authProto) && !isNone(privProto)) return 'AuthPriv'
  if (!isNone(authProto) && isNone(privProto)) return 'AuthNoPriv'
  if (isNone(authProto) && !isNone(privProto)) return 'AuthPriv'
  if (hasSecret || hasUser) return 'AuthPriv'
  return ''
}

function cleanCell(v) {
  if (v == null) return ''
  let s = String(v)
  s = s.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim()
  return s
}

function normalizeEntries(block) {
  if (!block) return []
  if (typeof block === 'string') return [{ hostname: block }]
  if (Array.isArray(block)) return block.map(v => (typeof v === 'string' ? { hostname: v } : { hostname: JSON.stringify(v) }))
  if (typeof block === 'object') {
    return Object.entries(block).map(([hostname, cfg]) => {
      let ip = ''
      let snmp_version = ''
      let community = ''
      let username = ''
      let extra_args = ''
      if (cfg && typeof cfg === 'object') {
        ip = cfg.ip_address || cfg.ip || ''
        snmp_version = normalizeVersion(cfg.snmp_version || cfg.snmp || '')
        community = cfg.community || cfg.snmp_community || ''
        username = cfg.username || cfg.snmp_username || ''
        if (typeof cfg.extra_args === 'string') {
          extra_args = cfg.extra_args
          const m1 = cfg.extra_args.match(/--snmp\.version\s+(\S+)/)
          if (m1) snmp_version = normalizeVersion(m1[1])
          const m3 = cfg.extra_args.match(/--snmp\.username\s+(\S+)/)
          if (m3 && !username) username = m3[1]
        }
      }
      return { hostname, ip_address: ip, snmp_version, community, username, extra_args }
    })
  }
  return []
}

function toCSVRowsForFile(filename, doc) {
  const customer = doc?.customer_name || filename.replace(/\.eyaml$/i, '')
  const globalTrapArgs = doc?.trap_extra_args || ''
  const globalTrapVer = normalizeVersion(parseTrapVersion(globalTrapArgs))

  const rows = []
  rows.push([`Customer: ${customer}`])
  rows.push(['File:', filename])
  rows.push(['Trap Extra Args:', globalTrapArgs])
  rows.push(['Derived SNMP Version (from trap args):', globalTrapVer])
  if (globalTrapVer === '3') {
    const sec = parseV3Security(globalTrapArgs)
    rows.push(['Derived V3 Security (from trap args):', sec])
  }
  rows.push([])

  SNMP_EXPORTERS.forEach((exp) => {
    const entries = normalizeEntries(doc?.[exp])
    rows.push([exp])
    rows.push(['Hostname','IP Address','SNMP Version','Security (V3)','Username','Trap Extra Args (V3 only)'])
    if (entries.length === 0) {
      rows.push(['(none found)','','','','',''])
    } else {
      entries.forEach(e => {
        const effectiveVersion = normalizeVersion(e.snmp_version || globalTrapVer || '')
        const effectiveArgs = effectiveVersion === '3' ? (e.extra_args || globalTrapArgs) : ''
        const security = effectiveVersion === '3' ? parseV3Security(effectiveArgs) : ''
        const username = effectiveVersion === '3' ? (e.username || '') : ''
        rows.push([
          cleanCell(e.hostname || ''),
          cleanCell(e.ip_address || ''),
          cleanCell(effectiveVersion || ''),
          cleanCell(security || ''),
          cleanCell(username || ''),
          cleanCell(effectiveArgs || '')
        ])
      })
    }
    rows.push([])
  })
  return rows
}

function rowsToCSV(rows) {
  return rows.map(r => r.map(cell => {
    const s = cleanCell(cell ?? '')
    if (s.includes('\"') || s.includes(',') || s.includes('\\n')) {
      return '\"' + s.replace(/\"/g, '\"\"') + '\"'
    }
    return s
  }).join(',')).join('\\n')
}

export default function App() {
  const [fileSummaries, setFileSummaries] = useState([])
  const [active, setActive] = useState(0)

  const handleFiles = async (files) => {
    const summaries = []
    for (const file of files) {
      try {
        const text = await file.text()
        const doc = YAML.load(text)
        const rows = toCSVRowsForFile(file.name, doc || {})
        summaries.push({ name: file.name, rows })
      } catch (e) {
        summaries.push({ name: file.name, rows: [], error: String(e) })
      }
    }
    setFileSummaries(summaries)
    setActive(0)
  }

  const labelFor = (sum) => {
    try {
      const first = sum.rows?.[0]?.[0] || ''
      const m = /^Customer:\s*(.*)$/.exec(first)
      if (m) return m[1].trim() || sum.name
    } catch {}
    return sum.name
  }

  const downloadAllCSVs = async () => {
    if (!fileSummaries.length) return
    const zip = new JSZip()
    const root = zip.folder('csv_exports')
    fileSummaries.forEach(sum => {
      const csv = rowsToCSV(sum.rows)
      let customer = ''
      try {
        const first = sum.rows?.[0]?.[0] || ''
        const m = /^Customer:\s*(.*)$/.exec(first)
        if (m) customer = m[1].trim()
      } catch {}
      const folderName = (customer || sum.name.replace(/\.eyaml$/i,'')).replace(/[^\w\-.]+/g, '_')
      const folder = root.folder(folderName)
      const csvName = (customer ? customer.replace(/[^\w\-.]+/g, '_') : sum.name.replace(/\.eyaml$/i,'')) + '.csv'
      folder.file(csvName, csv)
    })
    const blob = await zip.generateAsync({type:'blob'})
    saveAs(blob, 'eyaml-snmp-csvs.zip')
  }

  const downloadAllXLSX = async () => {
    if (!fileSummaries.length) return
    const zip = new JSZip();
    const root = zip.folder('excel_exports');
    for (const sum of fileSummaries) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('SNMP Summary', {properties:{defaultRowHeight:14}});
      const rows = sum.rows.map(r => r.map(c => cleanCell(c)));
      rows.forEach((r) => {
        const row = ws.addRow(r);
        row.height = 14;
        row.alignment = { vertical: 'middle', wrapText: false };
        if (r.length >= 3) {
          const v = String(r[2]).toLowerCase();
          const cell = row.getCell(3);
          if (v === '3' || v === 'v3' || v === 'snmpv3') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34EBC0' } };
          } else if (v === '2c' || v === '2') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBA134' } };
          }
        }
      });
      ws.eachRow((row) => {
        row.eachCell((cell) => {
          cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        });
      });
      const widths = [36, 18, 10, 16, 22, 54];
      widths.forEach((w,i)=> ws.getColumn(i+1).width = w);
      const buffer = await wb.xlsx.writeBuffer();
      let customer = '';
      try {
        const first = rows?.[0]?.[0] || '';
        const m = /^Customer:\s*(.*)$/.exec(first);
        if (m) customer = m[1].trim();
      } catch {}
      const fnameBase = (customer || sum.name.replace(/\.eyaml$/i,'')).replace(/[^\w\-.]+/g, '_');
      const fileName = fnameBase + '.xlsx';
      root.file(fileName, new Blob([buffer]));
    }
    const blob = await zip.generateAsync({type:'blob'});
    saveAs(blob, 'eyaml-snmp-xlsx.zip');
  }

  const downloadOneCSV = async (sum) => {
    const csv = rowsToCSV(sum.rows)
    let customer = ''
    try {
      const first = sum.rows?.[0]?.[0] || ''
      const m = /^Customer:\s*(.*)$/.exec(first)
      if (m) customer = m[1].trim()
    } catch {}
    const csvName = (customer || sum.name.replace(/\.eyaml$/i,'')) + '.csv'
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8'})
    saveAs(blob, csvName.replace(/[^\w\-.]+/g,'_'))
  }

  const downloadOneXLSX = async (sum) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('SNMP Summary', {properties:{defaultRowHeight:14}});
    const rows = sum.rows.map(r => r.map(c => cleanCell(c)));
    rows.forEach((r) => {
      const row = ws.addRow(r);
      row.height = 14;
      row.alignment = { vertical: 'middle', wrapText: false };
      if (r.length >= 3) {
        const v = String(r[2]).toLowerCase();
        const cell = row.getCell(3);
        if (v === '3' || v === 'v3' || v === 'snmpv3') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34EBC0' } };
        } else if (v === '2c' || v === '2') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBA134' } };
        }
      }
    });
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
      });
    });
    const widths = [36, 18, 10, 16, 22, 54];
    widths.forEach((w,i)=> ws.getColumn(i+1).width = w);
    const buffer = await wb.xlsx.writeBuffer();
    let customer = '';
    try {
      const first = rows?.[0]?.[0] || '';
      const m = /^Customer:\s*(.*)$/.exec(first);
      if (m) customer = m[1].trim();
    } catch {}
    const fnameBase = (customer || sum.name.replace(/\.eyaml$/i,'')).replace(/[^\w\-.]+/g, '_');
    const fileName = fnameBase + '.xlsx';
    saveAs(new Blob([buffer]), fileName);
  }

  const activeSum = fileSummaries[active]

  return (
    <div style={{fontFamily:'system-ui,Segoe UI,Roboto,Arial', padding:'2rem', maxWidth:'1200px', margin:'0 auto'}}>
      <h1>EYAML SNMP Helper</h1>
      <p><strong>Private by design:</strong> All parsing happens in your browser. No files are uploaded.</p>

      <input type="file" accept=".yaml,.yml,.eyaml,.txt" multiple onChange={(e)=>e.target.files && handleFiles(e.target.files)} />

      {fileSummaries.length > 0 && (
        <>
          <div style={{marginTop:'1rem'}}>
            <button onClick={downloadAllCSVs}>Download CSVs (ALL)</button>
            <span style={{display:'inline-block', width:'12px'}}></span>
            <button onClick={downloadAllXLSX}>Download Excel (ALL, styled)</button>
          </div>

          <div style={{display:'flex', gap:'8px', marginTop:'16px', flexWrap:'wrap'}}>
            {fileSummaries.map((sum, i) => (
              <button
                key={i}
                onClick={() => setActive(i)}
                style={{
                  padding:'8px 12px',
                  border:'1px solid #ddd',
                  borderBottom: active === i ? '2px solid #333' : '1px solid #ddd',
                  background: active === i ? '#f2f2f2' : 'white',
                  borderRadius:'8px',
                  cursor:'pointer'
                }}
                title={sum.name}
              >
                {labelFor(sum)}
              </button>
            ))}
          </div>

          {activeSum && (
            <div style={{marginTop:'12px', padding:'12px', border:'1px solid #ddd', borderRadius:'8px'}}>
              <div style={{display:'flex', gap:'8px', marginBottom:'10px'}}>
                <button onClick={() => downloadOneCSV(activeSum)}>Download CSV (this tab)</button>
                <button onClick={() => downloadOneXLSX(activeSum)}>Download Excel (this tab, styled)</button>
              </div>
              {activeSum.error ? (
                <p style={{color:'crimson'}}>Error: {activeSum.error}</p>
              ) : (
                <div style={{overflowX:'auto'}}>
                  <table style={{borderCollapse:'collapse', minWidth:'980px'}}>
                    <tbody>
                      {activeSum.rows.map((r,i) => (
                        <tr key={i}>
                          {r.length === 0 ? (
                            <td colSpan={6} style={{height:'10px'}}></td>
                          ) : r.length === 1 ? (
                            <td colSpan={6} style={{fontWeight:'600', background:'#f7f7f7', padding:'6px 8px', border:'1px solid #eee'}}>{r[0]}</td>
                          ) : (
                            r.map((c,j) => (
                              <td key={j} style={{border:'1px solid #eee', padding:'6px 8px', whiteSpace:'pre'}}>{cleanCell(c)}</td>
                            ))
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </>
      )}
    </div>
  )
}
