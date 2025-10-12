SNMP Helper
Client-only EYAML → CSV/XLSX generator for SNMP exporters.

- V2 → 2c normalisation, V3 security detection (NoAuthNoPriv/AuthNoPriv/AuthPriv)
- Username column for V3; trap args ignored for V2c rows
- Styled Excel export: version-colored fills (#34ebc0 for v3, #eba134 for 2c)
- **Tabbed UI**: one tab per uploaded file/customer; per-tab and all-in-one exports.

Run
npm install
npm run dev
