import {readFile,writeFile} from 'node:fs/promises';
import {auditComparison} from '../src/comparison/audit';
const path=process.argv[2];if(!path)throw new Error('Usage: vite-node scripts/audit-comparison.mts saved-state.json [report.json]');
await writeFile(process.argv[3]||'reports/all-style-audit.json',JSON.stringify(auditComparison(JSON.parse(await readFile(path,'utf8'))),null,2));
