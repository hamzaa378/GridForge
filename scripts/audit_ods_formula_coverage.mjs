#!/usr/bin/env node
/** Measure which frozen Excel parity formulas ODS writes as portable ODF formulas. */
import fs from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportOds} from '../dist/engine/ods.js';
import {zipRead} from '../dist/engine/zip.js';

const cases=JSON.parse(fs.readFileSync(new URL('../tests/excel_parity_506_cases.json',import.meta.url),'utf8')).cases;
const rows=[];
for(const item of cases){
 // The translator only needs formula syntax; omit recipe cells so every ODS
 // package contains exactly one formula and the measured cell is unambiguous.
 const book=new Workbook(400,52);book.applyEdits([{ref:'Z20',raw:item.formula}],false);
 const document=new SpreadsheetDocument(book),archive=await zipRead(exportOds(document)),xml=archive.get('content.xml'),formulaCell=xml.match(/<table:table-cell\b[^>]*(?:\/>|>)/g)?.find(cell=>cell.includes('table:formula=')||cell.includes('gridforge:source-formula='));
 // Formula cells carry native table:formula attributes. If translation fails,
 // GridForge writes its source formula only as a private recovery extension.
 const exportedFormula=formulaCell?.includes('table:formula=')??false,libreOfficeSpecific=formulaCell?.includes('ORG.LIBREOFFICE.')??false,microsoftCompatibility=formulaCell?.includes('COM.MICROSOFT.')??false,recoveryMetadata=formulaCell?.includes('gridforge:source-formula=')??false;
 rows.push({function:item.function,formula:item.formula,openFormula:exportedFormula&&!libreOfficeSpecific&&!microsoftCompatibility,microsoftCompatibility,libreOfficeSpecific,gridForgeRecoveryOnly:recoveryMetadata&&!exportedFormula,recoveryMetadataPresent:recoveryMetadata});
}
const openFormula=rows.filter(row=>row.openFormula).length,microsoftCompatibility=rows.filter(row=>row.microsoftCompatibility).length,libreOfficeSpecific=rows.filter(row=>row.libreOfficeSpecific).length,privateOnly=rows.filter(row=>row.gridForgeRecoveryOnly).length,recoveryCopies=rows.filter(row=>row.recoveryMetadataPresent).length,summary={fixtureCount:rows.length,openFormula,microsoftCompatibility,libreOfficeSpecific,gridForgeRecoveryOnly:privateOnly,recoveryMetadataCopies:recoveryCopies,openFormulaPercent:Number((openFormula/rows.length*100).toFixed(2)),method:'One frozen desktop-Excel parity formula per case. Standard OpenFormula, Microsoft-namespace, LibreOffice-specific, and GridForge-only serialization are counted separately; none proves full calculation parity.'};
const report={schema:'gridforge-ods-formula-coverage/1',summary,formulas:rows};
fs.writeFileSync(new URL('../docs/ODS_FORMULA_COVERAGE.json',import.meta.url),JSON.stringify(report,null,2)+'\n');
const lines=['# Frozen formula fixture: ODS export coverage','',`Standard OpenFormula expressions: **${openFormula}/${rows.length} (${summary.openFormulaPercent}%)**.`,`Microsoft compatibility-namespace formulas: **${microsoftCompatibility}/${rows.length}**.`,`LibreOffice-specific formulas: **${libreOfficeSpecific}/${rows.length}**.`,`GridForge-only recovery formulas: **${privateOnly}/${rows.length}**; metadata recovery copies are present on ${recoveryCopies} formulas.`, '',summary.method,'','## LibreOffice-specific formulas with a GridForge recovery copy','',...rows.filter(row=>row.libreOfficeSpecific).map(row=>`- **${row.function}**: \`${row.formula}\``),'','## Microsoft compatibility-namespace formulas','',...rows.filter(row=>row.microsoftCompatibility).map(row=>`- **${row.function}**: \`${row.formula}\``),'','## Formulas stored only in GridForge recovery metadata','',...rows.filter(row=>row.gridForgeRecoveryOnly).map(row=>`- **${row.function}**: \`${row.formula}\``),''];
fs.writeFileSync(new URL('../docs/ODS_FORMULA_COVERAGE.md',import.meta.url),lines.join('\n'));
console.log(JSON.stringify(summary,null,2));
