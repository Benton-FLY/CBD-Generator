import fs from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
import assert from 'node:assert/strict';
import {managerFixture} from './manager-fixture';
import {buildComparisonWorkbook} from '../src/comparison/exporter';
const dir='reports/.manager-baseline';await fs.mkdir(dir,{recursive:true});
for(const file of ['exporter.ts','groups.ts','types.ts'])await fs.writeFile(`${dir}/${file}`,execFileSync('git',['show',`61f9719:src/comparison/${file}`]));
const {buildComparisonWorkbook:baseline}=await import(`${process.cwd()}/${dir}/exporter.ts`);
const state=await managerFixture();
for(const options of [undefined,{stylePairIds:[state.styleMatches[0].id]}]){
 const old=baseline(state,options),now=buildComparisonWorkbook(state,options);
 assert.deepEqual(now.worksheets.map(s=>s.model),old.worksheets.map(s=>s.model));
}
console.log('PASS: existing Current/All worksheet values, formulas, formatting, names and print settings unchanged.');
for(const file of ['exporter.ts','groups.ts','types.ts'])await fs.unlink(`${dir}/${file}`);await fs.rmdir(dir);
