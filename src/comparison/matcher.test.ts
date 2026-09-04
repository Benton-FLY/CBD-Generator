import {describe,expect,it} from 'vitest';
import {assignMaterialMatches,matchMaterials,matchStyles,scoreMaterialCandidate,validateMaterialMatchIntegrity} from './matcher';
import {extractMaterialAttributes,normalizeMaterial,normalizeMaterialName,normalizeStyle} from './normalize';
import type {CbdMaterialRow,CbdStyle,MaterialMatchCluster} from './types';

let sequence=0;
const row=(material:string,options:Partial<CbdMaterialRow>={}):CbdMaterialRow=>({id:options.id||`row-${sequence++}`,group:'TRIMS',material,size:'',unit:'YD',cost:1,usage:1,loss:.05,extended:1.05,remark:'',width:'',order:sequence,...options});
const style=(id:string,side:'reference'|'current',name:string,items:CbdMaterialRow[]=[]):CbdStyle=>({id,side,fileName:`${id}.xlsx`,sheetName:name,styleName:name,summary:{},groupOrder:[],materials:items});
const pair=(reference:CbdMaterialRow,current:CbdMaterialRow)=>assignMaterialMatches([reference],[current])[0];

describe('material normalization and attributes',()=>{
 it('normalizes punctuation and retains identity-bearing codes and dimensions',()=>{expect(normalizeStyle('27 RAYCE JERSEY CBD SHEET')).toBe('RAYCE JERSEY');expect(normalizeMaterial('FLY RACING 26 JERSEY+JX-306',true)).toContain('JX-306');expect(normalizeMaterialName('NYLON WEBBING (25M/M)')).toContain('25MM');expect(extractMaterialAttributes('VT_SIZE WOVEN LABEL (#AF132 A01)').codes).toContain('#AF132 A01')});
 it.each([
  ['VT_SIZE WOVEN LABEL FLY RACING 20 F-16 (#AF132 A01)','SIZE WOVEN LABEL FLY RACING 20 F-16 (#AF132 A01)','CODE MATCH'],
  ['MAIN BUCKLE FLY RACING','FLY RACING MAIN BUCKLE SET','TOKEN MATCH'],
  ['HANG TAG FLY RACING 26 F-16 PANT (3 PAGES & 1 EYELET)','HANG TAG FLY RACING 25 F-16 PANT (3 PAGES & 1 EYELET) PFAS FREE (JX)','TOKEN MATCH'],
  ['PLASTIC SQUARE BUCKLE 30MM','PLASTIC SQUARE BUCKLE 30MM - INNER/OUTER WIDTH 30/38MM','NORMALIZED'],
  ['NYLON WEBBING (25M/M)','NEW - NYLON WEBBING (25M/M) P (HANSHIN)','NORMALIZED'],
 ])('automatically matches %s', (a,b,status)=>expect(pair(row(a),row(b))).toMatchObject({referenceRowIds:expect.any(Array),currentRowId:expect.any(String),status}));
});

describe('false-positive protection',()=>{
 it.each([
  ['WEBBING 10MM','WEBBING 25MM'],['VELCRO A','VELCRO B'],['ZIPPER LEFT','ZIPPER RIGHT'],['PATCH FRONT','PATCH BACK'],['LABEL #AF132 A01','LABEL #AF999 B02'],['PRINT MATERIAL','PRINT MATERIAL'],
 ])('does not auto-link conflicting or generic %s / %s',(a,b)=>{const left=row(a),right=row(b);if(a==='PRINT MATERIAL')right.id+='-other';const scored=scoreMaterialCandidate(left,right);if(a==='PRINT MATERIAL')expect(assignMaterialMatches([left,row(a)],[right]).some(c=>c.currentRowId===right.id&&c.referenceRowIds.length)).toBe(false);else expect(scored.score).toBeLessThan(90)});
 it('requires compatible units',()=>expect(scoreMaterialCandidate(row('MAIN BUCKLE',{unit:'PCS'}),row('MAIN BUCKLE',{unit:'YD'})).score).toBe(0));
});

describe('global assignment, locks, and N:1',()=>{
 it('is independent of input order and never duplicates a comparison row',()=>{const refs=[row('MAIN BUCKLE FLY RACING',{id:'r1'}),row('NYLON WEBBING 25MM',{id:'r2'})],curr=[row('NEW NYLON WEBBING 25MM',{id:'c2'}),row('FLY RACING MAIN BUCKLE SET',{id:'c1'})],signature=(clusters:MaterialMatchCluster[])=>clusters.filter(c=>c.currentRowId&&c.referenceRowIds.length).map(c=>`${c.referenceRowIds[0]}:${c.currentRowId}`).sort();expect(signature(assignMaterialMatches(refs,curr))).toEqual(signature(assignMaterialMatches([...refs].reverse(),[...curr].reverse())));const ids=assignMaterialMatches(refs,curr).map(c=>c.currentRowId).filter(Boolean);expect(new Set(ids).size).toBe(ids.length)});
 it('keeps close competing candidates in review',()=>{const ref=row('PLASTIC SQUARE BUCKLE 30MM',{id:'r'}),a=row('PLASTIC SQUARE BUCKLE 30MM INNER',{id:'a'}),b=row('PLASTIC SQUARE BUCKLE 30MM OUTER',{id:'b'}),clusters=assignMaterialMatches([ref],[a,b]);expect(clusters.some(c=>c.status==='REVIEW')).toBe(true);expect(clusters.some(c=>c.referenceRowIds.includes('r')&&c.currentRowId)).toBe(false)});
 it('preserves manual locks during re-auto matching',()=>{const ref=row('MAIN BUCKLE',{id:'r'}),current=row('MAIN BUCKLE',{id:'c'}),locked:MaterialMatchCluster={id:'locked',referenceRowIds:['r'],currentRowId:'c',relationType:'one-to-one',matchSource:'manual',manualLocked:true,finalGroup:'TRIMS',status:'MANUAL',confidence:1};expect(assignMaterialMatches([ref],[current],[locked])).toEqual([locked])});
 it('merges equivalent reference rows only when totals reconcile and counts comparison money once',()=>{const refs=[row('WEBBING POLY 10MM',{id:'r1',usage:1,extended:2}),row('WEBBING POLY 10MM',{id:'r2',usage:2,extended:4})],current=row('WEBBING POLY 10MM',{id:'c',usage:3,extended:6}),clusters=assignMaterialMatches(refs,[current]),merged=clusters.find(c=>c.status==='MERGED N:1');expect(merged?.referenceRowIds).toEqual(['r1','r2']);expect(validateMaterialMatchIntegrity(refs,[current],clusters)).toMatchObject({valid:true,referenceTotal:6,comparisonTotal:6,coveredComparisonTotal:6})});
 it('does not merge repeated generic names without quantitative evidence',()=>{const refs=[row('PRINT MATERIAL',{id:'r1',usage:1,extended:2}),row('PRINT MATERIAL',{id:'r2',usage:2,extended:4})],current=row('PRINT MATERIAL',{id:'c',usage:9,extended:20});expect(assignMaterialMatches(refs,[current]).some(c=>c.status==='MERGED N:1')).toBe(false)});
});

describe('style and regression behavior',()=>{
 it('matches style aliases but not different identifiers',()=>{const result=matchStyles([style('r1','reference','27 RAYCE JERSEY'),style('r2','reference','27 F-16 JERSEY (3XL ~ 5XL)'),style('r3','reference','27 KINETIC YOUTH 3 JERSEY'),style('c1','current','RAYCE JERSEY'),style('c2','current','F-16 JERSEY (BIG SIZES)'),style('c3','current','KINETIC YOUTH 1 JERSEY')]);expect(result.find(x=>x.referenceId==='r1')?.currentId).toBe('c1');expect(result.find(x=>x.referenceId==='r2')?.currentId).toBe('c2');expect(result.find(x=>x.referenceId==='r3')?.currentId).toBeUndefined()});
 it('uses the comparison group and reports group changes',()=>{const reference=style('r','reference','R',[row('LABEL #AF132 A01',{id:'r0',group:'TRIMS'})]),current=style('c','current','C',[row('LABEL #AF132 A01',{id:'c0',group:'LABEL & PACKAGING'})]),set=matchMaterials({id:'m',referenceId:'r',currentId:'c',method:'Normalized',confidence:1,status:'Normalized'},[reference,current]);expect(set.clusters[0]).toMatchObject({finalGroup:'LABEL & PACKAGING',status:'GROUP CHANGED'})});
});
