import {Workbook} from './workbook.js';
import {freshTools,validateTools,type SheetTools} from './sheettools.js';
export interface SheetEntry {id:string;name:string;book:Workbook;tools:SheetTools}
export class SpreadsheetDocument {
  private items:SheetEntry[]=[];
  private nextId=1;
  activeId='';
  constructor(initial?:Workbook) {this.addSheet('Sheet 1',initial??new Workbook());}
  get sheets():readonly SheetEntry[]{return this.items;}
  get active():SheetEntry{return this.items.find(x=>x.id===this.activeId)??this.items[0]!;}
  private validateName(name:string,excludeId?:string){
    const n=name.trim();if(!n||n.length>31||/[\\/?*\[\]:]/.test(n))throw Error('Sheet name must be 1–31 characters and cannot contain \\ / ? * [ ] :.');
    if(this.items.some(s=>s.id!==excludeId&&s.name.toLowerCase()===n.toLowerCase()))throw Error('A sheet with that name already exists.');
    return n;
  }
  addSheet(name?:string,book?:Workbook):SheetEntry {
    if(this.items.length>=12)throw Error('This version supports a maximum of 12 sheets.');
    let title=name;if(!title){let n=1;while(this.items.some(x=>x.name.toLowerCase()===`sheet ${n}`))n++;title=`Sheet ${n}`;}
    title=this.validateName(title);
    const item={id:`s${this.nextId++}`,name:title,book:book??new Workbook(),tools:freshTools()};this.items.push(item);this.activeId=item.id;return item;
  }
  selectSheet(id:string){if(!this.items.some(x=>x.id===id))throw Error('Sheet not found.');this.activeId=id;}
  renameSheet(id:string,name:string){const found=this.items.find(x=>x.id===id);if(!found)throw Error('Sheet not found.');found.name=this.validateName(name,id);}
  duplicateSheet(id=this.activeId){const original=this.items.find(x=>x.id===id);if(!original)throw Error('Sheet not found.');
    let name=`${original.name} copy`,suffix=2;while(this.items.some(x=>x.name.toLowerCase()===name.toLowerCase()))name=`${original.name.slice(0,23)} ${suffix++}`;
    const duplicate=this.addSheet(name,Workbook.fromJSON(original.book.toJSON()));duplicate.tools=validateTools(JSON.parse(JSON.stringify(original.tools)),duplicate.book.rows,duplicate.book.cols);return duplicate;
  }
  deleteSheet(id=this.activeId){if(this.items.length===1)throw Error('At least one sheet must remain.');const idx=this.items.findIndex(x=>x.id===id);if(idx===-1)throw Error('Sheet not found.');this.items.splice(idx,1);if(this.activeId===id)this.activeId=this.items[Math.min(idx,this.items.length-1)]!.id;}
  toJSON(){return {format:'gridforge',version:3,activeId:this.activeId,sheets:this.items.map(x=>({id:x.id,name:x.name,workbook:x.book.toJSON(),tools:x.tools}))};}
  static fromJSON(raw:unknown):SpreadsheetDocument {
    if(!raw||typeof raw!=='object')throw Error('Invalid document JSON.');const data=raw as Record<string,unknown>;
    if(!Array.isArray(data.sheets)){return new SpreadsheetDocument(Workbook.fromJSON(raw));}
    if(data.format!=='gridforge'||![2,3].includes(data.version as number)||!data.sheets.length||data.sheets.length>12)throw Error('Invalid GridForge v2 document.');
    const doc=new SpreadsheetDocument();doc.items=[];doc.nextId=1;
    for(const s of data.sheets){if(!s||typeof s!=='object')throw Error('Invalid sheet.');const x=s as Record<string,unknown>;
      if(typeof x.name!=='string')throw Error('Invalid sheet name.');const name=doc.validateName(x.name);
      const id=`s${doc.nextId++}`;const book=Workbook.fromJSON(x.workbook);doc.items.push({id,name,book,tools:validateTools(x.tools,book.rows,book.cols)});
    }
    if(typeof data.activeId==='string'){
      const i=data.sheets.findIndex((x:unknown)=>!!x&&typeof x==='object'&&(x as Record<string,unknown>).id===data.activeId);
      doc.activeId=doc.items[Math.max(0,i)]!.id;
    }else doc.activeId=doc.items[0]!.id;
    return doc;
  }
}
