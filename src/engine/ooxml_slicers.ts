const SLICER_NS='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const DRAWING_SLICER_NS='http://schemas.microsoft.com/office/drawing/2010/slicer';
const MC_NS='http://schemas.openxmlformats.org/markup-compatibility/2006';
const XDR_NS='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS='http://schemas.openxmlformats.org/drawingml/2006/main';
const escape=(value:string)=>value.replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[char]!));

export interface NativeSlicerCacheSpec {
  cacheName:string;
  sourceName:string;
  pivotCacheId:number;
  pivotTables:Array<{tabId:number;name:string}>;
  items:string[];
  selectedItems?:string[]|null;
}

export interface NativeSlicerViewSpec {
  name:string;
  cacheName:string;
  caption:string;
  width:number;
  height:number;
  x:number;
  y:number;
}

/** Slicer cache part for a non-OLAP worksheet PivotCache, as defined by MS-XLSX. */
export function nativeSlicerCacheXml(spec:NativeSlicerCacheSpec):string {
  const selected=spec.selectedItems===null||spec.selectedItems===undefined?null:new Set(spec.selectedItems.map(value=>value.toLocaleLowerCase()));
  const items=spec.items.map((value,index)=>`<i x="${index}"${selected===null||selected.has(value.toLocaleLowerCase())?' s="1"':''}/>`).join('');
  const pivotTables=spec.pivotTables.map(pivot=>`<pivotTable tabId="${pivot.tabId}" name="${escape(pivot.name)}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicerCacheDefinition xmlns="${SLICER_NS}" xmlns:mc="${MC_NS}" name="${escape(spec.cacheName)}" sourceName="${escape(spec.sourceName)}"><pivotTables>${pivotTables}</pivotTables><data><tabular pivotCacheId="${spec.pivotCacheId}"><items count="${spec.items.length}">${items}</items></tabular></data></slicerCacheDefinition>`;
}

/** Slicer view part associated with a worksheet drawing shape by the same name. */
export function nativeSlicerViewXml(views:NativeSlicerViewSpec[]):string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicers xmlns="${SLICER_NS}">${views.map(view=>`<slicer name="${escape(view.name)}" cache="${escape(view.cacheName)}" caption="${escape(view.caption)}" columnCount="1" showCaption="1" style="SlicerStyleLight1" rowHeight="285750"/>`).join('')}</slicers>`;
}

/** Worksheet reference to its slicer part, placed in the specified worksheet extension. */
export function nativeSlicerListExtension(relationshipIds:string|string[]):string {
  const ids=Array.isArray(relationshipIds)?relationshipIds:[relationshipIds];
  return `<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" xmlns:x14="${SLICER_NS}"><x14:slicerList xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${ids.map(id=>`<x14:slicer r:id="${escape(id)}"/>`).join('')}</x14:slicerList></ext>`;
}

/** Office drawing anchor for a slicer view; Excel uses the choice and other readers get a shape fallback. */
export function nativeSlicerAnchorXml(view:NativeSlicerViewSpec,shapeId:number):string {
  const x=Math.round(view.x*9525),y=Math.round(view.y*9525),width=Math.round(view.width*9525),height=Math.round(view.height*9525),caption=escape(view.caption);
  return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${width}" cy="${height}"/><mc:AlternateContent><mc:Choice Requires="sle"><sle:slicer name="${escape(view.name)}"/></mc:Choice><mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${shapeId}" name="${caption}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${caption}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:absoluteAnchor>`;
}
