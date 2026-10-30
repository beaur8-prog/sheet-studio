import {xml,D} from './chart-xml.mjs';
import {anchorPlacement,drawingAnchorRect,positionAnchor,transformAnchor} from './anchor-position.mjs';
const keys=['x','y','width','height'];
export function transformAnchoredObjects(sheet,before,operation){
 if(sheet.kind==='chart')return;
 for(const chart of sheet.charts||[]){
  const original=chart.source?.original,placement=chart.placement||(chart.source?anchorPlacement(xml(chart.source.anchor).documentElement):'fixed'),rect={x:chart.x??40,y:chart.y??25,width:chart.width??500,height:chart.height??300};
  let source=chart.source?.anchor||`<xdr:absoluteAnchor xmlns:xdr="${D}"/>`;
  if(!original?.geometry||!keys.every(k=>rect[k]===original.geometry[k])||placement!==(original.placement||anchorPlacement(xml(source).documentElement)))source=positionAnchor(source,rect,before,placement);
  const anchor=transformAnchor(source,before,sheet,operation),next=drawingAnchorRect(xml(anchor).documentElement,sheet);next.width=Math.max(100,next.width);next.height=Math.max(100,next.height);Object.assign(chart,next);
  if(chart.source){chart.source.anchor=anchor;chart.source.original.geometry=next;chart.source.original.placement=placement;}
 }
 for(const group of sheet.xlsxDrawings||[])group.anchors=group.anchors.map(source=>transformAnchor(source,before,sheet,operation));
}
export const dimensionSnapshot=sheet=>({...sheet,columnWidths:{...sheet.columnWidths},rowHeights:{...sheet.rowHeights}});
