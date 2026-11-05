import {xmlEscape} from './chart-xml.mjs';
export function baseDataBarXML(rule,priority){
  const points=rule.cfvo.map(point=>`<cfvo type="${xmlEscape(point.type)}"${point.value===undefined?'':' val="'+xmlEscape(point.value)+'"'}/>`).join('');
  return `<conditionalFormatting xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sqref="${xmlEscape(rule.range)}"><cfRule type="dataBar" priority="${priority}"><dataBar minLength="10" maxLength="90" showValue="1">${points}<color rgb="FF${rule.barColor.slice(1).toUpperCase()}"/></dataBar></cfRule></conditionalFormatting>`;
}
