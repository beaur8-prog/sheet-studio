import {unwrap} from './clone.mjs';
import {xml,stringify,children,all,attr} from './chart-xml.mjs';
import {transformExtraSheetFeatures} from './worksheet-extra.mjs';

/** Keep preserved XML references in step with supported sheet edits. Formula
 * parsers are supplied by the caller to avoid a second, divergent ref parser.
 */
export function transformPreservedFeatures(sheet,{reference=f=>f,formula=f=>f,rebase=f=>f}={}){
  const feature=sheet.xlsxFeatures;
  if(feature){
    transformExtraSheetFeatures(feature,reference);
    const transform=(source,kind)=>{
      const doc=xml(source),root=doc.documentElement,groups=kind==='dataValidations'?children(root,'dataValidation'):[root];
      for(const group of groups){
        const original=attr(group,'sqref'),ranges=original.split(/\s+/).filter(Boolean).map(reference).filter(r=>!r.includes('#REF!'));
        if(!ranges.length){if(group===root)return null;root.removeChild(group);continue;}
        group.setAttribute('sqref',ranges.join(' '));
        for(const name of ['formula','formula1','formula2'])for(const node of all(group,name))node.textContent=formula(rebase(node.textContent,original.split(/\s+/)[0]));
        for(const node of all(group,'cfvo'))if(attr(node,'type')==='formula')node.setAttribute('val',formula(rebase(attr(node,'val'),original.split(/\s+/)[0])));
      }
      if(kind==='dataValidations'){const count=children(root,'dataValidation').length;if(!count)return null;root.setAttribute('count',String(count));}
      return stringify(doc);
    };
    if(feature.dataValidations){feature.dataValidations=transform(feature.dataValidations,'dataValidations');if(!feature.dataValidations)delete feature.dataValidations;}
    if(feature.conditionalFormatting)feature.conditionalFormatting=feature.conditionalFormatting.map(s=>transform(s,'conditionalFormatting')).filter(Boolean);
  }
  const cells=unwrap(sheet.cells);for(const key in cells){
    const cell=sheet.cells[key],original=cell.xlsxFormula;
    if(original){original.raw=formula(original.raw);original.formula=formula(original.formula);if(original.ref){original.ref=reference(original.ref);if(original.ref.includes('#REF!'))delete cell.xlsxFormula;}}
    if(cell.cachedFormula?.startsWith('='))cell.cachedFormula=formula(cell.cachedFormula);
    if(cell.xlsxExternal?.raw)cell.xlsxExternal.raw=formula(cell.xlsxExternal.raw);
    const metadata=cell.xlsxMetadata;if(metadata){if(metadata.raw.startsWith('='))metadata.raw=formula(metadata.raw);if(metadata.array){metadata.array=reference(metadata.array);if(metadata.array.includes('#REF!'))delete cell.xlsxMetadata;}}
  }
}
