import {unwrap} from './clone.mjs';
import {xml,stringify,children,child,all,attr} from './chart-xml.mjs';
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
    if(feature.conditionalExtensions)feature.conditionalExtensions=feature.conditionalExtensions.map(source=>{
      const doc=xml(source),root=doc.documentElement;
      for(const formats of children(root,'conditionalFormattings')){
        for(const group of children(formats,'conditionalFormatting')){
          const node=child(group,'sqref'),original=node?.textContent||'',ranges=original.trim().split(/\s+/).filter(Boolean).map(reference).filter(r=>!r.includes('#REF!'));
          if(!ranges.length){formats.removeChild(group);continue;}
          node.textContent=ranges.join(' ');for(const f of all(group,'f'))f.textContent=formula(rebase(f.textContent,original.trim().split(/\s+/)[0]));
        }
        if(!children(formats,'conditionalFormatting').length)root.removeChild(formats);
      }
      return children(root).length?stringify(doc):null;
    }).filter(Boolean);
    if(sheet.unsupportedConditionalCount&&!feature.conditionalFormatting?.length&&!feature.conditionalExtensions?.length)delete sheet.unsupportedConditionalCount;
  }
  const cells=unwrap(sheet.cells);for(const key in cells){
    const cell=sheet.cells[key],original=cell.xlsxFormula;
    if(original){original.raw=formula(original.raw);original.formula=formula(original.formula);if(original.ref){original.ref=reference(original.ref);if(original.ref.includes('#REF!'))delete cell.xlsxFormula;}}
    if(cell.cachedFormula?.startsWith('='))cell.cachedFormula=formula(cell.cachedFormula);
    if(cell.xlsxExternal?.raw)cell.xlsxExternal.raw=formula(cell.xlsxExternal.raw);
    const metadata=cell.xlsxMetadata;if(metadata){if(metadata.raw.startsWith('='))metadata.raw=formula(metadata.raw);if(metadata.array){metadata.array=reference(metadata.array);if(metadata.array.includes('#REF!'))delete cell.xlsxMetadata;}}
  }
}
