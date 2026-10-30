import JSZip from 'jszip';
import {createWorkbook} from '../../src/engine.mjs';
import {writeSpreadsheet} from '../../src/io.mjs';
import {xml,stringify,all,child,attr} from '../../src/chart-xml.mjs';
/** A synthetic, inert future-metadata extension with valid cm/vm indices.
 * This exercises opaque OOXML preservation, not stock/linked-data calculation. */
export async function valueMetadataFixture({unknown=false}={}){
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=SEQUENCE(3)'},D1:{raw:'=SEQUENCE(3)'},G1:{raw:'=SEQUENCE(3)'},G2:{raw:'',note:'metadata note'}};
 const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data),sheet=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
 for(const ref of ['A1','A2','D2','G2'])all(sheet,'c').find(c=>attr(c,'r')===ref).setAttribute('vm','1');
 if(unknown)child(all(sheet,'c').find(c=>attr(c,'r')==='A1'),'f').textContent='_xlfn.FUTUREARRAY(3)';
 zip.file('xl/worksheets/sheet1.xml',stringify(sheet));
 const metadata=xml(await zip.file('xl/metadata.xml').async('string')),root=metadata.documentElement,ns=root.namespaceURI,types=child(root,'metadataTypes'),type=metadata.createElementNS(ns,'metadataType');
 for(const [key,value] of Object.entries({name:'APP_PRIVATE',minSupportedVersion:120000,copy:1,pasteAll:1,pasteValues:1,merge:1,splitFirst:1,rowColShift:1,assign:1,coerce:1,cellMeta:0}))type.setAttribute(key,String(value));types.appendChild(type);types.setAttribute('count','2');
 const future=metadata.importNode(xml('<futureMetadata xmlns="'+ns+'" name="APP_PRIVATE" count="1"><bk><extLst><ext uri="urn:sheet-studio:test:opaque"><t:payload xmlns:t="urn:sheet-studio:test:opaque">original value metadata</t:payload></ext></extLst></bk></futureMetadata>').documentElement,true);root.insertBefore(future,child(root,'cellMetadata'));
 const values=metadata.importNode(xml('<valueMetadata xmlns="'+ns+'" count="1"><bk><rc t="2" v="0"/></bk></valueMetadata>').documentElement,true);root.insertBefore(values,child(root,'extLst')||null);zip.file('xl/metadata.xml',stringify(metadata));return zip.generateAsync({type:'uint8array'});
}
