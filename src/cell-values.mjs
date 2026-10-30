export const EXCEL_ERROR_VALUES=Object.freeze(['#N/A','#REF!','#VALUE!','#DIV/0!','#NUM!','#NAME?','#NULL!','#GETTING_DATA','#SPILL!','#CALC!','#FIELD!','#BLOCKED!','#UNKNOWN!','#CONNECT!','#BUSY!','#PYTHON!']);
const errorValues=new Set([...EXCEL_ERROR_VALUES,'#CYCLE!','#ERROR!']),tokens=[...errorValues].sort((a,b)=>b.length-a.length);
export const isErrorText=value=>typeof value==='string'&&errorValues.has(value);
export function readErrorToken(text){if(text[0]!=='#')return null;const upper=text.toUpperCase();return tokens.find(value=>upper.startsWith(value))||null;}

// Entry depends on the destination format. Changing a format alone never changes
// the type of an existing value, and values-only paste preserves its source type.
export const textEntryStyle=style=>style?.format==='text'||(!style?.format||style.format==='custom')&&/^@$/i.test(style?.numberFormat||'');
export function entryRaw(value,style){
  const raw=String(value??'');
  if(!textEntryStyle(style)||!raw||raw.startsWith("'"))return raw;
  return /^[=]/.test(raw)||!Number.isNaN(Number(raw))||/^(TRUE|FALSE)$/i.test(raw)||isErrorText(raw)||/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%$/i.test(raw.trim())?"'"+raw:raw;
}
