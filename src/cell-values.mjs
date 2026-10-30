export const isErrorText=value=>typeof value==='string'&&/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|SPILL!|CALC!|ERROR!)$/.test(value);

// Entry depends on the destination format. Changing a format alone never changes
// the type of an existing value, and values-only paste preserves its source type.
export const textEntryStyle=style=>style?.format==='text'||(!style?.format||style.format==='custom')&&/^@$/i.test(style?.numberFormat||'');
export function entryRaw(value,style){
  const raw=String(value??'');
  if(!textEntryStyle(style)||!raw||raw.startsWith("'"))return raw;
  return /^[=]/.test(raw)||!Number.isNaN(Number(raw))||/^(TRUE|FALSE)$/i.test(raw)||isErrorText(raw)||/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%$/i.test(raw.trim())?"'"+raw:raw;
}
