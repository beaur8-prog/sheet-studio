export const isErrorText=value=>typeof value==='string'&&/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|SPILL!|CALC!|ERROR!)$/.test(value);
