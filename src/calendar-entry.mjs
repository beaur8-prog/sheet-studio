import {textEntryStyle} from './cell-values.mjs';
import {numberFormatKind} from './number-formats.mjs';

// Restrict recognition to unambiguous ISO dates and 24-hour times in a date/time
// destination. General/Text identifiers keep their existing string identity.
// UTC arithmetic avoids locale guessing and DST-dependent off-by-one days.
export function calendarEntry(value,style={}, {date1904=false,kind}={}){
  if(typeof value!=='string'||textEntryStyle(style)||value.startsWith("'")||value.startsWith('='))return null;
  const formattedDate=style.format==='date'||(!style.format||style.format==='custom')&&numberFormatKind(style.numberFormat)==='date';
  if(!['date','time'].includes(kind)&&!formattedDate)return null;
  const text=value.trim(),date=/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(text),time=/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?$/.exec(text);
  if(!date&&!time)return null;
  let serial=0,hasTime=false;
  if(date){
    const [,year,month,day,hour,minute,second,millis]=date,y=+year,m=+month,d=+day;
    if(y<1900||y>9999||m<1||m>12||d<1)return null;
    const fakeLeap=y===1900&&m===2&&d===29,stamp=Date.UTC(y,m-1,d),actual=new Date(stamp);
    if(!fakeLeap&&(actual.getUTCFullYear()!==y||actual.getUTCMonth()!==m-1||actual.getUTCDate()!==d))return null;
    serial=fakeLeap?60:(stamp-Date.UTC(1899,11,31))/86400000+(stamp>=Date.UTC(1900,2,1)?1:0);
    if(date1904)serial-=1462;
    if(hour!==undefined){const h=+hour,n=+minute,s=+(second||0),ms=+(millis||'').padEnd(3,'0');if(h>23||n>59||s>59)return null;serial+=(h*3600+n*60+s+ms/1000)/86400;hasTime=true;}
  }else{
    const [,hour,minute,second,millis]=time,h=+hour,m=+minute,s=+(second||0),ms=+(millis||'').padEnd(3,'0');if(h>23||m>59||s>59)return null;serial=(h*3600+m*60+s+ms/1000)/86400;hasTime=true;
  }
  const general=(!style.format||style.format==='general')&&(!style.numberFormat||/^General$/i.test(style.numberFormat));
  const numberFormat=date?'yyyy-mm-dd'+(hasTime?' hh:mm:ss':''):'hh:mm'+(time[3]!==undefined?':ss':'');
  return {raw:String(serial),...(general?{style:{format:'date',numberFormat}}:{})};
}
