// The application model uses the same JSON representation as .sheet.json.
// Sending UTF-8 bytes between workers avoids cloning millions of cell objects
// twice. The parser validates the model before encoding; the private receiving
// port belongs only to that parser, and decoding never changes the live book.
export function encodeWorkbookWire(book){return new TextEncoder().encode(JSON.stringify(book));}
export function decodeWorkbookWire(bytes){
  if(!(bytes instanceof Uint8Array))throw Error('읽은 문서의 전달 형식이 올바르지 않습니다.');
  return JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes));
}
