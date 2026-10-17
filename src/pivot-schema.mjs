export const aggregations = Object.freeze({sum:'합계',count:'개수',countNumbers:'숫자 개수',average:'평균',min:'최솟값',max:'최댓값'});
export const scalarKey = value => JSON.stringify([typeof value, value]);
export const pivotLabel = value => value === '' || value === null ? '(비어 있음)' : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value);

export function validatePivot(spec) {
  const field = v => typeof v === 'string' && v.length > 0 && v.length <= 255;
  const axis = v => Array.isArray(v) && v.length <= 8 && v.every(field) && new Set(v).size === v.length;
  const scalar = v => v === null || typeof v === 'boolean' || (typeof v === 'string' && v.length <= 32768) || (typeof v === 'number' && Number.isFinite(v));
  if (!spec || spec.version !== 1 || !field(spec.sourceSheetId) || typeof spec.sourceRange !== 'string' || !/^(?:[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*|#REF!)$/.test(spec.sourceRange) || !axis(spec.rows) || !axis(spec.columns) || !Array.isArray(spec.values) || !spec.values.length || spec.values.length > 16 || spec.values.some(v => !v || !field(v.field) || !Object.hasOwn(aggregations,v.aggregate)) || !Array.isArray(spec.filters) || spec.filters.length > 16 || spec.filters.some(f => !f || !field(f.field) || !Array.isArray(f.values) || f.values.length > 10000 || !f.values.every(scalar)) || typeof spec.totals !== 'boolean' || !['asc','desc','source'].includes(spec.sort)) throw new Error('피벗테이블 설정이 올바르지 않습니다.');
  if (new Set([...spec.rows,...spec.columns]).size !== spec.rows.length+spec.columns.length) throw new Error('같은 필드를 행과 열에 중복 배치할 수 없습니다.');
  if (new Set(spec.filters.map(f=>f.field)).size !== spec.filters.length) throw new Error('같은 필터 필드를 중복 배치할 수 없습니다.');
  return spec;
}
