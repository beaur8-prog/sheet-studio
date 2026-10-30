/** A file's stored result is a labelled presentation value, never proof that
 * an unsupported formula or external connection was recalculated here. */
export function calculationDisplay(result){
  return result.cached&&!result.stale?{...result,value:result.cachedValue,error:result.cachedError||null,calculationError:result.error}:result;
}
