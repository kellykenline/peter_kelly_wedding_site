/** Bind this Apps Script to the LIVE Google Sheet. Deploy as a Web App. */
const SHEET_NAME='RSVP Tracking';
const HEADERS={first:'First Name',last:'Last Name',child:'Child',family:'Family Group',salutation:'Salutation',code:'Codes',submitted:'Submitted At',attending:'Attending',meal:'Meal Selection',brunch:'Attending Breakfast',notAttending:'Not Attending',comments:'Other Comments'};
/** Basic Apps Script throttling. Shared quotas protect against bulk guessing;
 * per-code quotas slow repeated requests against the same household.
 * Apps Script web apps do not reliably provide visitor IP addresses, so this
 * is best-effort abuse prevention, NOT per-visitor rate limiting.
 */
const LOOKUP_WINDOW_SECONDS = 600;
const MAX_GLOBAL_LOOKUPS = 120;
const MAX_LOOKUPS_PER_CODE = 12;
function limitLookup(code) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw Error('RSVP service is busy. Please try again shortly.');
  try {
    const cache = CacheService.getScriptCache();
    const bucket = Math.floor(Date.now() / (LOOKUP_WINDOW_SECONDS * 1000));
    const globalKey = 'rsvp:lookup:global:' + bucket;
    // Hash the submitted code so raw invitation codes are not cache keys.
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(code));
    const hex = digest.slice(0,12).map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
    const codeKey = 'rsvp:lookup:code:' + bucket + ':' + hex;
    const counts = cache.getAll([globalKey, codeKey]);
    const globalCount = Number(counts[globalKey] || 0);
    const codeCount = Number(counts[codeKey] || 0);
    if (globalCount >= MAX_GLOBAL_LOOKUPS || codeCount >= MAX_LOOKUPS_PER_CODE) {
      throw Error('Too many RSVP lookups. Please try again in about 10 minutes.');
    }
    cache.putAll({[globalKey]: String(globalCount + 1), [codeKey]: String(codeCount + 1)}, LOOKUP_WINDOW_SECONDS + 20);
  } finally { lock.releaseLock(); }
}
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw Error('Invalid request');
    const p = JSON.parse(e.postData.contents);
    if (!p || !['lookup','save'].includes(p.action)) throw Error('Unknown action');
    if (p.action === 'lookup') {
      limitLookup(clean(p.code).toUpperCase());
      return respond(handle(p));
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return respond(handle(p)); }
    finally { lock.releaseLock(); }
  } catch (err) { return respond({ok:false,error:String(err.message || err)}); }
}
function respond(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON)}
function clean(v){return String(v==null?'':v).trim()}
function safe(v,max){const s=clean(v).slice(0,max);return /^[=+@\-\t\r]/.test(s)?"'"+s:s}
function handle(p){const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);if(!sheet)throw Error('RSVP Tracking tab not found');const data=sheet.getDataRange().getValues();const names=data[0].map(clean);const ix={};Object.keys(HEADERS).forEach(k=>{ix[k]=names.indexOf(HEADERS[k]);if(ix[k]<0)throw Error('Missing column: '+HEADERS[k])});const code=clean(p.code).toUpperCase();if(!/^[A-Z0-9-]{3,40}$/.test(code))throw Error('Invalid household code');const matched=[];for(let i=1;i<data.length;i++)if(clean(data[i][ix.code]).toUpperCase()===code&&clean(data[i][ix.first]))matched.push({row:i+1,values:data[i]});if(!matched.length)throw Error('No household found for that code. Please check your invitation.');if(matched.length>30)throw Error('Household too large; please contact the couple.');
if(p.action==='save'){if(!Array.isArray(p.guests)||p.guests.length!==matched.length)throw Error('All household guests must be included');const byRow=new Map(p.guests.map(g=>[Number(g.row),g]));if(byRow.size!==matched.length)throw Error('Duplicate guest');const comment=safe(p.comments,2000);const now=new Date();const writes=[];for(const item of matched){const g=byRow.get(item.row);if(!g)throw Error('Invalid guest list');const attending=clean(g.attending),meal=clean(g.meal),brunch=clean(g.brunch);if(!['Yes','No'].includes(attending))throw Error('Select attendance for everyone');const child=clean(item.values[ix.child])==='1';const allowed=child?['Chicken Fingers','Pasta']:['Filet','Chicken','Vegetarian'];if(attending==='Yes'&&(!allowed.includes(meal)||!['Yes','No'].includes(brunch)))throw Error('Invalid meal or brunch choice');writes.push({row:item.row,attending,meal:attending==='Yes'?meal:'',brunch:attending==='Yes'?brunch:'',notAttending:attending==='No'?'Yes':'',comment,now})}
// Only update RSVP columns. Do not overwrite names, codes, addresses, or family grouping.
for(const w of writes){const set=(key,value)=>sheet.getRange(w.row,ix[key]+1).setValue(value);set('submitted',w.now);set('attending',w.attending==='Yes'?'Yes':'');set('meal',w.meal);set('brunch',w.brunch);set('notAttending',w.notAttending);set('comments',w.comment)}SpreadsheetApp.flush();}
else if(p.action!=='lookup')throw Error('Unknown action');
const fresh=sheet.getDataRange().getValues();const first=matched[0].values;return {ok:true,label:clean(first[ix.salutation])||'Your Household',comments:clean(fresh[matched[0].row-1][ix.comments]),guests:matched.map(item=>{const r=fresh[item.row-1],att=clean(r[ix.attending]),decline=clean(r[ix.notAttending]);return {row:item.row,firstName:clean(r[ix.first]),lastName:clean(r[ix.last]),child:clean(r[ix.child])==='1',attending:att==='Yes'?'Yes':decline==='Yes'?'No':'',meal:clean(r[ix.meal]),brunch:clean(r[ix.brunch])}})};
}

/** No public household lookups via GET, and no JSONP responses. */
function doGet(e) {
  return ContentService.createTextOutput('RSVP service active. Use the RSVP page to submit requests.')
    .setMimeType(ContentService.MimeType.TEXT);
}
