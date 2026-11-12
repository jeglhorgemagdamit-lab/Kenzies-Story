/** Kenzie's Story — Google Sheets database.
 * Create this project from Extensions > Apps Script in the RSVP spreadsheet.
 * Script properties: RSVP_SHARED_SECRET (random secret, also set in Netlify).
 * Only the Netlify function should know the secret.
 */
const FAMILY_SHEET = 'Families';
const GUEST_SHEET = 'Guests';
const FAMILY_HEADERS = ['Invitation Code','Family','Contact','RSVP Status','Adults','Kids','Total Guests','Last Updated'];
const GUEST_HEADERS = ['Invitation Code','Family','Guest Name','Category','Last Updated'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kenzie RSVP')
    .addItem('Set up guest sheets', 'setupSheets')
    .addItem('Generate codes for new families', 'generateMissingCodes')
    .addToUi();
}
function setupSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ensureHeader_(ss,FAMILY_SHEET,FAMILY_HEADERS);
  ensureHeader_(ss,GUEST_SHEET,GUEST_HEADERS);
  SpreadsheetApp.getUi().alert('Ready! Enter family names in the Families tab, column B.');
}
function ensureHeader_(ss,name,headers){
  let sheet=ss.getSheetByName(name)||ss.insertSheet(name);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}
function generateMissingCodes(){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FAMILY_SHEET);
  if(!sheet) throw new Error('Use Kenzie RSVP > Set up guest sheets first.');
  const last=sheet.getLastRow();if(last<2)return;
  const values=sheet.getRange(2,1,last-1,2).getValues();
  const existing=new Set(values.map(r=>String(r[0]).trim().toUpperCase()).filter(Boolean));
  let count=0;
  values.forEach((r,i)=>{
    if(String(r[1]).trim() && !String(r[0]).trim()){
      let code;
      do { code='KZ-'+Utilities.getUuid().replace(/-/g,'').slice(0,20).toUpperCase(); } while(existing.has(code));
      existing.add(code);
      sheet.getRange(i+2,1).setValue(code);
      sheet.getRange(i+2,4).setValue('Pending');
      count++;
    }
  });
  SpreadsheetApp.getUi().alert(count+' new family code(s) created. Keep the sheet private.');
}
function doPost(e){
  try {
    if(!e||!e.postData||!e.postData.contents) return json_({ok:false,message:'Missing request.'});
    if(e.postData.contents.length>18000) return json_({ok:false,message:'Request too large.'});
    const data=JSON.parse(e.postData.contents);
    const secret=PropertiesService.getScriptProperties().getProperty('RSVP_SHARED_SECRET');
    if(!secret || !data.secret || data.secret!==secret) return json_({ok:false,message:'Unauthorized request.'});
    if(data.action==='lookup') return json_(lookup_(data));
    if(data.action==='submit') return json_(submit_(data));
    return json_({ok:false,message:'Unknown action.'});
  }catch(err){console.error(err);return json_({ok:false,message:'Unable to process RSVP. Please try again.'});}
}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function sheets_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const families=ss.getSheetByName(FAMILY_SHEET),guests=ss.getSheetByName(GUEST_SHEET);
  if(!families||!guests)throw new Error('Missing sheets.');return {families,guests};
}
function clean_(value,max){return String(value==null?'':value).trim().slice(0,max);}
function code_(v){return clean_(v,64).toUpperCase();}
function findFamily_(sheet,code){
  if(!/^KZ-[A-F0-9]{20}$/.test(code)) return null;
  const last=sheet.getLastRow();if(last<2)return null;
  // Read only the code column for lookups; no other family's data leaves the backend.
  const match=sheet.getRange(2,1,last-1,1).createTextFinder(code).matchEntireCell(true).matchCase(false).findNext();
  return match?match.getRow():null;
}
function safeCell_(s){return /^[=+\-@\t\r]/.test(s)?"'"+s:s;}
function lookup_(data){
  const code=code_(data.code),s=sheets_(),row=findFamily_(s.families,code);
  if(!row)return {ok:false,message:'This invitation code was not found. Please check your family link.'};
  const familyData=s.families.getRange(row,1,1,8).getValues()[0];
  const status=clean_(familyData[3],20)||'Pending';
  const guests=[];
  const last=s.guests.getLastRow();
  if(last>=2){
    s.guests.getRange(2,1,last-1,4).getValues().forEach(r=>{
      if(code_(r[0])===code)guests.push({name:String(r[2]),category:String(r[3])});
    });
  }
  return {ok:true,family:String(familyData[1]),status,contact:String(familyData[2]||''),guests};
}
function submit_(data){
  const code=code_(data.code),contact=clean_(data.contact,100);
  if(!contact)return {ok:false,message:'Please provide a contact name.'};
  const yes=data.attendance==='yes';
  if(!yes && data.attendance!=='no')return {ok:false,message:'Please choose whether you are attending.'};
  const inputGuests=Array.isArray(data.guests)?data.guests:[];
  if(inputGuests.length>35)return {ok:false,message:'Please contact the organizer for a larger party.'};
  const guests=yes?inputGuests.map(g=>({name:clean_(g.name,100),category:clean_(g.category,10)})):[];
  if(yes&&(!guests.length||guests.some(g=>!g.name||!['Adult','Kid'].includes(g.category))))
    return {ok:false,message:'Enter each guest name and select Adult or Kid.'};
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))return {ok:false,message:'The guest list is busy. Please try again.'};
  try{
    const s=sheets_(),row=findFamily_(s.families,code);
    if(!row)return {ok:false,message:'This invitation code was not found.'};
    const family=String(s.families.getRange(row,2).getValue());
    const adult=guests.filter(g=>g.category==='Adult').length;
    const kid=guests.length-adult;
    const now=new Date();
    // One row per family, overwritten when that family updates their RSVP.
    s.families.getRange(row,3,1,6).setValues([[safeCell_(contact),yes?'Attending':'Declined',adult,kid,guests.length,now]]);
    // Replace this family's guest rows without changing other families' records.
    const last=s.guests.getLastRow();
    if(last>=2){
      const codes=s.guests.getRange(2,1,last-1,1).getValues();
      for(let i=codes.length-1;i>=0;i--)if(code_(codes[i][0])===code)s.guests.deleteRow(i+2);
    }
    if(guests.length){
      const rows=guests.map(g=>[code,safeCell_(family),safeCell_(g.name),g.category,now]);
      const start=s.guests.getLastRow()+1;
      s.guests.getRange(start,1,rows.length,5).setValues(rows);
    }
    SpreadsheetApp.flush();
    return {ok:true,family,adults:adult,kids:kid,total:guests.length};
  }finally{lock.releaseLock();}
}
