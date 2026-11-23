/** Constance & Hervé — Google Sheet Mariage 2027. Installer comme Web App. */
const SHEET_ID='1CqqkFtWsWQsHKg9AsxlaJvkLx3w7qXnbPuPgvMYSr5g';
const TAB='Réponses site mariage';
function doPost(e){
  try{
    const p=JSON.parse(e.postData.contents||'{}');
    const name=String(p.name||'').trim().slice(0,120);
    if(name.length<2||!['oui','non'].includes(p.attending)) throw Error('Réponse invalide');
    const adults=p.attending==='oui'?Number(p.adults):0;
    const children=p.attending==='oui'?Number(p.children):0;
    if(!Number.isInteger(adults)||!Number.isInteger(children)||adults<0||children<0||adults+children>25) throw Error('Nombre invalide');
    const safe=x=>{const s=String(x||'').trim().slice(0,1000);return /^[=+\-@]/.test(s)?"'"+s:s};
    const lock=LockService.getScriptLock();lock.waitLock(20000);
    try{
      const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
      if(!sheet)throw Error('Onglet introuvable');
      const now=new Date(),id=Utilities.getUuid();
      sheet.appendRow([id,now,safe(name),'',p.attending,adults,children,safe(p.lodging),'',safe(p.notes),now,'Nouvelle réponse']);
      return ContentService.createTextOutput(JSON.stringify({ok:true,id})).setMimeType(ContentService.MimeType.JSON);
    }finally{lock.releaseLock()}
  }catch(err){return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message)})).setMimeType(ContentService.MimeType.JSON)}
}