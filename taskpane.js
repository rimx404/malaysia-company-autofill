const statusBox=document.getElementById('status');
function show(s){statusBox.textContent=s;}
Office.onReady(async info=>{
 if(info.host!==Office.HostType.Excel){show('Open this add-in inside Excel.');return;}
 try{
  await Excel.run(async context=>{
   const sheet=context.workbook.worksheets.getItemOrNullObject('Company Master');
   sheet.load('name');await context.sync();
   if(sheet.isNullObject){show('Company Master sheet not found.');return;}
   sheet.onChanged.add(async event=>{
    try{
     if(event.changeType!=='RangeEdited')return;
     const addr=event.address.split('!').pop().replace(/\$/g,'');
     const m=/^D(\d+)$/.exec(addr);
     if(!m||Number(m[1])<7)return;
     await Excel.run(async ctx=>{
      const cell=ctx.workbook.worksheets.getItem('Company Master').getRange(addr);
      cell.load('values');await ctx.sync();
      show('Detected company in '+addr+': '+String(cell.values[0][0]||'(empty)')+'\nNo cells were changed.');
     });
    }catch(e){show('Detection error: '+e.message);}
   });
   await context.sync();show('Listening for edits to Company Master column D (rows 7+).');
  });
 }catch(e){show('Setup error: '+e.message);}
});
