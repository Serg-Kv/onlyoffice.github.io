const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const assert = require('node:assert/strict');
const source = fs.readFileSync(process.argv[2] || path.join(__dirname, 'cell/mock-data-generator.js'), 'utf8');
function setup(o = {}) {
    const fields = o.fields || ['Name', 'Email'];
    const state = { writes: [], prompts: [], actions: [], commands: 0, active: 0, selections: 0 };
    const cells = new Map();
    const firstRow = o.firstRow || 1, firstCol = o.firstCol || 1;
    const key = (s,r,c) => `${s}:${r}:${c}`;
    const occupiedKey = key(0, firstRow + 1, firstCol);
    if ('occupied' in o) cells.set(occupiedKey, {value:o.occupied});
    if (o.formula) cells.set(occupiedKey, {value:'',formula:'=IF(TRUE,"",1)'});
    function range(s,r,c,h,w) {
        return {
            Row:r, Col:c, GetColumnsCount:()=>w, GetRowsCount:()=>h,
            GetValue2:()=>w===1?fields[0]:[Array.from({length:w},(_,i)=>fields[i] ?? '')],
            Resize:(hh,ww)=>range(s,r,c,hh,ww),
            GetRows:i=>range(s,r+i-1,c,1,w), GetCells:i=>range(s,r,c+i-1,1,1),
            GetValue:()=>cells.get(key(s,r,c))?.value ?? '',
            GetFormula:()=>cells.get(key(s,r,c))?.formula ?? '',
            SetValue:value=>{state.writes.push({sheet:s,row:r,col:c,value});cells.set(key(s,r,c),{value});}
        };
    }
    function sheet(s) {
        return {
            get Selection(){state.selections++;return o.noSelection?null:range(s,firstRow,firstCol,o.height||1,o.width||fields.length);},
            GetRange:address=>{
                if(address==='bad'){if(o.rangeThrows)throw Error('bad range');return null;}
                const m=/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(address);
                if(!m)throw Error('bad address: '+address);
                const col=x=>[...x].reduce((v,ch)=>v*26+ch.charCodeAt(0)-64,0);
                return range(s,+m[2],col(m[1]),+(m[4]||m[2])-m[2]+1,col(m[3]||m[1])-col(m[1])+1);
            }
        };
    }
    const ctx={RegisteredFunction:function(x){Object.assign(this,x);},window:{AgentState:{ToolError:Error}},Asc:{scope:{},Editor:{}},Api:{GetActiveSheet:()=>sheet(state.active)}};
    ctx.AI={ActionType:{Chat:1},Request:{create:()=>({modelUI:{name:'test'},chatRequest:async prompt=>{
        state.prompts.push(prompt);
        if(o.reject)throw Error('offline');
        if(o.editDuringAI)cells.set(occupiedKey,{value:'USER EDIT'});
        if(o.switchSheet)state.active=1;
        const m=/exactly (\d+) rows x (\d+) columns/.exec(prompt);
        return o.response ?? JSON.stringify(Array.from({length:+m[1]},()=>Array.from({length:+m[2]},()=> 'sample')));
    }})}};
    vm.createContext(ctx);
    ctx.Asc.Editor.callCommand=async fn=>{
        state.commands++;
        if(o.editBeforeWrite && state.commands===4)cells.set(occupiedKey,{value:'USER EDIT'});
        return vm.runInContext('('+fn.toString()+')()',ctx);
    };
    ctx.Asc.Editor.callMethod=async(name,args)=>{
        state.actions.push([name,args[0]]);
        if(o.groupFails && name==='StartAction' && args[0]==='GroupActions')throw Error('group startup failed');
    };
    const result=vm.runInContext(source,ctx);
    const tool=process.argv.includes('--bundle')?ctx.HELPERS.cell.find(f=>f.name==='mockDataGenerator'):result;
    return {tool,state};
}
let pass=0,fail=0;
async function test(name,o,params,verify){
    const s=setup(o);let error;
    try{await s.tool.call(params);}catch(e){error=e;}
    try{verify(s.state,error,s.tool);pass++;console.log('PASS '+name);}
    catch(e){fail++;console.log('FAIL '+name+': '+e.message);}
}
const rejectEarly=(s,e)=>{assert(e);assert.equal(s.prompts.length,0);assert.equal(s.writes.length,0);};
(async()=>{
    for(const o of [{},{fields:['Name']},{height:6},{firstRow:7,firstCol:27}])
        await test('basic '+JSON.stringify(o),o,{rows:2},(s,e)=>{assert.ifError(e);assert.equal(s.writes.length,2*(o.fields?.length||2));assert.equal(s.writes[0].row,(o.firstRow||1)+1);assert.equal(s.writes[0].col,o.firstCol||1);});
    await test('whole row trims to 3 columns',{fields:['Name','Email','Age'],width:16384},{rows:2},(s,e)=>{assert.ifError(e);assert.equal(s.writes.length,6);});
    await test('interior empty preserved, trailing empty trimmed',{fields:['Name','','Age',' ']},{rows:2},(s,e)=>{assert.ifError(e);assert.equal(s.writes.length,6);});
    await test('51 columns rejected',{fields:Array(51).fill('x')},{rows:2},rejectEarly);
    await test('empty headers rejected',{fields:['',' ']},{rows:2},rejectEarly);
    await test('default 10 rows',{}, {},(s,e)=>{assert.ifError(e);assert.equal(s.writes.length,20);});
    await test('100 rows allowed',{}, {rows:100},(s,e)=>{assert.ifError(e);assert.equal(s.writes.length,200);});
    for(const rows of [null,0,-1,101,1000,1.5,'2',NaN,Infinity])await test('rows '+String(rows),{},{rows},rejectEarly);
    for(const range of [null,'',42,'bad'])await test('range '+String(range),{},{rows:2,range},rejectEarly);
    await test('range throws',{rangeThrows:true},{rows:2,range:'bad'},rejectEarly);
    for(const occupied of ['value',' ',0,false])await test('occupied '+JSON.stringify(occupied),{occupied},{rows:2},rejectEarly);
    await test('empty-result formula',{formula:true},{rows:2},rejectEarly);
    await test('edit during AI',{editDuringAI:true},{rows:2},(s,e)=>{assert(e);assert.equal(s.writes.length,0);assert.equal(s.prompts.length,1);});
    await test('edit between final check and write',{editBeforeWrite:true},{rows:2},(s,e)=>{assert(e,'concurrent edit was overwritten');assert.equal(s.writes.length,0);});
    await test('sheet switch during AI',{switchSheet:true},{rows:2},(s,e)=>{assert(e || s.writes.every(w=>w.sheet===0),'wrote to newly active sheet');});
    for(const o of [{reject:true},{groupFails:true}])await test('cleanup '+JSON.stringify(o),o,{rows:2},(s,e)=>{assert(e);assert.equal(s.writes.length,0);assert(s.actions.some(a=>a[0]==='EndAction' && a[1]==='Block'),'Block left open');});
    for(const response of ['not JSON','[[1]]'])await test('bad response '+response,{response},{rows:2},(s,e)=>{assert(e);assert.equal(s.writes.length,0);});
    console.log(`RESULT ${pass} passed, ${fail} failed`);process.exitCode=fail?1:0;
})().catch(e=>{console.error(e);process.exitCode=1;});
