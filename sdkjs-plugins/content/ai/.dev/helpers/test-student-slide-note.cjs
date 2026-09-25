const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');

// Reuse editor mocks, but load the submitted source or generated bundle directly.
let harness = fs.readFileSync(path.join(__dirname, 'test-add-note-to-slide.cjs'), 'utf8');
harness = harness.slice(0, harness.indexOf('(async () =>'));
const start = harness.indexOf('    vm.runInContext(fs.readFileSync');
const end = harness.indexOf('    return { tool', start);
const target = process.argv[2] || path.join(__dirname, 'slide/add-note-to-slide.js');
const bundle = process.argv.includes('--bundle');
const loader = bundle
    ? `vm.runInContext(fs.readFileSync(${JSON.stringify(target)}, 'utf8'), c); const tool = c.HELPERS.slide.find(f => f.name === 'addNoteToSlide');`
    : `const tool = vm.runInContext(fs.readFileSync(${JSON.stringify(target)}, 'utf8'), c);`;
harness = harness.slice(0, start) + loader + '\n' + harness.slice(end);
const setup = new Function('require', harness + '\nreturn setup;')(require);
let passed = 0, failed = 0;
async function test(name, opts, params, verify) {
    const s = setup(opts);
    let error;
    try { await s.tool.call(params); } catch (e) { error = e; }
    try { verify(s, error); passed++; console.log('PASS ' + name); }
    catch (e) { failed++; console.log('FAIL ' + name + ': ' + e.message); }
}
const rejected = (s, e) => { assert(e, 'expected rejection'); assert.equal(s.writes.length, 0); };
const ended = (s, kind) => s.actions.some(a => a[0] === 'EndAction' && a[1] === kind);
(async () => {
    for (const params of [{text:'Exact'}, {text:'Exact', slideNumber:2}]) {
        await test('literal ' + JSON.stringify(params), {}, params, (s,e) => {
            assert.ifError(e); assert.equal(s.calls,1);
            assert.deepEqual(s.writes, [[params.slideNumber === 2 ? 1 : 0, 'Exact']]);
        });
    }
    for (const opts of [{directCellText:true}, {}, {changeSelection:true, directCellText:true}, {empty:true}]) {
        await test('AI ' + JSON.stringify(opts), opts, {request:'notes'}, (s,e) => {
            assert.ifError(e); assert.equal(s.calls,2); assert.equal(s.writes[0][0],0);
            if (!opts.empty) assert(s.prompts[0].includes('73 visitors'), 'table content missing from prompt');
            assert(ended(s,'Block') && ended(s,'GroupActions'));
        });
    }
    for (const slideNumber of [0,null,'2',-1,1.5,NaN,Infinity]) {
        await test('invalid slideNumber ' + String(slideNumber), {}, {text:'Exact',slideNumber}, (s,e) => {
            rejected(s,e); assert.equal(s.calls,0,'invalid number reached editor API');
        });
    }
    for (const [opts,params] of [[{},{text:'x',slideNumber:99}],[{noSlide:true},{text:'x'}],[{noModel:true},{request:'x'}],[{},{}],[{},{text:'x',request:'y'}]]) {
        await test('rejection ' + JSON.stringify([opts,params]), opts, params, rejected);
    }
    for (const opts of [{reject:true},{groupFails:true}]) {
        await test('action failure ' + JSON.stringify(opts), opts, {request:'x'}, (s,e) => {
            rejected(s,e); assert(ended(s,'Block'));
            if (opts.reject) assert(ended(s,'GroupActions'));
        });
    }
    const s=setup();
    assert(s.tool.examples.some(e=>e.prompt==='generate talking points for this slide' && !('slideNumber' in e.arguments)));
    console.log('Metadata current-slide example: PASS');
    console.log(`RESULT: ${passed} passed, ${failed} failed; failures are retained, not patched in the student source.`);
    process.exitCode = failed ? 1 : 0;
})().catch(e=>{console.error(e);process.exitCode=1;});
