const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

function setup(opts = {}) {
    const writes = [], actions = [], prompts = [], logs = [];
    let selected = 0, calls = 0;
    const slides = [0, 1].map(index => ({
        GetSlideIndex: () => index,
        AddNotesText: opts.noApi ? undefined : text => { if (opts.writeFail) return false; writes.push([index, text]); return true; },
        GetAllShapes: () => opts.empty ? [] : [{ GetDocContent: () => ({ GetElementsCount: () => 1, GetElement: () => ({ GetText: () => 'Library pilot' }) }) }],
        GetAllTables: () => opts.empty ? [] : [{ GetRow: row => row > 0 ? null : ({
            GetCellsCount: () => 1,
            GetCell: () => opts.directCellText ? { GetText: () => '73 visitors' } : { GetContent: () => ({ GetText: () => '73 visitors' }) }
        }) }]
    }));
    const c = {
        console: { log: (...a) => logs.push(a), warn: (...a) => logs.push(a), error: (...a) => logs.push(a) },
        RegisteredFunction: function (o) { Object.assign(this, o); },
        Asc: { scope: {}, Editor: {} }, window: { AgentState: { ToolError: Error } },
        Api: { GetPresentation: () => ({ GetCurrentSlide: () => opts.noSlide ? null : slides[selected], GetSlideByIndex: i => slides[i], GetSlidesCount: () => 2 }) },
        AI: { ActionType: { Chat: 1 }, Request: { create: () => opts.noModel ? null : ({
            modelUI: { name: 'test' }, chatRequest: async prompt => {
                prompts.push(prompt);
                if (opts.changeSelection) selected = 1;
                if (opts.reject) throw Error('offline');
                return opts.emptyResponse ? '' : 'Generated notes';
            }
        }) } }
    };
    vm.createContext(c);
    c.Asc.Editor.callCommand = async fn => { calls++; return vm.runInContext('(' + fn.toString() + ')()', c); };
    c.Asc.Editor.callMethod = async (name, args) => {
        actions.push([name, args[0]]);
        if (opts.groupFails && name === 'StartAction' && args[0] === 'GroupActions') throw Error('group failed');
    };
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../../scripts/helpers/helpers.js'), 'utf8'), c);
    const tool = c.HELPERS.slide.flat().find(f => f.name === 'addNoteToSlide');
    assert.equal(c.HELPERS.names.slide.addNoteToSlide, 'Insert Note');
    assert(tool.examples.some(e => e.prompt === 'generate talking points for this slide' && e.arguments.slideNumber === undefined));
    return { tool, c, writes, actions, prompts, logs, get calls() { return calls; } };
}

(async () => {
    let count = 0;
    for (const slideNumber of [undefined, 2]) {
        const s = setup();
        await s.tool.call({ text: 'Exact wording', slideNumber });
        assert.equal(s.calls, 1);
        assert.equal(s.prompts.length, 0);
        assert.deepEqual(s.writes, [[slideNumber === 2 ? 1 : 0, 'Exact wording']]);
        count++;
    }
    for (const opts of [{}, { directCellText: true }, { changeSelection: true }, { empty: true }]) {
        const s = setup(opts);
        await s.tool.call({ request: 'Generate talking points for this slide' });
        assert.equal(s.calls, 2);
        assert.equal(s.writes[0][0], 0);
        if (!opts.empty) assert(s.prompts[0].includes('73 visitors') && s.prompts[0].includes('Library pilot'));
        assert.equal(Object.keys(s.c.Asc.scope).length, 0);
        count++;
    }
    for (const slideNumber of [null, 0, -1, 1.5, '2', 99]) {
        const s = setup();
        await assert.rejects(() => s.tool.call({ text: 'x', slideNumber }));
        assert.equal(s.writes.length, 0);
        count++;
    }
    for (const [opts, params] of [
        [{}, {}], [{}, { text: 'x', request: 'y' }], [{}, { text: ' ' }],
        [{ noSlide: true }, { text: 'x' }], [{ noApi: true }, { text: 'x' }],
        [{ writeFail: true }, { text: 'x' }], [{ noModel: true }, { request: 'x' }],
        [{ reject: true }, { request: 'x' }], [{ groupFails: true }, { request: 'x' }],
        [{ emptyResponse: true }, { request: 'x' }]
    ]) {
        const s = setup(opts);
        await assert.rejects(() => s.tool.call(params));
        assert.equal(s.writes.length, 0);
        assert.equal(Object.keys(s.c.Asc.scope).length, 0);
        if (opts.reject || opts.groupFails) assert(s.actions.some(a => a[0] === 'EndAction' && a[1] === 'Block'));
        if (opts.reject) assert(s.actions.some(a => a[0] === 'EndAction' && a[1] === 'GroupActions'));
        count++;
    }
    console.log(`PASS: ${count} slide-note regression scenarios`);
})().catch(error => { console.error(error); process.exitCode = 1; });
