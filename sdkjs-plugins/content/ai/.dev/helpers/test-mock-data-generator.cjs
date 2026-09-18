const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function setup(options = {}) {
    const writes = [], actions = [], prompts = [], normalized = [];
    let selectionReads = 0;
    const fields = options.fields || ['Name', 'Email'];
    const cell = (r, c) => ({
        GetValue: () => options.occupied && r === 3 && c === fields.length ? options.value : '',
        GetFormula: () => options.formula && r === 3 && c === fields.length ? '=IF(TRUE,"",1)' : '',
        SetValue: value => writes.push([r, c, value])
    });
    const header = {
        GetAddress: () => '$A$1:$B$1',
        GetValue2: () => fields.length === 1 ? fields[0] : [fields],
        Resize: (rows, cols) => ({
            GetRowsCount: () => options.overflow ? rows - 1 : rows,
            GetColumnsCount: () => cols,
            GetRows: r => ({ GetCells: c => cell(r, c) })
        })
    };
    const input = {
        GetColumnsCount: () => fields.length,
        Resize: (rows, cols) => { normalized.push([rows, cols]); return header; }
    };
    const sheet = {
        get Selection() { selectionReads++; return input; },
        GetRange: address => {
            if (address === 'bad') {
                if (options.rangeThrows) throw Error('invalid range');
                return null;
            }
            return address === '$A$1:$B$1' ? header : input;
        }
    };
    const context = {
        console: { log() {}, warn() {}, error() {} },
        RegisteredFunction: function (obj) { Object.assign(this, obj); },
        window: { AgentState: { ToolError: Error } },
        Api: { GetActiveSheet: () => sheet },
        Asc: { scope: {}, Editor: {} },
        AI: { ActionType: { Chat: 1 }, Request: { create: () => ({
            modelUI: { name: 'test' },
            chatRequest: async prompt => {
                prompts.push(prompt);
                if (options.reject) throw Error('provider offline');
                return options.response === undefined ? JSON.stringify([
                    fields.map(() => 'sample 1'), fields.map(() => 'sample 2')
                ]) : options.response;
            }
        }) } }
    };
    vm.createContext(context);
    context.Asc.Editor.callCommand = async fn => vm.runInContext('(' + fn.toString() + ')()', context);
    context.Asc.Editor.callMethod = async (name, args) => {
        actions.push([name, args[0]]);
        if (options.groupFails && name === 'StartAction' && args[0] === 'GroupActions') throw Error('group failed');
    };
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../../scripts/helpers/helpers.js'), 'utf8'), context);
    const tool = context.HELPERS.cell.flat().find(item => item.name === 'mockDataGenerator');
    assert.equal(context.HELPERS.names.cell.mockDataGenerator, 'Generate mock data');
    assert(context.HELPERS.slide.flat().some(item => item.name === 'addNoteToSlide'));
    return { tool, context, writes, actions, prompts, normalized, get selectionReads() { return selectionReads; } };
}

(async () => {
    let count = 0;
    for (const fields of [['Name', 'Email'], ['Name']]) {
        const s = setup({ fields });
        await s.tool.call({ rows: 2 });
        assert.equal(s.writes.length, 2 * fields.length);
        assert.equal(s.normalized[0][0], 1);
        assert(s.prompts[0].includes('inert data'));
        assert.equal(Object.keys(s.context.Asc.scope).length, 0);
        count++;
    }
    for (const rows of [null, 0, -1, 501, 100000, 1.5, '2', NaN]) {
        const s = setup();
        await assert.rejects(() => s.tool.call({ rows }), /rows/);
        assert.equal(s.selectionReads, 0);
        count++;
    }
    for (const rangeThrows of [false, true]) {
        const s = setup({ rangeThrows });
        await assert.rejects(() => s.tool.call({ rows: 2, range: 'bad' }), /invalid/);
        assert.equal(s.selectionReads, 0);
        assert.equal(s.prompts.length, 0);
        count++;
    }
    for (const range of [null, '', 42, {}]) {
        const s = setup();
        await assert.rejects(() => s.tool.call({ rows: 2, range }), /range/);
        assert.equal(s.selectionReads, 0);
        count++;
    }
    for (const options of [
        { occupied: true, value: 'existing' }, { occupied: true, value: ' ' },
        { occupied: true, value: 0 }, { occupied: true, value: false }, { formula: true },
        { overflow: true }, { reject: true }, { groupFails: true },
        { response: 'not JSON' }, { response: '[[{}], [{}]]', fields: ['Name'] }
    ]) {
        const s = setup(options);
        await assert.rejects(() => s.tool.call({ rows: 2 }));
        assert.equal(s.writes.length, 0);
        if (options.reject || options.groupFails) assert(s.actions.some(a => a[0] === 'EndAction' && a[1] === 'Block'));
        if (options.reject) assert(s.actions.some(a => a[0] === 'EndAction' && a[1] === 'GroupActions'));
        assert.equal(Object.keys(s.context.Asc.scope).length, 0);
        count++;
    }
    const s = setup();
    await s.tool.call({ rows: 2, range: 'A1:B6' });
    assert.equal(s.selectionReads, 0);
    assert.equal(s.normalized[0][0], 1);
    console.log(`PASS: ${count + 1} mock-generator regression scenarios`);
})().catch(error => { console.error(error); process.exitCode = 1; });
