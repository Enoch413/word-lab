const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname, '../app/app.js'), 'utf8');

function element() {
  const attributes = new Map();
  return {
    handlers: new Map(),
    value: '', textContent: '', innerHTML: '', hidden: false, disabled: false,
    className: '', dataset: {},
    classList: { add() {}, remove() {}, toggle() {} },
    setAttribute(key, value) { attributes.set(key, String(value)); },
    getAttribute(key) { return attributes.get(key) ?? null; },
    removeAttribute(key) { attributes.delete(key); },
    focus() { this.focused = true; },
    addEventListener(type, handler) { this.handlers.set(type, handler); },
    querySelectorAll() { return []; },
  };
}

function app() {
  const ids = new Map();
  const context = vm.createContext({
    document: {
      addEventListener() {},
      getElementById(id) { if (!ids.has(id)) ids.set(id, Object.assign(element(),{id})); return ids.get(id); },
      querySelectorAll() { return []; },
    },
    window: { confirm: () => true, addEventListener() {}, clearTimeout() {}, setTimeout() {}, navigator:{locks:{request:async(name,options,task)=>task()}} },
    localStorage: { setItem() {} },
    TextEncoder, TextDecoder, Blob, console: { error() {} },
  });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  run('cacheElements();state.inputEditable=true;');
  return { context, run, ids };
}

function snapshot() {
  return {
    appName: 'WORD LAB', version: 1, activePresetId: 'mon-wed',
    settings: { questionCount: 5, sessionCount: 1 },
    presets: {
      'mon-wed': { teacherName: '선생님', layoutLocked: false, classes: [{ className: 'Original', rawText: 'apple\t사과', quizMode: 'word-to-meaning' }] },
      'tue-thu': { teacherName: '', layoutLocked: false, classes: [] },
    },
  };
}

function generationApp() {
  const instance = app();
  instance.context.fixtureSnapshot = snapshot();
  instance.run(`
    globalThis.calls = {targets:0, pdfClasses:[], snapshots:[], zipEntries:[], statuses:[], toasts:[]};
    elements.questionCountInput.value='5';elements.sessionCountInput.value='1';
    captureActivePresetClassesFromDom = () => {};
    state.presets = fixtureSnapshot.presets;
    getSettings = () => ({questionCount:5, sessionCount:1});
    collectClassModels = () => [{className:'Original', quizMode:'word-to-meaning', parsed:parseVocabulary('apple\\t사과')}];
    ensureWritableSaveRootHandle = async () => true;
    getBaseOutputFolderName = () => 'Teacher_2030-01-02';
    resolveAvailableOutputFolderName = async name => name;
    createGenerationTarget = async () => { calls.targets++; return {}; };
    buildSheetPages = options => { calls.pdfClasses.push(options.className); return []; };
    buildAnswerSheetPages = () => [];
    renderPagesToPdfBuffer = async () => new ArrayBuffer(2);
    writeGeneratedPdfFile = async () => {};
    writeGeneratedTextFile = async (target, name, text) => calls.snapshots.push(JSON.parse(text));
    createZipArchive = entries => { calls.zipEntries = entries; return new ArrayBuffer(2); };
    writeGeneratedZipFile = async () => {};
    setRunStatus = (text, type) => calls.statuses.push({text,type});
    showToast = text => calls.toasts.push(text);
    updateOutputFolderPreview = async () => {};
  `);
  return instance;
}

test('O01 output progress identifies the class, session and successfully saved PDF count', async () => {
  const {run}=generationApp();
  await run('handleGenerate()');
  const statuses=Array.from(run('calls.statuses'),item=>item.text);
  assert.ok(statuses.some(text=>/PDF 0\/2개 저장됨 · Original 1차 시험지 생성 중/.test(text)));
  assert.ok(statuses.some(text=>/PDF 1\/2개 저장됨 · Original 1차 답지 생성 중/.test(text)));
  assert.ok(statuses.some(text=>/PDF 2\/2개 저장됨 · 복원 JSON·ZIP 저장 중/.test(text)));
  assert.match(statuses.at(-1),/생성 완료/);
});

test('F01 rejects unrelated/malformed snapshots before replacing any state', async () => {
  const { run, context } = app();
  const invalid = [null, [], {}, { ...snapshot(), appName:'Other' }, { ...snapshot(), version:2 }];
  const malformed = snapshot(); malformed.presets['mon-wed'].classes = [null]; invalid.push(malformed);
  const badText = snapshot(); badText.presets['mon-wed'].classes[0].rawText = {}; invalid.push(badText);
  const badSettings = snapshot(); badSettings.settings.questionCount = 'five'; invalid.push(badSettings);
  for (const value of invalid) {
    context.candidate = value;
    assert.throws(() => run('normalizeSnapshotImport(candidate)'));
  }
  context.fixtureSnapshot = snapshot();
  run(`state.presets = fixtureSnapshot.presets; globalThis.saves=0; persistState=()=>saves++; showToast=()=>{};`);
  await run(`handleSnapshotImport({target:{files:[{text:async()=> '{}'}]}})`);
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), 'Original');
  assert.equal(run('saves'), 0);
});

test('F01 supports current v1 and existing v1 files without optional fields', () => {
  const { run, context } = app();
  context.candidate = snapshot();
  delete context.candidate.presets['mon-wed'].teacherName;
  delete context.candidate.presets['mon-wed'].layoutLocked;
  const normalized = run('normalizeSnapshotImport(candidate)');
  assert.equal(normalized.presets['mon-wed'].classes[0].rawText, 'apple\t사과');
});

test('F01 cancelled import/demo leave current data untouched; confirmed replacement can be undone', async () => {
  const { run, context } = app();
  context.fixtureSnapshot = snapshot();
  run(`
    state.presets = fixtureSnapshot.presets;
    captureActivePresetClassesFromDom=()=>{};
    getSettings=()=>({questionCount:5,sessionCount:1});
    renderActivePreset=()=>{}; syncPresetUi=()=>{}; setCheckedValue=()=>{};
    applySettings=()=>{}; persistState=()=>true; updateSummary=()=>{}; showToast=()=>{};
    window.confirm=()=>false;
  `);
  await run(`handleSnapshotImport({target:{files:[{text:async()=>JSON.stringify({...fixtureSnapshot,presets:{...fixtureSnapshot.presets,'mon-wed':{classes:[]}}})}]}})`);
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), 'Original');
  run('seedDemoData();');
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), 'Original');
  run('window.confirm=()=>true; seedDemoData();');
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), '2강서A');
  run('undoLastReplacement();');
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), 'Original');
});

test('F02 quota failure is returned and remains visible until a successful retry', () => {
  const { run, ids } = app();
  run(`localStorage.setItem=()=>{throw new Error('QuotaExceededError')};`);
  assert.equal(run('setStoredStateRaw("new data")'), false);
  run('captureActivePresetClassesFromDom=()=>{}; getSettings=()=>({questionCount:5,sessionCount:1});');
  assert.equal(run('persistState()'), false);
  assert.equal(run('state.hasUnsavedChanges'), true);
  assert.equal(ids.get('storage-warning').hidden, false);
  run('localStorage.setItem=()=>{};');
  assert.equal(run('persistState()'), true);
  assert.equal(run('state.hasUnsavedChanges'), false);
  assert.equal(ids.get('storage-warning').hidden, true);
});

test('F02 JSON failure is reported as partial completion; ZIP still contains recovery JSON', async () => {
  const { run } = generationApp();
  run(`writeGeneratedTextFile=async()=>{throw new Error('snapshot write failed')};`);
  await run('handleGenerate()');
  assert.notEqual(run('calls.statuses.at(-1).type'), 'success');
  assert.match(run('calls.statuses.at(-1).text'), /복원.*실패/);
  assert.equal(run('calls.zipEntries.some(entry=>entry.path.endsWith(SNAPSHOT_FILE_NAME))'), true);
});

test('F02 ZIP failure and PDF failure preserve visible partial results', async () => {
  let instance = generationApp();
  instance.run(`writeGeneratedZipFile=async()=>{throw new Error('zip write failed')};`);
  await instance.run('handleGenerate()');
  assert.notEqual(instance.run('calls.statuses.at(-1).type'), 'success');
  assert.match(instance.run('calls.statuses.at(-1).text'), /ZIP.*실패/);
  instance = generationApp();
  instance.run(`globalThis.writes=0; writeGeneratedPdfFile=async()=>{if(++writes===2) throw new Error('disk full')};`);
  await instance.run('handleGenerate()');
  assert.match(instance.run('calls.statuses.at(-1).text'), /1개/);
  assert.match(instance.run('calls.statuses.at(-1).text'), /Teacher_2030-01-02/);
  assert.equal(instance.run('state.isGenerating'),false);
});

test('F02 aborted file writes release the writable and retain the original error', async () => {
  const { run } = app();
  run(`globalThis.aborts=0; globalThis.failure=new Error('write failed'); globalThis.folder={getFileHandle:async()=>({createWritable:async()=>({write:async()=>{throw failure},close:async()=>{},abort:async()=>{aborts++}})})};`);
  await assert.rejects(run('writeBinaryFile(folder,"test.pdf",new ArrayBuffer(1))'), /write failed/);
  assert.equal(run('aborts'), 1);
});

test('F03 two clicks while permission is pending start only one job; cancellation unlocks', async () => {
  const { run } = generationApp();
  run(`globalThis.permissionResolvers=[]; ensureWritableSaveRootHandle=()=>new Promise(resolve=>permissionResolvers.push(resolve)); globalThis.first=handleGenerate(); globalThis.second=handleGenerate(); permissionResolvers.forEach(resolve=>resolve(false));`);
  await run('Promise.all([first,second])');
  assert.equal(run('permissionResolvers.length'), 1);
  assert.equal(run('calls.targets'), 0);
  assert.equal(run('state.isGenerating'), false);
  run('ensureWritableSaveRootHandle=async()=>true;');
  await run('handleGenerate()');
  assert.equal(run('calls.targets'), 1);
  run(`permissionResolvers=[]; ensureWritableSaveRootHandle=()=>new Promise(resolve=>permissionResolvers.push(resolve)); first=handleGenerate(); second=handleGenerate(); permissionResolvers.forEach(resolve=>resolve(true));`);
  await run('Promise.all([first,second])');
  assert.equal(run('calls.targets'),2);
  assert.equal(run('permissionResolvers.length'),1);
});

test('F03 PDF and recovery JSON use the same input captured before permission/render awaits', async () => {
  const { run } = generationApp();
  run(`
    ensureWritableSaveRootHandle=async()=>{state.presets['mon-wed'].classes[0].className='Changed while awaiting permission'; return true};
    renderPagesToPdfBuffer=async()=>{state.presets['mon-wed'].classes[0].rawText='book\\t책'; return new ArrayBuffer(2)};
  `);
  await run('handleGenerate()');
  assert.equal(run('calls.pdfClasses[0]'), 'Original');
  assert.equal(run('calls.snapshots[0].presets["mon-wed"].classes[0].className'), 'Original');
  assert.equal(run('calls.snapshots[0].presets["mon-wed"].classes[0].rawText'), 'apple\t사과');
  assert.equal(run('calls.statuses.at(-1).type'), 'success');
});

test('F03 all editing controls are disabled and previous disabled states are restored', () => {
  const { run, context } = app();
  const controls = [element(), element(), element()]; controls[1].disabled = true;
  context.controls = controls;
  run('document.querySelectorAll=()=>controls; setGenerating(true);');
  assert.deepEqual(controls.map(control=>control.disabled), [true,true,true]);
  run('setGenerating(false);');
  assert.deepEqual(controls.map(control=>control.disabled), [false,true,false]);
});

test('F04 detailed validation is wired to the page and the correct card field', () => {
  const { run, context, ids } = app();
  const name = element(), text = element(), note = element(); note.id='card-validation';
  const card = { querySelector: selector => ({'.class-name-input':name,'.paste-area':text,'.class-validation':note})[selector] };
  context.cards = [card];
  run(`collectClassCards=()=>cards;`);
  run(`setValidationNotes([{type:'error',text:'반 이름이 비어 있습니다.',classIndex:0,field:'className'}]);`);
  assert.match(ids.get('validation-notes').innerHTML, /반 이름이 비어/);
  assert.equal(note.hidden, false);
  assert.equal(name.getAttribute('aria-invalid'), 'true');
  assert.equal(name.getAttribute('aria-describedby'), 'card-validation');
  run('setValidationNotes([]);');
  assert.equal(note.hidden, true);
  assert.equal(name.getAttribute('aria-invalid'), null);
  const html=fs.readFileSync(path.join(__dirname,'../app/index.html'),'utf8');
  assert.match(html, /id="validation-notes"/);
});

test('F04 validation identifies missing names, duplicate folder names, empty words and shortages', () => {
  const { run, context } = app();
  context.models = [
    {className:'',classIndex:1,parsed:{entries:[{word:'a',meaning:'x'}]}},
    {className:'A',classIndex:2,parsed:{entries:[{word:'a',meaning:'x'}]}},
    {className:'a',classIndex:3,parsed:{entries:[{word:'a',meaning:'x'}]}},
    {className:'Empty',classIndex:4,parsed:{entries:[]}},
  ];
  const messages=run('validateClassModels(models,{questionCount:5,sessionCount:3})');
  assert.ok(messages.some(message=>message.type==='error' && message.classIndex===1 && message.field==='className'));
  assert.ok(messages.some(message=>message.type==='error' && message.classIndex===3));
  assert.ok(messages.some(message=>message.type==='warning' && message.classIndex===4 && message.field==='rawText'));
  assert.ok(messages.some(message=>message.text.includes('문항은 1개')));
});

test('F01 a pending import cannot overwrite an in-progress generation', async () => {
  const { run, context } = app();
  context.fixtureSnapshot = snapshot();
  run(`state.presets=fixtureSnapshot.presets; globalThis.readFile=null; globalThis.pending=handleSnapshotImport({target:{files:[{text:()=>new Promise(resolve=>readFile=resolve)}]}}); state.isGenerating=true; readFile(JSON.stringify({...fixtureSnapshot,presets:{...fixtureSnapshot.presets,'mon-wed':{classes:[]}}}));`);
  await run('pending');
  assert.equal(run('state.presets["mon-wed"].classes[0].className'), 'Original');
});

test('F02 manual backup preserves both groups and can be imported when auto-save is unavailable', async () => {
  const { run, context } = app();
  context.fixtureSnapshot=snapshot();
  run(`
    state.presets=fixtureSnapshot.presets;
    captureActivePresetClassesFromDom=()=>{};
    getSettings=()=>({questionCount:5,sessionCount:1});
    globalThis.downloadedBlob=null; globalThis.clicked=false; globalThis.downloadName='';
    globalThis.URL={createObjectURL:blob=>{downloadedBlob=blob;return 'blob:test'},revokeObjectURL:()=>{}};
    document.createElement=()=>({set download(value){downloadName=value},click(){clicked=true}});
    state.hasUnsavedChanges=true;
    exportSnapshot();
  `);
  assert.equal(run('clicked'),true);
  assert.equal(run('downloadName'),'WORD_LAB_단어복원.json');
  const text=await run('downloadedBlob.text()');
  context.candidate=JSON.parse(text);
  assert.equal(run('normalizeSnapshotImport(candidate).presets["mon-wed"].classes[0].rawText'),'apple\t사과');
  assert.equal(run('state.hasUnsavedChanges'),true);
});

test('F02 close failure attempts abort without hiding the original error', async () => {
  const { run } = app();
  run(`globalThis.aborts=0; globalThis.folder={getFileHandle:async()=>({createWritable:async()=>({write:async()=>{},close:async()=>{throw new Error('close failed')},abort:async()=>{aborts++;throw new Error('abort failed')}})})};`);
  await assert.rejects(run('writeTextFile(folder,"test.json","{}")'),/close failed/);
  assert.equal(run('aborts'),1);
});

test('F02 warns before leaving with unsaved input or an active job', () => {
  const { run, context } = app();
  const handlers=new Map();
  context.window.addEventListener=(type,handler)=>handlers.set(type,handler);
  run('bindGlobalEvents();');
  for (const [unsaved,generating,expected] of [[true,false,true],[false,true,true],[false,false,false]]) {
    context.unsaved=unsaved; context.generating=generating;
    run('state.hasUnsavedChanges=unsaved;state.isGenerating=generating;');
    let prevented=false;
    handlers.get('beforeunload')({preventDefault(){prevented=true},returnValue:null});
    assert.equal(prevented,expected);
  }
});

test('F01 the actual delete handler preserves cancelled input and captures a recoverable snapshot', () => {
  const { run, context, ids } = app();
  const cards=[];
  ids.get('class-card-template').content={cloneNode(){
    const fields=Object.fromEntries(['.class-name-input','.paste-area','.class-title','.remove-class-btn','.class-validation'].map(key=>[key,element()]));
    const card={...element(),fields,querySelector:selector=>fields[selector],querySelectorAll:()=>[],remove(){cards.splice(cards.indexOf(card),1)}};
    return {card,querySelector:()=>card};
  }};
  ids.get('class-list').appendChild=card=>cards.push(card);
  context.cards=cards;
  run(`collectClassCards=()=>cards;updateClassCard=()=>{};applyLayoutLockToCards=()=>{};updateSummary=()=>{};showToast=()=>{};getSettings=()=>({questionCount:5,sessionCount:1});addClassCard({className:'Keep',rawText:'apple\\t사과'});addClassCard({className:'Delete',rawText:'book\\t책'});window.confirm=()=>false;`);
  // Exercise the handler attached by addClassCard rather than a copy of deletion logic.
  cards[1].fields['.remove-class-btn'].handlers.get('click')();
  assert.equal(cards.length,2);
  run('window.confirm=()=>true;');
  cards[1].fields['.remove-class-btn'].handlers.get('click')();
  assert.equal(cards.length,1);
  assert.equal(run('state.undoSnapshot.presets["mon-wed"].classes[1].rawText'),'book\t책');
  run('captureActivePresetClassesFromDom=()=>{};renderActivePreset=()=>{};syncPresetUi=()=>{};setCheckedValue=()=>{};applySettings=()=>{};undoLastReplacement();');
  assert.equal(run('state.presets["mon-wed"].classes.length'),2);
});

test('F05 the selected material date is printed on both quiz and answer sheets', async () => {
  const { run }=generationApp();
  run(`elements.outputDateInput.value='2030-01-02';globalThis.dates=[];buildSheetPages=options=>{dates.push(options.dateLabel);return []};buildAnswerSheetPages=options=>{dates.push(options.dateLabel);return []};`);
  await run('handleGenerate()');
  assert.deepEqual(Array.from(run('dates')),['2030.01.02','2030.01.02']);
});
