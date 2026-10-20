const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');

function app(){
  const context=vm.createContext({document:{addEventListener(){},querySelectorAll(){return []}},window:{},console:{error(){}},setTimeout,clearTimeout,TextEncoder,TextDecoder});
  vm.runInContext(source,context);
  vm.runInContext('state.inputEditable=true',context);
  return {context,run:code=>vm.runInContext(code,context)};
}

test('F06 literal quotes inside a TSV field cannot merge following words',()=>{
  const {run}=app();
  const parsed=run(`parseVocabulary('say\\t말하다 "hello\\nbook\\t책')`);
  assert.deepEqual(Array.from(parsed.entries,entry=>entry.word),['say','book']);
  assert.match(parsed.entries[0].meaning,/"hello/);
});

test('F06 quoted multiline cells, escaped quotes, CRLF and duplicate meanings remain supported',()=>{
  const {run}=app();
  const parsed=run(`parseVocabulary('apple\\t"사과\\r\\n능금"\\r\\nAPPLE\\t"과일 ""apple"""\\r\\nbook\\t책')`);
  assert.equal(parsed.entries.length,2);
  assert.equal(parsed.mergedCount,1);
  assert.equal(parsed.invalidLines.length,0);
  assert.match(parsed.entries[0].meaning,/사과.*능금/);
  assert.match(parsed.entries[0].meaning,/"apple"/);
});

test('F06 unfinished quoted fields are reported with physical line numbers, never silently accepted',()=>{
  const {run}=app();
  const parsed=run(`parseVocabulary('apple\\t"사과\\n능금"\\nbook\\t"닫히지 않음')`);
  assert.equal(parsed.entries.length,1);
  assert.equal(parsed.invalidLines.length,1);
  assert.equal(parsed.invalidLines[0].lineNumber,3);
  assert.match(parsed.invalidLines[0].reason,/따옴표/);
});

test('F07 height-based pagination preserves all entries, order, numbering and page metadata',()=>{
  const {run}=app();
  run(`
    elements.renderStage={innerHTML:'',appendChild:()=>{}};
    buildPrintablePage=options=>({_printOptions:options,classList:{add(){}},querySelectorAll:()=>[],remove(){},getBoundingClientRect(){return {height:Math.max(1123,300+Math.ceil(options.items.length/4)*options.items[0].height)}}});
    globalThis.items=Array.from({length:100},(_,index)=>({word:String(index),height:220}));
    globalThis.input=[buildPrintablePage({items:items.slice(0,50),itemStartIndex:0,totalItemCount:100,className:'검사반'}),buildPrintablePage({items:items.slice(50),itemStartIndex:50,totalItemCount:100,className:'검사반'})];
  `);
  const pages=run('paginatePrintablePages(input)');
  assert.ok(pages.length>1);
  assert.equal(pages.length,9);
  assert.deepEqual(Array.from(pages).flatMap(page=>Array.from(page._printOptions.items,item=>item.word)),Array.from({length:100},(_,i)=>String(i)));
  let offset=0;
  pages.forEach((page,index)=>{
    assert.ok(page.getBoundingClientRect().height<=1123);
    assert.equal(page._printOptions.itemStartIndex,offset);
    assert.equal(page._printOptions.pageNumber,index+1);
    assert.equal(page._printOptions.pageCount,pages.length);
    offset+=page._printOptions.items.length;
  });
});

test('F07 an oversized single item fails explicitly instead of shrinking or clipping text',()=>{
  const {run}=app();
  run(`elements.renderStage={innerHTML:'',appendChild:()=>{}};buildPrintablePage=options=>({_printOptions:options,classList:{add(){}},querySelectorAll:()=>[],remove(){},getBoundingClientRect:()=>({height:3000})});`);
  assert.throws(()=>run(`paginatePrintablePages([buildPrintablePage({items:[{word:'huge'}],itemStartIndex:0,totalItemCount:1})])`),/한 문항|인쇄 영역/);
});

test('S01 local restoration and events do not wait for optional authentication',async()=>{
  const {run}=app();
  run(`
    globalThis.order=[];globalThis.resolveAuth=null;
    cacheElements=()=>{elements.authStatus={hidden:true,textContent:''};elements.inputScopeBtn={hidden:true}};
    initializeOutputDateControl=()=>{};initializeNativeBridge=()=>{};
    initializeCurrentUserContext=()=>new Promise(resolve=>resolveAuth=resolve);
    restoreState=()=>order.push('restored');renderActivePreset=()=>order.push('rendered');
    bindGlobalEvents=()=>order.push('events');syncPresetUi=()=>{};syncSaveRootUi=()=>{};
    restoreSaveRootHandle=async()=>{};updateOutputFolderPreview=async()=>{};updateSummary=()=>{};setRunStatus=()=>{};
    acquireInputEditing=async()=>true;
    globalThis.initialized=initializeApp();
  `);
  try{
    assert.deepEqual(Array.from(run('order')),['restored','rendered','events']);
    await Promise.resolve();
  }finally{
    run('resolveAuth(null);');
    await run('initialized');
  }
});

test('S01 a delayed admin response cannot replace edited input or change its storage scope',async()=>{
  const {run}=app();
  run(`
    elements.authStatus={hidden:true,textContent:''};elements.inputScopeBtn={hidden:true};
    globalThis.resolveAuth=null;initializeCurrentUserContext=()=>new Promise(resolve=>resolveAuth=resolve);
    globalThis.originalKey=getStateStorageKey();globalThis.ready=initializeOptionalUserContext();
    state.localRevision++;state.presets['mon-wed'].teacherName='Edited';
    resolveAuth({user:{uid:'admin-A'},profile:{role:'admin'}});
  `);
  await run('ready');
  assert.equal(run('state.presets["mon-wed"].teacherName'),'Edited');
  assert.equal(run('getStateStorageKey()'),run('originalKey'));
  assert.equal(run('state.currentUser'),null);
});

test('S02 output work uses a cross-tab exclusive lock and rejects unsupported coordination',async()=>{
  const {run}=app();
  run(`globalThis.requested=null;globalThis.tasks=0;window.navigator={locks:{request:async(name,options,task)=>{requested={name,mode:options.mode};return task()}}};`);
  assert.equal(await run('withOutputLock(async()=>{tasks++;return 42})'),42);
  assert.equal(run('requested.mode'),'exclusive');
  assert.equal(run('tasks'),1);
  run('window.navigator={};');
  await assert.rejects(run('withOutputLock(async()=>{tasks++})'),/브라우저|동시/);
  assert.equal(run('tasks'),1);
});

test('S01 synchronous auth callbacks and timeouts both release their subscription',async()=>{
  const {run}=app();
  run(`globalThis.timers=[];globalThis.unsubscribed=0;setTimeout=callback=>{timers.push(callback);return timers.length};clearTimeout=()=>{};`);
  const user=await run(`waitForAuthUser(success=>{success({uid:'test'});return ()=>unsubscribed++})`);
  assert.equal(user.uid,'test');
  assert.equal(run('unsubscribed'),1);
  run(`globalThis.pending=waitForAuthUser(()=>()=>unsubscribed++);timers.at(-1)();`);
  assert.equal(await run('pending'),null);
  assert.equal(run('unsubscribed'),2);
});

test('S02 two tabs sharing an output root choose different reserved names without deleting existing entries',async()=>{
  const directories=new Set();
  const files=new Set(['Teacher_2030-01-02.zip']);
  const root={
    async getDirectoryHandle(name,options){if(options.create){directories.add(name);return {name}}if(directories.has(name))return {name};throw {name:'NotFoundError'}},
    async getFileHandle(name){if(files.has(name))return {name};throw {name:'NotFoundError'}},
  };
  let tail=Promise.resolve();
  const locks={request(name,options,task){const result=tail.then(task);tail=result.catch(()=>{});return result}};
  const tabs=[app(),app()];
  for(const tab of tabs){tab.context.root=root;tab.context.locks=locks;tab.run('state.saveRootHandle=root;window.navigator={locks};');}
  const results=await Promise.all(tabs.map(tab=>tab.run(`withOutputLock(async()=>{const name=await resolveAvailableOutputFolderName('Teacher_2030-01-02');await createGenerationTarget(name);return name})`)));
  assert.deepEqual(results,['Teacher_2030-01-02_2','Teacher_2030-01-02_3']);
  assert.equal(files.has('Teacher_2030-01-02.zip'),true);
  assert.equal(directories.size,2);
});

test('S03 vendor content matches pinned provenance across checkout line endings; jsPDF produces A4 buffers',()=>{
  const vendorDir=process.env.WORDLAB_TEST_VENDOR_DIR || path.join(__dirname,'../app/vendor');
  const manifest=JSON.parse(fs.readFileSync(path.join(vendorDir,'manifest.json'),'utf8'));
  for(const library of manifest.libraries){
    const hash=require('node:crypto').createHash('sha256').update(fs.readFileSync(path.join(vendorDir,library.file),'utf8').replace(/\r\n/g,'\n')).digest('hex');
    assert.equal(hash,library.sha256Lf,library.name);
    assert.match(library.source,/^https:\/\/(raw\.githubusercontent\.com|registry\.npmjs\.org)\//);
  }
  const {jsPDF}=require(path.join(vendorDir,'jspdf.umd.min.js'));
  assert.equal(jsPDF.version,'4.2.1');
  const pdf=new jsPDF({unit:'mm',format:'a4',compress:true});
  pdf.text('regression',10,10);
  assert.ok(pdf.output('arraybuffer').byteLength>1000);
  assert.ok(Math.abs(pdf.internal.pageSize.getWidth()-210)<0.1);
  assert.ok(Math.abs(pdf.internal.pageSize.getHeight()-297)<0.1);
});

test('F10 semantic headings and both radio group names remain in the document',()=>{
  const html=fs.readFileSync(path.join(__dirname,'../app/index.html'),'utf8');
  assert.deepEqual(Array.from(html.matchAll(/<h([1-6])\b/g),match=>Number(match[1])),[1,2,3]);
  assert.match(html,/id="preset-toggle"[^>]*role="group"[^>]*aria-label="반 그룹"/);
  assert.match(html,/class-mode-toggle"[^>]*role="group"[^>]*aria-label="출제 방향"/);
});
