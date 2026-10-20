const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const helperSource=fs.readFileSync(path.join(__dirname,'regressions.test.cjs'),'utf8').split("test('O01")[0];
const {app,generationApp,snapshot}=new Function('require','__dirname',helperSource+';return {app,generationApp,snapshot};')(require,__dirname);

function storageApp(shared){
  const instance=app();
  instance.context.localStorage={getItem:key=>shared.get(key)||null,setItem:(key,value)=>shared.set(key,value)};
  instance.context.setTimeout=setTimeout;instance.context.clearTimeout=clearTimeout;
  instance.run(`captureActivePresetClassesFromDom=()=>{};setCheckedValue=()=>{};renderActivePreset=()=>{};syncPresetUi=()=>{};applyLayoutLockToCards=()=>{};updateSummary=()=>{};showToast=()=>{};`);
  return instance;
}

function inputLocks(){
  let occupied=false;
  return {async request(name,options,task){
    assert.equal(name,'word-lab-input');assert.equal(options.ifAvailable,true);
    if(occupied)return task(null);
    occupied=true;
    try{return await task({name})}finally{occupied=false}
  }};
}

test('R01 a second tab cannot save a stale snapshot; after handoff it restores the latest input',async()=>{
  const shared=new Map(),locks=inputLocks();
  const left=storageApp(shared),right=storageApp(shared);
  for(const tab of [left,right]){tab.context.locks=locks;tab.run('state.inputEditable=false;window.navigator={locks};');}
  assert.equal(await left.run('acquireInputEditing()'),true);
  left.run(`state.presets['mon-wed'].teacherName='latest A';persistState();`);
  assert.equal(await right.run('acquireInputEditing()'),false);
  assert.equal(right.run(`state.presets['tue-thu'].teacherName='stale B';persistState()`),false);
  assert.equal(JSON.parse(shared.get('word-lab-desktop-state-v1')).presets['mon-wed'].teacherName,'latest A');
  left.run('state.inputEditable=false;state.releaseInputLock();');
  await new Promise(setImmediate);
  assert.equal(await right.run('acquireInputEditing()'),true);
  assert.equal(right.run(`state.presets['mon-wed'].teacherName`),'latest A');
  right.run(`state.presets['tue-thu'].teacherName='latest B';persistState();`);
  const saved=JSON.parse(shared.get('word-lab-desktop-state-v1'));
  assert.equal(saved.presets['mon-wed'].teacherName,'latest A');
  assert.equal(saved.presets['tue-thu'].teacherName,'latest B');
  right.run('state.releaseInputLock()');
});

test('R01 read-only enforcement leaves backup/output available and preserves disabled states',()=>{
  const {context,run,ids}=app();
  const controls=['teacher-name','generate-btn','pick-folder-btn','export-snapshot-btn','acquire-input-btn','session-count-increase'].map(id=>ids.get(id));
  ids.get('session-count-input').value='5';
  ids.get('session-count-increase').disabled=true;
  context.document.querySelectorAll=()=>controls;
  run('state.inputEditable=false;syncInputEditingUi();');
  assert.equal(ids.get('teacher-name').disabled,true);
  for(const id of ['generate-btn','pick-folder-btn','export-snapshot-btn','acquire-input-btn'])assert.equal(ids.get(id).disabled,false,id);
  run('state.inputEditable=true;syncInputEditingUi();');
  assert.equal(ids.get('teacher-name').disabled,false);
  assert.equal(ids.get('session-count-increase').disabled,true);
});

test('R01 unsupported coordination fails closed; external storage changes preserve the visible draft',async()=>{
  const {context,run,ids}=app();
  context.window.navigator={};
  run('state.inputEditable=false;');
  assert.equal(await run('acquireInputEditing()'),false);
  assert.equal(run('persistState()'),false);
  assert.equal(ids.get('acquire-input-btn').disabled,true);
  const listeners=new Map();context.window.addEventListener=(name,handler)=>listeners.set(name,handler);
  const controls=[ids.get('teacher-name'),ids.get('export-snapshot-btn')];
  context.document.querySelectorAll=()=>controls;
  run(`state.inputEditable=true;bindGlobalEvents();elements.teacherNameInput.value='live draft';setGenerating(true);`);
  listeners.get('storage')({key:'word-lab-desktop-state-v1'});
  run('setGenerating(false)');
  assert.equal(run('state.hasUnsavedChanges'),true);
  assert.equal(ids.get('teacher-name').value,'live draft');
  assert.equal(ids.get('teacher-name').disabled,true);
  assert.equal(ids.get('export-snapshot-btn').disabled,false);
});

test('R02 the latest local edit survives same-account reload; both original storage keys remain',async()=>{
  const shared=new Map();
  const initial=snapshot();initial.presets['mon-wed'].teacherName='Local before';
  shared.set('word-lab-desktop-state-v1',JSON.stringify(initial));
  const scoped=snapshot();scoped.presets['mon-wed'].teacherName='Admin old';
  const adminKey='word-lab-desktop-state-v1:WORD LAB:admin:synthetic-A';
  shared.set(adminKey,JSON.stringify(scoped));
  const first=storageApp(shared);
  first.run(`state.activeStorageKey=getStateStorageKey();restoreState();initializeCurrentUserContext=()=>new Promise(resolve=>globalThis.resolveAuth=resolve);globalThis.pending=initializeOptionalUserContext();state.presets['mon-wed'].teacherName='Latest edited';persistState();resolveAuth({user:{uid:'synthetic-A'},profile:{role:'admin'}});`);
  await first.run('pending');
  const reloaded=storageApp(shared);
  reloaded.run(`state.activeStorageKey=getStateStorageKey();restoreState();initializeCurrentUserContext=async()=>({user:{uid:'synthetic-A'},profile:{role:'admin'}});`);
  await reloaded.run('initializeOptionalUserContext()');
  assert.equal(reloaded.run(`state.presets['mon-wed'].teacherName`),'Latest edited');
  assert.equal(reloaded.run('getStateStorageKey()'),'word-lab-desktop-state-v1');
  assert.equal(JSON.parse(shared.get(adminKey)).presets['mon-wed'].teacherName,'Admin old');
  reloaded.run('switchInputScope()');
  assert.equal(reloaded.run(`state.presets['mon-wed'].teacherName`),'Admin old');
  assert.equal(JSON.parse(shared.get('word-lab-desktop-state-v1')).presets['mon-wed'].teacherName,'Latest edited');
  reloaded.run('switchInputScope()');
  assert.equal(reloaded.run(`state.presets['mon-wed'].teacherName`),'Latest edited');
});

test('R02 different accounts, cancelled switching and cross-scope undo cannot overwrite either original',async()=>{
  const shared=new Map(),{context,run}=storageApp(shared);
  run(`state.presets['mon-wed'].teacherName='Local latest';persistState();initializeCurrentUserContext=async()=>({user:{uid:'A'},profile:{role:'admin'}});`);
  await run('initializeOptionalUserContext()');
  run(`prepareStateReplacement('test');`);
  const admin=snapshot();admin.presets['mon-wed'].teacherName='Admin original';
  shared.set('word-lab-desktop-state-v1:WORD LAB:admin:A',JSON.stringify(admin));
  context.window.confirm=()=>false;run('switchInputScope()');
  assert.equal(run('getStateStorageKey()'),'word-lab-desktop-state-v1');
  context.window.confirm=()=>true;run('switchInputScope();undoLastReplacement();');
  assert.equal(run(`state.presets['mon-wed'].teacherName`),'Admin original');
  assert.equal(run('elements.undoNotice.hidden'),true);
  run('switchInputScope()');
  assert.equal(run('elements.undoNotice.hidden'),false);
  const other=storageApp(shared);
  shared.set('word-lab-desktop-state-v1:scope',JSON.stringify({mode:'admin',uid:'A'}));
  other.run(`restoreState();initializeCurrentUserContext=async()=>({user:{uid:'B'},profile:{role:'admin'}});`);
  await other.run('initializeOptionalUserContext()');
  assert.equal(other.run('getStateStorageKey()'),'word-lab-desktop-state-v1');
  assert.equal(shared.has('word-lab-desktop-state-v1:WORD LAB:admin:B'),false);
  assert.equal(JSON.parse(shared.get('word-lab-desktop-state-v1:WORD LAB:admin:A')).presets['mon-wed'].teacherName,'Admin original');
});

test('R02 scope preference write failure is visible; a failed switch leaves the original scope intact',()=>{
  const {context,run}=storageApp(new Map());
  run(`state.availableUserContext={user:{uid:'A'},profile:{role:'admin'}};state.activeStorageKey=getStateStorageKey();state.presets['mon-wed'].teacherName='Keep';`);
  context.localStorage.setItem=(key)=>{if(key.endsWith(':scope'))throw new Error('quota')};
  assert.equal(run('persistState()'),false);
  assert.equal(run('state.hasUnsavedChanges'),true);
  run('state.hasUnsavedChanges=false;switchInputScope()');
  assert.equal(run('getStateStorageKey()'),'word-lab-desktop-state-v1');
  assert.equal(run(`state.presets['mon-wed'].teacherName`),'Keep');
});

test('R03 raw out-of-range, fractional and empty question counts stop before permission/output',async()=>{
  for(const value of ['3','101','5.5','','NaN']){
    const {run}=generationApp();
    run(`elements.questionCountInput.value=${JSON.stringify(value)};globalThis.permissions=0;ensureWritableSaveRootHandle=async()=>{permissions++;return true};`);
    await run('handleGenerate()');
    assert.equal(run('calls.targets'),0,value);
    assert.equal(run('permissions'),0,value);
    assert.equal(run('elements.questionCountInput.getAttribute("aria-invalid")'),'true',value);
    assert.equal(run('elements.questionCountInput.focused'),true,value);
  }
});

test('R03 correcting invalid input clears ARIA errors; 5 and 100 each generate the visible count',async()=>{
  const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');
  for(const count of [5,100]){
    const {run}=generationApp();
    run(source.match(/function getSettings\([^]*?(?=\nfunction )/)[0]);
    run(`elements.questionCountInput.value='3';setValidationNotes([getQuestionCountError()]);elements.questionCountInput.value='${count}';state.presets['mon-wed'].classes[0].rawText=Array.from({length:${count}},(_,i)=>'word'+i+'\\t뜻'+i).join('\\n');globalThis.printedCount=0;buildSheetPages=options=>{printedCount=options.items.length;return []};`);
    await run('handleGenerate()');
    assert.equal(run('printedCount'),count);
    assert.equal(run('elements.questionCountInput.getAttribute("aria-invalid")'),null);
    assert.equal(run('elements.questionCountValidation.hidden'),true);
  }
});

test('R04 late initialization cannot replace generation or terminal status',async()=>{
  for(const message of ['PDF 1/2개 저장됨 · 생성 중','생성 완료','일부 저장 실패']){
    const {run}=app();
    run(`cacheElements=()=>{};initializeOutputDateControl=()=>{};initializeNativeBridge=()=>{};restoreState=()=>{};renderActivePreset=()=>{};bindGlobalEvents=()=>{};syncPresetUi=()=>{};syncSaveRootUi=()=>{};acquireInputEditing=async()=>true;initializeOptionalUserContext=async()=>{};restoreSaveRootHandle=()=>new Promise(resolve=>globalThis.resumeRestore=resolve);updateOutputFolderPreview=async()=>{};updateSummary=()=>{};globalThis.pending=initializeApp();`);
    await Promise.resolve();await Promise.resolve();
    run(`setRunStatus(${JSON.stringify(message)},'running');resumeRestore();`);
    await run('pending');
    assert.equal(run('elements.runStatus.textContent'),message);
  }
});

test('V01 late and repeated auth preserves the backup draft after storage replacement, deletion or clear',async()=>{
  for(const eventKey of ['word-lab-desktop-state-v1',null]){
    for(const remove of [false,true]){
      const shared=new Map(),{context,run}=storageApp(shared),listeners=new Map();
      context.window.addEventListener=(name,handler)=>listeners.set(name,handler);
      const local=snapshot(),admin=snapshot();
      local.presets['mon-wed'].teacherName='Backup this draft';
      admin.presets['mon-wed'].teacherName='Admin original';
      const adminKey='word-lab-desktop-state-v1:WORD LAB:admin:A';
      shared.set('word-lab-desktop-state-v1',JSON.stringify(local));
      shared.set(adminKey,JSON.stringify(admin));
      shared.set('word-lab-desktop-state-v1:scope',JSON.stringify({mode:'admin',uid:'A'}));
      run(`restoreState();bindGlobalEvents();initializeCurrentUserContext=()=>new Promise(resolve=>globalThis.finishAuth=resolve);globalThis.pending=initializeOptionalUserContext();`);
      if(remove)shared.delete('word-lab-desktop-state-v1');
      else shared.set('word-lab-desktop-state-v1','external replacement');
      listeners.get('storage')({key:eventKey});
      run(`finishAuth({user:{uid:'A'},profile:{role:'admin'}})`);
      await run('pending');
      run(`initializeCurrentUserContext=async()=>({user:{uid:'A'},profile:{role:'admin'}})`);
      await run('initializeOptionalUserContext()');
      assert.equal(run(`buildSnapshotPayload().presets['mon-wed'].teacherName`),'Backup this draft');
      assert.equal(run('getStateStorageKey()'),'word-lab-desktop-state-v1');
      assert.equal(run('state.inputEditable'),false);
      assert.equal(run('state.inputConflict'),true);
      assert.equal(run('state.hasUnsavedChanges'),true);
      assert.equal(run('elements.storageWarning.hidden'),false);
      assert.equal(shared.get(adminKey),JSON.stringify(admin));
    }
  }
});

test('V01 authentication never reloads an unsaved draft in its existing admin scope',async()=>{
  const shared=new Map(),{run}=storageApp(shared);
  const original=snapshot();original.presets['mon-wed'].teacherName='Saved admin';
  shared.set('word-lab-desktop-state-v1:WORD LAB:admin:A',JSON.stringify(original));
  shared.set('word-lab-desktop-state-v1:scope',JSON.stringify({mode:'admin',uid:'A'}));
  run(`state.currentUser={uid:'A'};state.userProfile={role:'admin'};state.activeStorageKey=getStateStorageKey(state.currentUser,state.userProfile);state.presets['mon-wed'].teacherName='Unsaved admin draft';state.hasUnsavedChanges=true;initializeCurrentUserContext=async()=>({user:{uid:'A'},profile:{role:'admin'}});`);
  await run('initializeOptionalUserContext()');
  assert.equal(run(`buildSnapshotPayload().presets['mon-wed'].teacherName`),'Unsaved admin draft');
  assert.equal(run('state.hasUnsavedChanges'),true);
});

function folderApp(){
  const instance=storageApp(new Map());
  instance.context.window.indexedDB={};
  instance.run(`syncSaveRootUi=()=>{};persistSaveRootHandle=async handle=>globalThis.storedFolder=handle;globalThis.clears=0;clearStoredSaveRootHandle=async()=>clears++;globalThis.newFolder={name:'New folder'};`);
  return instance;
}

test('V02 a late startup handle read or failure cannot replace or clear a newer folder choice',async()=>{
  for(const fail of [false,true]){
    const {run}=folderApp();
    run(`getStoredSaveRootHandle=()=>new Promise((resolve,reject)=>globalThis.finishRead=${fail?'reject':'resolve'});globalThis.restoring=restoreSaveRootHandle();`);
    await run('activateSaveRootHandle(newFolder)');
    run(fail?`finishRead(new Error('old read failed'))`:`finishRead({name:'Old folder',queryPermission:async()=> 'granted'})`);
    await run('restoring');
    assert.equal(run('state.saveRootHandle'),run('newFolder'));
    assert.equal(run('state.rememberedSaveRootHandle'),run('newFolder'));
    assert.equal(run('state.saveRootName'),'New folder');
    assert.equal(run('storedFolder'),run('newFolder'));
    assert.equal(run('clears'),0);
  }
});

test('V02 a late startup permission result or failure cannot replace the newer selected folder',async()=>{
  for(const outcome of ['granted','denied','error']){
    const {run}=folderApp();
    run(`globalThis.oldFolder={name:'Old folder',queryPermission:()=>new Promise((resolve,reject)=>globalThis.finishPermission=${outcome==='error'?'reject':'resolve'})};getStoredSaveRootHandle=async()=>oldFolder;globalThis.restoring=restoreSaveRootHandle();`);
    await Promise.resolve();
    assert.equal(run('typeof finishPermission'),'function');
    await run('activateSaveRootHandle(newFolder)');
    run(outcome==='error'?`finishPermission(new Error('old permission failed'))`:`finishPermission('${outcome}')`);
    await run('restoring');
    assert.equal(run('state.saveRootHandle'),run('newFolder'));
    assert.equal(run('state.rememberedSaveRootHandle'),run('newFolder'));
    assert.equal(run('state.saveRootName'),'New folder');
    assert.equal(run('clears'),0);
  }
});

test('V02 delayed old-handle cleanup does not delete or disconnect a folder chosen while the database opens',async()=>{
  for(const reconnect of [false,true]){
    const {run}=folderApp();
    const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');
    run(source.match(/async function clearStoredSaveRootHandle\([^]*?(?=\n(?:async )?function )/)[0]);
    run(`globalThis.deletes=0;globalThis.closed=0;openHandleDatabase=()=>new Promise(resolve=>globalThis.opened=resolve);getStoredSaveRootHandle=async()=>{throw new Error('bad old handle')};state.rememberedSaveRootHandle={queryPermission:async()=>{throw new Error('bad old permission')}};globalThis.restoring=${reconnect?'restoreRememberedSaveRoot(true)':'restoreSaveRootHandle()'};`);
    await new Promise(setImmediate);
    assert.equal(run('typeof opened'),'function');
    await run('activateSaveRootHandle(newFolder)');
    run(`opened({close(){closed++},transaction(){const tx={objectStore:()=>({delete(){deletes++;return {}}})};Promise.resolve().then(()=>tx.oncomplete?.());return tx}})`);
    const usable=await run('restoring');
    if(reconnect)assert.equal(usable,true);
    assert.equal(run('state.saveRootHandle'),run('newFolder'));
    assert.equal(run('storedFolder'),run('newFolder'));
    assert.equal(run('deletes'),0);
    assert.equal(run('closed'),1);
  }
});

test('V02 an older folder save waiting for the database cannot replace the newer saved selection',async()=>{
  const {run}=folderApp();
  const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');
  run(source.match(/async function persistSaveRootHandle\([^]*?(?=\n(?:async )?function )/)[0]);
  run(`globalThis.opens=[];globalThis.storedName='';globalThis.closed=0;openHandleDatabase=()=>new Promise(resolve=>opens.push(resolve));globalThis.first=activateSaveRootHandle({name:'Old choice'});globalThis.second=activateSaveRootHandle(newFolder);globalThis.database={close(){closed++},transaction(){const tx={objectStore:()=>({put(handle){storedName=handle.name;return {}}})};Promise.resolve().then(()=>tx.oncomplete?.());return tx}};opens[1](database);`);
  await run('second');
  run('opens[0](database)');await run('first');
  assert.equal(run('state.saveRootName'),'New folder');
  assert.equal(run('storedName'),'New folder');
  assert.equal(run('closed'),2);
});

test('V02 a stale reconnect result or failure preserves the newer folder and reports it usable',async()=>{
  for(const outcome of ['granted','denied','error']){
    const {run}=folderApp();
    run(`state.rememberedSaveRootHandle={name:'Old folder',queryPermission:()=>new Promise((resolve,reject)=>globalThis.finishPermission=${outcome==='error'?'reject':'resolve'}),requestPermission:async()=> 'denied'};globalThis.restoring=restoreRememberedSaveRoot(true);`);
    await run('activateSaveRootHandle(newFolder)');
    run(outcome==='error'?`finishPermission(new Error('old reconnect failed'))`:`finishPermission('${outcome}')`);
    assert.equal(await run('restoring'),true);
    assert.equal(run('state.saveRootHandle'),run('newFolder'));
    assert.equal(run('state.rememberedSaveRootHandle'),run('newFolder'));
    assert.equal(run('clears'),0);
  }
});

test('V02 ordinary startup restoration still connects granted handles and keeps denied handles reconnectable',async()=>{
  for(const permission of ['granted','prompt','denied']){
    const {run}=folderApp();
    run(`globalThis.oldFolder={name:'Old folder',queryPermission:async()=> '${permission}'};getStoredSaveRootHandle=async()=>oldFolder;`);
    await run('restoreSaveRootHandle()');
    assert.equal(run('state.rememberedSaveRootHandle'),run('oldFolder'));
    assert.equal(run('state.saveRootName'),'Old folder');
    assert.equal(run('state.saveRootHandle'),permission==='granted'?run('oldFolder'):null);
    assert.equal(run('clears'),0);
  }
  const {run}=folderApp();
  run(`getStoredSaveRootHandle=async()=>{throw new Error('invalid old handle')};`);
  await run('restoreSaveRootHandle()');
  assert.equal(run('clears'),1);
  assert.equal(run('state.saveRootHandle'),null);
});

test('V02 a late native startup response cannot overwrite a completed explicit native folder choice',async()=>{
  for(const fail of [false,true]){
    const {run}=folderApp();
    run(`hasNativeHost=()=>true;callNativeHost=method=>method==='chooseSaveRoot'?Promise.resolve({remembered:true,connected:true,name:'New native folder'}):new Promise((resolve,reject)=>globalThis.finishStatus=${fail?'reject':'resolve'});globalThis.restoring=restoreSaveRootHandle();`);
    await run('chooseNativeSaveRoot()');
    run(fail?`finishStatus(new Error('old native read failed'))`:`finishStatus({remembered:true,connected:true,name:'Old native folder'})`);
    await run('restoring');
    assert.equal(run('state.saveRootName'),'New native folder');
    assert.equal(run('state.saveRootHandle.native'),true);
  }
});

test('V02 output after an overlapping startup restore writes only to the newer folder and keeps existing ZIPs',async()=>{
  const {context,run}=generationApp();
  const stores=[new Map([['old.txt','KEEP OLD FILE']]),new Map([['Teacher_2030-01-02.zip','KEEP EXISTING ZIP']])];
  function folder(store,prefix='',directories=new Set([''])){
    return {name:prefix||'Root',queryPermission:async()=> 'granted',
      async getDirectoryHandle(name,options={}){const key=prefix?prefix+'/'+name:name;if(!directories.has(key)&&!options.create)throw {name:'NotFoundError'};directories.add(key);return folder(store,key,directories)},
      async getFileHandle(name,options={}){const key=prefix?prefix+'/'+name:name;if(!store.has(key)&&!options.create)throw {name:'NotFoundError'};return {createWritable:async()=>{let pending;return {write:async value=>pending=value,close:async()=>store.set(key,pending),abort:async()=>{}}}}}
    };
  }
  context.oldRoot=folder(stores[0]);context.newRoot=folder(stores[1]);context.window.indexedDB={};
  context.window.html2canvas={};context.window.jspdf={jsPDF:{}};
  const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');
  for(const name of ['ensureWritableSaveRootHandle','resolveAvailableOutputFolderName','createGenerationTarget','writeGeneratedPdfFile','writeGeneratedTextFile','writeGeneratedZipFile']){
    run(source.match(new RegExp('async function '+name+'\\([^]*?(?=\\n(?:async )?function )'))[0]);
  }
  run(`syncSaveRootUi=()=>{};persistSaveRootHandle=async()=>{};getStoredSaveRootHandle=()=>new Promise(resolve=>globalThis.finishRead=resolve);globalThis.restoring=restoreSaveRootHandle();`);
  await run('activateSaveRootHandle(newRoot)');
  run(`newRoot.queryPermission=async()=>{finishRead(oldRoot);await restoring;return 'granted'};`);
  await run('handleGenerate()');
  assert.deepEqual(Array.from(stores[0]),[['old.txt','KEEP OLD FILE']]);
  assert.equal(stores[1].get('Teacher_2030-01-02.zip'),'KEEP EXISTING ZIP');
  assert.equal(Array.from(stores[1].keys()).filter(name=>name.endsWith('.pdf')).length,2);
  assert.equal(stores[1].has('Teacher_2030-01-02_2.zip'),true);
  assert.equal(stores[1].has('Teacher_2030-01-02_2/WORD_LAB_단어복원.json'),true);
  assert.equal(run('calls.statuses.at(-1).type'),'success');
});
