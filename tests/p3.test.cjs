const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const {performance}=require('node:perf_hooks');
const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');

function app(){
  let allocated=0;
  const TrackedBytes=new Proxy(Uint8Array,{construct(target,args){
    if(typeof args[0]==='number') allocated+=args[0];
    return new target(...args);
  }});
  const context=vm.createContext({document:{addEventListener(){}},window:{},console:{error(){}},TextEncoder,TextDecoder,ArrayBuffer,Uint8Array:TrackedBytes,Date});
  vm.runInContext(source,context);
  vm.runInContext('state.inputEditable=true',context);
  return {context,run:code=>vm.runInContext(code,context),reset:()=>{allocated=0},allocated:()=>allocated};
}

function crc32(bytes){
  let crc=0xffffffff;
  for(const byte of bytes){crc^=byte;for(let bit=0;bit<8;bit++)crc=(crc>>>1)^((crc&1)?0xedb88320:0)}
  return (crc^0xffffffff)>>>0;
}

test('O02 ZIP preserves UTF-8 paths, binary slices, empty files, CRC and central directory offsets',()=>{
  const {context,run}=app();
  const files=[
    {path:'자료\\반 A/시험.pdf',data:new Uint8Array([91,0,255,42,92]).subarray(1,4)},
    {path:'자료/복원.json',data:'{"단어":"사과 🍎"}'},
    {path:'자료/빈 파일',data:new Uint8Array(0)},
  ].map(file=>({...file,modifiedAt:new Date(2030,0,2,3,4,6)}));
  context.files=files;
  const zip=Buffer.from(run('createZipArchive(files)'));
  const end=zip.length-22;
  assert.equal(zip.readUInt32LE(end),0x06054b50);
  assert.equal(zip.readUInt16LE(end+10),files.length);
  let central=zip.readUInt32LE(end+16);
  const centralStart=central;
  for(const file of files){
    const expectedName=file.path.replace(/\\/g,'/');
    const expectedData=typeof file.data==='string'?Buffer.from(file.data):Buffer.from(file.data);
    assert.equal(zip.readUInt32LE(central),0x02014b50);
    assert.equal(zip.readUInt16LE(central+8),0x0800);
    const nameLength=zip.readUInt16LE(central+28);
    assert.equal(zip.subarray(central+46,central+46+nameLength).toString(),expectedName);
    const local=zip.readUInt32LE(central+42);
    assert.equal(zip.readUInt32LE(local),0x04034b50);
    assert.equal(zip.readUInt16LE(local+6),0x0800);
    assert.equal(zip.readUInt32LE(local+14),crc32(expectedData));
    assert.equal(zip.readUInt32LE(central+16),crc32(expectedData));
    assert.equal(zip.readUInt32LE(local+18),expectedData.length);
    assert.equal(zip.readUInt16LE(local+26),nameLength);
    assert.equal(zip.subarray(local+30,local+30+nameLength).toString(),expectedName);
    assert.deepEqual(zip.subarray(local+30+nameLength,local+30+nameLength+expectedData.length),expectedData);
    central+=46+nameLength;
  }
  assert.equal(central,end);
  assert.equal(zip.readUInt32LE(end+12),end-centralStart);
  assert.equal(Buffer.from(run('createZipArchive([])')).length,22);
  assert.throws(()=>run('createZipArchive([{path:"bad",data:{}}])'),/데이터 형식/);
});

test('O02 ZIP allocates the final archive once instead of copying every PDF into intermediate records',()=>{
  const instance=app();
  instance.run('globalThis.files=Array.from({length:32},(_,i)=>({path:`큰 자료/반/${i}.pdf`,data:new Uint8Array(1024*1024),modifiedAt:new Date(2030,0,2)}))');
  instance.reset();
  const start=performance.now();
  const zip=instance.run('createZipArchive(files)');
  const elapsed=performance.now()-start;
  if(process.env.WORDLAB_BENCHMARK)console.log(JSON.stringify({zipBytes:zip.byteLength,allocatedBytes:instance.allocated(),elapsedMs:elapsed}));
  assert.ok(instance.allocated()<=zip.byteLength+1024,`temporary/final allocations ${instance.allocated()}, final ${zip.byteLength}`);
});

test('O02 classic ZIP rejects entry counts and UTF-8 names that would truncate header fields',()=>{
  const {run}=app();
  assert.throws(()=>run('createZipArchive(new Array(65536))'),/65,535/);
  assert.throws(()=>run('createZipArchive([{path:"한".repeat(21846),data:""}])'),/경로가 너무/);
  assert.equal(new DataView(run('createZipArchive([{path:"한".repeat(21845),data:""}])')).getUint16(26,true),65535);
  run('toUint8Array=()=>({length:0xffffffff});computeCrc32=()=>0;');
  assert.throws(()=>run('createZipArchive([{path:"huge.pdf",data:null}])'),/4 GiB/);
});

test('O02 synchronous saving retains the final keystroke in both groups without a pending timer',()=>{
  const {context,run}=app();
  run(`
    globalThis.saved='';elements.storageWarning={hidden:true};elements.classList={};
    getSettings=()=>({questionCount:100,sessionCount:5});
    setStoredStateRaw=text=>{saved=text;return true};
    globalThis.cards=Array.from({length:20},(_,i)=>({dataset:{classId:String(i)},querySelector:selector=>({value:selector==='.class-name-input'?'반 '+i:Array.from({length:500},(_,j)=>'word'+j+'\\t단어 뜻 '+j).join('\\n')})}));
    collectClassCards=()=>cards;getClassCardQuizMode=()=> 'word-to-meaning';
    state.presets['tue-thu'].classes=[{className:'다른 그룹',rawText:'keep\\t보존'}];
  `);
  const start=performance.now();
  for(let i=0;i<30;i++)assert.equal(run('persistState()'),true);
  const elapsed=performance.now()-start;
  const saved=JSON.parse(run('saved'));
  assert.equal(saved.presets['mon-wed'].classes.length,20);
  assert.match(saved.presets['mon-wed'].classes[19].rawText,/word499\t단어 뜻 499$/);
  assert.equal(saved.presets['tue-thu'].classes[0].rawText,'keep\t보존');
  assert.equal(run('state.hasUnsavedChanges'),false);
  if(process.env.WORDLAB_BENCHMARK)console.log(JSON.stringify({classes:20,wordsPerClass:500,jsonUtf8Bytes:Buffer.byteLength(run('saved')),meanCaptureSerializeMs:elapsed/30}));
});
