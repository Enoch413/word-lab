const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const os=require('node:os');
const net=require('node:net');
const http=require('node:http');
const {spawn,spawnSync}=require('node:child_process');

test('F08/F09/F12 isolated Windows server: availability, allowed assets and checkout identity', {skip:process.platform!=='win32',timeout:15000},async t=>{
  const root=fs.mkdtempSync(path.join(os.tmpdir(),'word lab 서버 '));
  fs.mkdirSync(path.join(root,'app/vendor'),{recursive:true});
  fs.mkdirSync(path.join(root,'.git'));
  for(const file of ['index.html','app/index.html','app/app.js','app/styles.css','app/firebase-config.js','app/vendor/html2canvas.min.js','app/vendor/jspdf.umd.min.js']){
    fs.copyFileSync(path.join(__dirname,'..',file),path.join(root,file));
  }
  fs.writeFileSync(path.join(root,'.git/HEAD'),'synthetic fixture');
  fs.writeFileSync(path.join(root,'private.json'),'synthetic fixture');
  const port=await new Promise(resolve=>{const probe=net.createServer();probe.listen(0,'127.0.0.1',()=>{const port=probe.address().port;probe.close(()=>resolve(port))})});
  const child=spawn('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',process.env.WORDLAB_TEST_SERVER_SOURCE || path.join(__dirname,'../wordlab-server.ps1'),'-Port',String(port),'-Root',root],{windowsHide:true,stdio:'ignore'});
  function request(url,method='GET'){
    return new Promise((resolve,reject)=>{
      const req=http.request({hostname:'127.0.0.1',port,path:url,method,agent:false},res=>{let length=0;res.on('data',chunk=>length+=chunk.length);res.on('end',()=>resolve({status:res.statusCode,headers:res.headers,length}))});
      req.on('error',reject);req.setTimeout(2000,()=>req.destroy(new Error('request timed out')));req.end();
    });
  }
  try{
    let ready=false;
    for(let i=0;i<60;i++){try{if((await request('/app/index.html')).status===200){ready=true;break}}catch{}await new Promise(resolve=>setTimeout(resolve,30));}
    assert.ok(ready,'server must start in a spaced Korean path');
    await t.test('only public assets are served',async()=>{
      for(const file of ['/app/index.html','/app/app.js','/app/styles.css','/app/vendor/jspdf.umd.min.js']) assert.equal((await request(file)).status,200);
      for(const file of ['/.git/HEAD','/private.json','/app/README.md','/app/%2e%2e/.git/HEAD','/app/..%5c.git%5cHEAD']) assert.ok([403,404].includes((await request(file)).status),file);
    });
    await t.test('HEAD, method restrictions and root identity',async()=>{
      const head=await request('/app/index.html','HEAD');
      assert.equal(head.length,0);
      assert.ok(Number(head.headers['content-length'])>0);
      assert.equal(head.headers['x-word-lab-root'],Buffer.from(path.resolve(root)).toString('base64'));
      assert.equal((await request('/app/index.html','POST')).status,405);
    });
    await t.test('START health check accepts only its own checkout and retains the existing path fix',()=>{
      const batch=fs.readFileSync(path.join(__dirname,'../START.bat'),'utf8');
      assert.ok(batch.includes('-Root "%~dp0."'));
      const command=batch.split(/\r?\n/).find(line=>line.startsWith('powershell ') && line.includes('exit 2')).match(/-Command "(.*)" >nul/)[1].replace('%APP_URL%',`http://127.0.0.1:${port}/app/index.html`);
      for(const [folder,expected] of [[root,0],[root+' other',2]]){
        const child=spawnSync('powershell.exe',['-NoProfile','-Command',command],{env:{...process.env,WORDLAB_ROOT:folder},windowsHide:true,stdio:'ignore',timeout:4000});
        assert.equal(child.status,expected);
      }
    });
    await t.test('an idle connection cannot indefinitely block a normal request',async()=>{
      const idle=net.createConnection({host:'127.0.0.1',port});
      idle.on('error',()=>{});
      await new Promise(resolve=>idle.once('connect',resolve));
      try{assert.equal((await request('/app/index.html')).status,200)}finally{idle.destroy()}
    });
    await t.test('oversized request headers are rejected',async()=>{
      const response=await new Promise((resolve,reject)=>{
        const socket=net.createConnection({host:'127.0.0.1',port});let data='';
        socket.setTimeout(2000,()=>socket.destroy(new Error('header timeout')));
        socket.on('error',error=>data ? resolve(data) : reject(error));socket.on('data',chunk=>data+=chunk);socket.on('end',()=>resolve(data));
        socket.on('connect',()=>socket.write('GET /app/index.html HTTP/1.1\r\nX-Test: '+ 'x'.repeat(33000)+'\r\n\r\n'));
      });
      assert.match(response,/^HTTP\/1\.1 431/);
    });
  }finally{
    child.kill();
    await new Promise(resolve=>child.once('exit',resolve));
    assert.equal(path.dirname(path.resolve(root)),path.resolve(os.tmpdir()));
    assert.ok(path.basename(root).startsWith('word lab 서버 '));
    fs.rmSync(root,{recursive:true,force:true});
  }
});
