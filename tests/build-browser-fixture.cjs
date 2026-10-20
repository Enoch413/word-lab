// Standalone synthetic PDF checks; no authentication or user-folder access.
const fs=require('node:fs');
const path=require('node:path');
const target=process.argv[2];
if(!target || !path.isAbsolute(target)) throw new Error('Pass an absolute fixture output directory');
const appRoot=path.join(__dirname,'../app');
const appTarget=path.join(target,'app');
fs.mkdirSync(path.join(appTarget,'vendor'),{recursive:true});
for(const file of ['styles.css','vendor/html2canvas.min.js','vendor/jspdf.umd.min.js']) fs.copyFileSync(path.join(appRoot,file),path.join(appTarget,file));
const source=fs.readFileSync(path.join(appRoot,'app.js'),'utf8');
const fixtureVersion=require('node:crypto').createHash('sha256').update(source).update(fs.readFileSync(path.join(appTarget,'styles.css'))).digest('hex').slice(0,12);
const wanted=['buildSheetPages','buildAnswerSheetPages','buildPrintablePages','buildPrintablePage','buildMatchingPrintablePage','paginatePrintablePages','renderSheetColumn','getPromptAndAnswer','splitIntoColumns','chunkItems','escapeHtml','renderPagesToPdfBuffer','waitForLayout','hasNativeHost','withOutputLock','createZipArchive','normalizeZipPath','toUint8Array','computeCrc32','getCrc32Table','getDosDateTime'];
const functions=wanted.map(name=>{
  const match=source.match(new RegExp('(?:async )?function '+name+'\\([^]*?(?=\\n(?:async )?function |$)'));
  if(!match) throw new Error(name);
  return match[0];
}).join('\n');
fs.writeFileSync(path.join(appTarget,'app.js'),`const PRINTABLE_ITEMS_PER_PAGE=50;
const elements={renderStage:document.getElementById('render-stage')};
${functions}
const mediaRules=Array.from(document.styleSheets).filter(sheet=>sheet.href?.split('?')[0].endsWith('styles.css')).flatMap(sheet=>Array.from(sheet.cssRules)).filter(rule=>rule.media);
document.getElementById('motion-mode').onchange=event=>{
  const simulated=event.target.value;
  mediaRules.forEach(rule=>{
    rule.originalMedia ??= rule.media.mediaText;
    if(rule.originalMedia.includes('prefers-reduced-motion'))rule.media.mediaText=simulated==='system'?rule.originalMedia:simulated==='reduce'?'all':'not all';
    if(rule.originalMedia.includes('(hover: hover)'))rule.media.mediaText=simulated==='touch'?'not all':rule.originalMedia;
  });
  document.getElementById('motion-result').textContent=JSON.stringify({mode:simulated,scroll:getComputedStyle(document.documentElement).scrollBehavior,toastTransform:getComputedStyle(document.getElementById('sample-toast')).transform,toastTransition:getComputedStyle(document.getElementById('sample-toast')).transitionDuration,buttonTransition:getComputedStyle(document.getElementById('sample-button')).transitionDuration,disabledTransform:getComputedStyle(document.getElementById('sample-disabled')).transform});
};
let releaseTestLock;
document.getElementById('lock').onclick=async()=>{
  document.getElementById('lock-status').textContent='잠금 대기';
  try{
    await withOutputLock(async()=>{
      document.getElementById('lock-status').textContent='잠금 획득';
      await new Promise(resolve=>{releaseTestLock=resolve});
    });
    document.getElementById('lock-status').textContent='잠금 해제';
  }catch(error){document.getElementById('lock-status').textContent=String(error)}
};
document.getElementById('release').onclick=()=>releaseTestLock?.();
document.getElementById('run').onclick=async()=>{
  const result=document.getElementById('result');
  result.textContent='검사 중';document.getElementById('run').disabled=true;
  const findings=[];
  const zipEntries=[];
  try{
    const cases=[
      {name:'normal-quiz',count:50,mode:'word-to-meaning',answer:false,long:false,expectedPages:1},
      {name:'normal-answer',count:50,mode:'meaning-to-word',answer:true,long:false,expectedPages:1},
      {name:'mixed-answer-50',count:50,mode:'word-to-meaning',answer:true,mixed:true,expectedPages:1},
      {name:'normal-quiz-100',count:100,mode:'word-to-meaning',answer:false,long:false,expectedPages:2},
      {name:'normal-answer-100',count:100,mode:'word-to-meaning',answer:true,long:false,expectedPages:2},
      {name:'long-answer-100',count:100,mode:'word-to-meaning',answer:true,long:true},
      {name:'long-quiz-100',count:100,mode:'meaning-to-word',answer:false,long:true}
    ];
    for(const item of cases){
      result.textContent='검사 중: '+item.name;
      const items=Array.from({length:item.count},(_,index)=>({word:'word'+index,meaning:item.long?'아주 긴 뜻과 설명 '.repeat(18):item.mixed?'뜻과 설명 '.repeat(index%7===0?8:2):'뜻'+index}));
      const options={className:'검사반',sessionNumber:1,items,quizMode:item.mode,dateLabel:'2030.01.02'};
      const layouts=[];
      for(const answerKey of [false,true]){
        const candidate=answerKey?buildAnswerSheetPages(options):buildSheetPages(options);
        elements.renderStage.innerHTML='';candidate.forEach(page=>elements.renderStage.appendChild(page));
        await waitForLayout();
        const matched=paginatePrintablePages(candidate);
        elements.renderStage.innerHTML='';matched.forEach(page=>elements.renderStage.appendChild(page));
        layouts.push(matched.map(page=>{
          const origin=page.getBoundingClientRect();
          return Array.from(page.querySelectorAll('.sheet-row'),row=>{
            const rect=row.getBoundingClientRect();
            return {number:row.querySelector('.sheet-num').textContent,prompt:row.querySelector('.sheet-prompt').textContent,x:rect.x-origin.x,y:rect.y-origin.y,height:rect.height};
          });
        }));
      }
      if(JSON.stringify(layouts[0])!==JSON.stringify(layouts[1])) throw new Error(item.name+': quiz/key positions differ');
      const pages=item.answer?buildAnswerSheetPages(options):buildSheetPages(options);
      const bytes=await renderPagesToPdfBuffer(pages,item.name);
      zipEntries.push({path:'합성 자료/검사반/'+item.name+'.pdf',data:bytes});
      const printed=Array.from(elements.renderStage.querySelectorAll('.sheet-page'));
      const numbers=printed.flatMap(page=>Array.from(page.querySelectorAll('.sheet-num'),node=>Number(node.textContent.replace('.',''))));
      const heights=printed.map(page=>page.getBoundingClientRect().height);
      if(item.count===50){
        const page=printed[0];
        const bottoms=Array.from(page.querySelectorAll('.sheet-column'),column=>column.lastElementChild.getBoundingClientRect().bottom);
        const footer=page.querySelector('.sheet-footer');
        const gap=footer.getBoundingClientRect().top-Math.max(...bottoms);
        if(Math.max(...bottoms)-Math.min(...bottoms)>1 || gap>parseFloat(getComputedStyle(footer).marginTop)+1)throw new Error(item.name+': columns do not fill to the footer');
      }
      if(numbers.length!==item.count || numbers.some((value,index)=>value!==index+1)) throw new Error(item.name+': missing or reordered entries');
      if(heights.some(height=>height>1123)) throw new Error(item.name+': page overflow');
      findings.push({name:item.name,pages:printed.length,entries:numbers.length,heights,bytes:bytes.byteLength});
      if(item.expectedPages && printed.length!==item.expectedPages) throw new Error(item.name+': expected '+item.expectedPages+' pages, got '+printed.length);
      const link=document.createElement('a');link.download=item.name+'.pdf';link.textContent=item.name+' PDF';link.href=URL.createObjectURL(new Blob([bytes],{type:'application/pdf'}));link.style.display='block';document.getElementById('files').appendChild(link);
      if(item.name==='long-answer-100') document.getElementById('preview').appendChild(printed[0].cloneNode(true));
    }
    const oversized=Array.from({length:50},(_,i)=>({word:'word'+i,meaning:'아주 긴 뜻과 설명 '.repeat(18)}));
    try{
      await renderPagesToPdfBuffer(buildSheetPages({className:'긴 설명 검사반',sessionNumber:1,items:oversized,quizMode:'word-to-meaning',dateLabel:'2030.01.02'}),'oversized');
      throw new Error('oversized 50-item sheet was silently accepted');
    }catch(error){if(!/한 페이지/.test(String(error)))throw error;}
    zipEntries.push({path:'합성 자료/복원.json',data:JSON.stringify({appName:'WORD LAB',note:'UTF-8 단어 🍎',findings})});
    const zip=createZipArchive(zipEntries);
    const zipLink=document.createElement('a');zipLink.download='wordlab-p3-regression.zip';zipLink.textContent='합성 자료 ZIP';zipLink.href=URL.createObjectURL(new Blob([zip],{type:'application/zip'}));document.getElementById('files').appendChild(zipLink);
    result.textContent=JSON.stringify({pass:true,findings,zipBytes:zip.byteLength},null,2);
  }catch(error){result.textContent=JSON.stringify({pass:false,error:String(error),findings},null,2)}
  finally{document.getElementById('run').disabled=false}
};`);
fs.writeFileSync(path.join(appTarget,'index.html'),`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>WORD LAB PDF 회귀 검사</title><link rel="stylesheet" href="styles.css?v=${fixtureVersion}"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;800&display=swap" rel="stylesheet"></head><body><h1>합성 자료 PDF 회귀 검사</h1><p>실제 출력 함수·CSS·PDF 라이브러리 사용. 인증·사용자 폴더 접근 없음.</p><label>모션 조건 검사 <select id="motion-mode"><option value="system">시스템 설정</option><option value="normal">일반 모션 CSS</option><option value="reduce">모션 감소 CSS</option><option value="touch">터치 CSS</option></select></label><p>아래 선택은 이 검사 화면의 미디어 규칙만 활성/비활성화합니다. OS 설정은 바꾸지 않습니다.</p><button class="btn btn-green" id="sample-button">활성 버튼</button><button class="btn btn-green" id="sample-disabled" disabled>비활성 버튼</button><span class="toast" id="sample-toast">토스트 표본</span><pre id="motion-result"></pre><button id="lock">출력 잠금 획득</button><button id="release">출력 잠금 해제</button><p id="lock-status">잠금 검사 전</p><button id="run">PDF 검사 실행</button><pre id="result"></pre><div id="files"></div><div id="preview"></div><div class="render-stage" id="render-stage" aria-hidden="true"></div><script src="vendor/html2canvas.min.js"></script><script src="vendor/jspdf.umd.min.js"></script><script src="app.js?v=${fixtureVersion}"></script></body></html>`);
console.log(target);
