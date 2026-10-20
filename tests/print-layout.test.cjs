const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(process.env.WORDLAB_TEST_SOURCE || path.join(__dirname,'../app/app.js'),'utf8');

function app(count=50,long=false){
  const context=vm.createContext({document:{addEventListener(){}},window:{},console});
  vm.runInContext(source,context);
  vm.runInContext(`
    elements.renderStage={innerHTML:'',appendChild(){}};
    globalThis.items=Array.from({length:${count}},(_,i)=>({word:'word'+i,meaning:'meaning'+i,answerHeight:${long?220:64}+(i%4)*4}));
    buildPrintablePage=options=>{
      const page={_printOptions:options,compact:false,remove(){}};
      page.classList={add(name){if(name==='sheet-page-compact')page.compact=true}};
      const naturalHeight=item=>options.answerKey?item.answerHeight:58;
      const rows=options.items.map(item=>({style:{},getBoundingClientRect(){return {height:Math.max(parseFloat(this.style.minHeight)||0,naturalHeight(item)-(page.compact?12:0))}}}));
      const header={style:{},getBoundingClientRect(){return {height:74}}};
      const info={style:{},getBoundingClientRect(){return {height:options.answerKey?64:60}}};
      page.querySelectorAll=selector=>selector==='.sheet-row'?rows:[header,info];
      page.getBoundingClientRect=()=>({height:Math.max(1123,(page.compact?170:240)+Math.max(...splitIntoColumns(rows,4).map(column=>column.reduce((height,row)=>height+row.getBoundingClientRect().height+(page.compact?3:4),0))))});
      return page;
    };
    globalThis.options={className:'검사반',sessionNumber:1,items,quizMode:'word-to-meaning',dateLabel:'2030.01.02',itemStartIndex:0,totalItemCount:items.length};
  `,context);
  return {run:code=>vm.runInContext(code,context)};
}

test('50 questions and answers share one A4 page, column boundaries and row heights',()=>{
  const {run}=app();
  const quiz=run('paginatePrintablePages([buildPrintablePage({...options,answerKey:false})])');
  const answers=run('paginatePrintablePages([buildPrintablePage({...options,answerKey:true})])');
  assert.equal(quiz.length,1);
  assert.equal(answers.length,1);
  assert.equal(quiz[0].getBoundingClientRect().height,1123);
  assert.equal(answers[0].getBoundingClientRect().height,1123);
  assert.deepEqual(Array.from(quiz[0].querySelectorAll('.sheet-row'),r=>r.getBoundingClientRect().height),Array.from(answers[0].querySelectorAll('.sheet-row'),r=>r.getBoundingClientRect().height));
  assert.equal(quiz[0].compact,answers[0].compact);
});

test('100 long questions use identical page partitions and preserve every numbered entry',()=>{
  const {run}=app(100,true);
  const quiz=run('paginatePrintablePages([buildPrintablePage({...options,answerKey:false})])');
  const answers=run('paginatePrintablePages([buildPrintablePage({...options,answerKey:true})])');
  const partitions=pages=>Array.from(pages,p=>Array.from(p._printOptions.items,item=>item.word));
  assert.deepEqual(partitions(quiz),partitions(answers));
  assert.deepEqual(partitions(quiz).flat(),Array.from({length:100},(_,i)=>'word'+i));
  for(const pages of [quiz,answers]){
    let offset=0;
    pages.forEach((page,index)=>{
      assert.ok(page.getBoundingClientRect().height<=1123);
      assert.equal(page._printOptions.itemStartIndex,offset);
      assert.equal(page._printOptions.pageNumber,index+1);
      assert.equal(page._printOptions.pageCount,pages.length);
      offset+=page._printOptions.items.length;
    });
  }
});

test('50 oversized entries fail clearly instead of clipping, shrinking or producing mismatched pages',()=>{
  const {run}=app(50,true);
  assert.throws(()=>run('paginatePrintablePages([buildPrintablePage({...options,answerKey:false})])'),/한 페이지|한 쪽/);
});
