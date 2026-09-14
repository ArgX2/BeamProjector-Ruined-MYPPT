import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import * as pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs';
import {PptxViewer, parseZip, buildPresentation, RECOMMENDED_ZIP_LIMITS} from '@aiden0z/pptx-renderer/browser';

// Bundled worker handler also works under file://, without fetching a worker script.
globalThis.pdfjsWorker = pdfWorker;
const pdfAssets = __PDF_ASSETS__;
function asset(name) { const raw = atob(pdfAssets[name] || ''); if (!raw) throw Error('Missing PDF resource'); return Uint8Array.from(raw, c=>c.charCodeAt(0)); }
class CMaps { async fetch({name}) { return {cMapData:asset(name+'.bcmap'),compressionType:1}; } }
class Fonts { async fetch({filename}) { return asset(filename); } }
const $ = id => document.getElementById(id);
const definitions = [
  ['brightness','밝기',-70,70,1,'%'],['contrast','대비 저하',0,80,1,'%'],
  ['saturation','채도 저하',0,100,1,'%'],['ambient','주변광 · 흰색 번짐',0,60,1,'%'],
  ['warmth','누런 색감',0,80,1,'%'],['blur','초점 흐림',0,5,.1,'px'],['vignette','주변부 어두움',0,80,1,'%'],['bloom','블루밍 · 빛 번짐',0,100,1,'%']
];
const presets = {
 classroom:[-8,18,22,12,8,.4,8,12], bright:[8,28,30,27,7,.5,7,28],
 old:[-24,32,38,15,25,1.1,25,22], dark:[-4,8,10,3,2,.2,5,8]
};
let values={}, mode='projected', zoom=100, fitMode=true, width=1280,height=720;
let doc=null, current=0, total=1, busy=false, pageHandle=null, imageURL=null, viewToken=0;
let note='파일을 끌어 놓거나 ‘파일 열기’를 선택하세요.';
let resolution=null, resolutionToken=0, rasterSource=null, rasterReady=false;
function syncResolution(){
 const use=mode==='projected'&&resolution&&rasterReady;
 $('surface').style.visibility=use?'hidden':'visible';
 $('rasterSurface').hidden=!use;
}
async function snapshotSource(){
 const direct=$('surface').firstElementChild;
 if(direct instanceof HTMLCanvasElement||direct instanceof HTMLImageElement)return direct;
 // Flatten PPTX HTML/SVG at its intrinsic size before downsampling.
 await document.fonts.ready;
 const root=document.createElement('div');root.setAttribute('xmlns','http://www.w3.org/1999/xhtml');root.style.cssText=`position:relative;width:${width}px;height:${height}px;background:white;overflow:hidden;`;
 async function copy(node){
  if(node.nodeType!==1)return node.cloneNode(true);
  let clone=node.cloneNode(false);
  if(node instanceof HTMLCanvasElement){clone=document.createElement('img');clone.src=node.toDataURL();}
  if(node instanceof HTMLImageElement){await node.decode();const c=canvas(node.naturalWidth,node.naturalHeight);c.getContext('2d').drawImage(node,0,0);clone.src=c.toDataURL();clone.removeAttribute('srcset');}
  const style=getComputedStyle(node);for(const prop of style)clone.style.setProperty(prop,style.getPropertyValue(prop));
  for(const child of node.childNodes)clone.append(await copy(child));
  return clone;
 }
 for(const child of $('surface').children)root.append(await copy(child));
 const markup=new XMLSerializer().serializeToString(root);
 const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="100%" height="100%">${markup}</foreignObject></svg>`;
 const img=new Image();img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);await img.decode();return img;
}
async function updateResolution(){
 const token=++resolutionToken;rasterReady=false;syncResolution();
 if(!resolution){$('resolutionInfo').textContent='원본 선명도';refreshBloom();return;}
 $('resolutionInfo').textContent='해상도 적용 중…';
 try{
  const pending=rasterSource||(rasterSource=snapshotSource());const src=await pending;
  if(token!==resolutionToken)return;
  const scale=Math.min(resolution[0]/width,resolution[1]/height);
  const w=Math.max(1,Math.round(width*scale)),h=Math.max(1,Math.round(height*scale));
  const c=canvas(w,h),ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';ctx.drawImage(src,0,0,w,h);
  c.style.width=width+'px';c.style.height=height+'px';c.style.imageRendering='pixelated';$('rasterSurface').replaceChildren(c);rasterReady=true;
  $('resolutionInfo').textContent=`영상 영역 ${w} × ${h} px · 비율 유지`;
  layout();refreshBloom();effect();
 }catch(e){if(token!==resolutionToken)return;rasterSource=null;$('resolutionInfo').textContent='해상도 변환 실패 · 원본 선명도로 표시';refreshBloom();}
}
function contentChanged(){rasterSource=null;updateResolution();}
function chooseResolution(custom=false){
 if($('resolution').value==='native'){resolution=null;}else if(custom||$('resolution').value==='custom'){
  const w=Number($('resolutionWidth').value),h=Number($('resolutionHeight').value);
  if(!Number.isFinite(w)||!Number.isFinite(h)||w<160||w>3840||h<90||h>2160){$('resolutionWidth').value=resolution?.[0]||1024;$('resolutionHeight').value=resolution?.[1]||768;return;}
  resolution=[Math.round(w),Math.round(h)];$('resolution').value='custom';
 }else resolution=$('resolution').value.split('x').map(Number);
 $('resolutionDimensions').hidden=!resolution;
 if(resolution){$('resolutionWidth').value=resolution[0];$('resolutionHeight').value=resolution[1];}
 updateResolution();
}
function message(text,error=false){$('message').textContent=text;$('message').classList.toggle('error',error);}
function busyState(on){busy=on;document.body.classList.toggle('busy',on);$('file').disabled=on;$('open').disabled=on;$('sample').disabled=on;navigation();}
function navigation(){$('prev').disabled=busy||current===0;$('next').disabled=busy||current>=total-1;$('page').disabled=busy;$('page').value=current+1;$('page').max=total;$('pageCount').textContent=`/ ${total}`;}
function effect(){
 const p=mode==='projected';
 const base=`brightness(${1+values.brightness/100}) contrast(${1-values.contrast/100}) saturate(${1-values.saturation/100}) sepia(${values.warmth/100})`;
 $('surface').style.filter=p?`${base} blur(${values.blur}px)`:'none';
 $('rasterSurface').style.filter=$('surface').style.filter;syncResolution();
 // A bright-pass copy spreads only highlights; screen compositing leaves dark pixels alone.
 $('bloom').style.filter=`${base} brightness(0.7) contrast(5) blur(${2+values.bloom*.18}px)`;
 $('bloom').style.opacity=p?values.bloom/100:0;
 $('ambient').style.opacity=p?values.ambient/100:0;$('vignette').style.opacity=p?values.vignette/100:0;
 $('original').setAttribute('aria-pressed',!p);$('projected').setAttribute('aria-pressed',p);$('modeLabel').textContent=p?'프로젝터 미리보기':'원본 미리보기';
}
function setMode(next){mode=next;effect();}
function refreshBloom(){
 const layer=$('bloom');layer.replaceChildren();
 const source=resolution&&rasterReady?$('rasterSurface'):$('surface');
 for(const child of source.children)layer.append(child.cloneNode(true));
 const sources=source.querySelectorAll('canvas'),targets=layer.querySelectorAll('canvas');
 sources.forEach((c,i)=>targets[i]?.getContext('2d').drawImage(c,0,0));
 layer.querySelectorAll('[id]').forEach(el=>{if(!el.closest('svg'))el.removeAttribute('id');});
}
function setValue(key,value,custom=true){
 const d=definitions.find(x=>x[0]===key);if(!d||!Number.isFinite(value))return;
 value=Math.round(Math.max(d[2],Math.min(d[3],value))/d[4])*d[4];value=Number(value.toFixed(1));
 values[key]=value;$(key+'Range').value=value;$(key+'Number').value=value;
 if(custom)$('preset').value='custom';effect();
}
for(const [key,label,min,max,step,unit] of definitions){
 const row=document.createElement('div');row.className='control';
 row.innerHTML=`<div class="control-head"><label for="${key}Range">${label}</label><div class="value-box"><input id="${key}Number" aria-label="${label} 값" type="number" min="${min}" max="${max}" step="${step}"><span>${unit}</span></div></div><input id="${key}Range" type="range" min="${min}" max="${max}" step="${step}">`;
 $('controls').append(row);
 $(key+'Range').addEventListener('input',e=>setValue(key,+e.target.value));
 $(key+'Number').addEventListener('input',e=>{if(e.target.value!==''&&e.target.validity.valid)setValue(key,+e.target.value);});
 $(key+'Number').addEventListener('change',e=>{if(e.target.value==='')e.target.value=values[key];else setValue(key,+e.target.value);});
}
function preset(name){const p=presets[name];if(!p)return;definitions.forEach((d,i)=>setValue(d[0],p[i],false));$('preset').value=name;}
function layout(){
 if(fitMode){const pad=window.innerWidth<=720?40:80;zoom=Math.min(($('viewport').clientWidth-pad)/width,($('viewport').clientHeight-pad)/height)*100;zoom=Math.max(1,Math.min(400,zoom));}
 $('stage').style.width=width*zoom/100+'px';$('stage').style.height=height*zoom/100+'px';
 $('surface').style.width=width+'px';$('surface').style.height=height+'px';$('surface').style.transform=`scale(${zoom/100})`;
 $('rasterSurface').style.width=width+'px';$('rasterSurface').style.height=height+'px';$('rasterSurface').style.transform=`scale(${zoom/100})`;
 $('bloom').style.width=width+'px';$('bloom').style.height=height+'px';$('bloom').style.transform=`scale(${zoom/100})`;
 $('zoomValue').textContent=Math.round(zoom)+'%';$('fit').setAttribute('aria-pressed',fitMode);
 $('minus').disabled=!fitMode&&zoom<=5;$('plus').disabled=zoom>=400;
}
function zoomStep(delta){fitMode=false;zoom=Math.max(5,Math.min(400,(delta>0?Math.floor(zoom/5)+1:Math.ceil(zoom/5)-1)*5));layout();}
function clearDocument(){viewToken++;resolutionToken++;rasterSource=null;rasterReady=false;pageHandle?.dispose();pageHandle=null;doc?.viewer?.destroy();doc?.pdfTask?.destroy();doc=null;if(imageURL){URL.revokeObjectURL(imageURL);imageURL=null;}$('surface').replaceChildren();$('bloom').replaceChildren();$('rasterSurface').replaceChildren();syncResolution();}
function canvas(w,h){const c=document.createElement('canvas');c.width=w;c.height=h;return c;}
function sample(){
 if(busy)return;clearDocument();width=1280;height=720;current=0;total=1;
 const c=canvas(width,height),g=c.getContext('2d');g.fillStyle='#ffffff';g.fillRect(0,0,width,height);
 g.fillStyle='#2462e9';g.fillRect(72,68,42,5);g.font='18px "Segoe UI", "Malgun Gothic", sans-serif';g.fillStyle='#667389';g.fillText('PROJECTION TEST',72,108);
 g.fillStyle='#18283e';g.font='bold 52px "Segoe UI", "Malgun Gothic", sans-serif';g.fillText('발표 전에, 한 번 더.',72,194);
 g.font='22px "Segoe UI", "Malgun Gothic", sans-serif';g.fillStyle='#758296';g.fillText('색감과 대비가 강의실에서도 잘 보이는지 확인하세요.',72,244);
 const colors=['#2158db','#0092be','#11a78c','#edbd27','#e77548','#b75797'];
 colors.forEach((col,i)=>{g.fillStyle=col;g.fillRect(72+i*190,301,174,148);g.fillStyle='#5a677b';g.font='16px monospace';g.fillText(col.toUpperCase(),72+i*190,478);});
 [0,1,2,3,4,5,6,7,8,9].forEach(i=>{g.fillStyle=`rgb(${i*28.333},${i*28.333},${i*28.333})`;g.fillRect(72+i*114,525,114,48);});
 g.fillStyle='#17263e';g.font='24px "Malgun Gothic", sans-serif';g.fillText('큰 제목과 진한 글씨',72,625);
 g.fillStyle='#a9b1bc';g.font='17px "Malgun Gothic", sans-serif';g.fillText('옅은 글씨와 가는 선은 어떻게 보이나요?',690,625);
 g.strokeStyle='#dce2e9';g.beginPath();g.moveTo(690,642);g.lineTo(1200,642);g.stroke();
 $('surface').append(c);contentChanged();$('filename').textContent='색감 테스트 · 샘플';note='파일을 끌어 놓거나 ‘파일 열기’를 선택하세요.';message(note);fitMode=true;layout();navigation();effect();
}
async function renderPage(index){
 if(!doc)return;const token=++viewToken;pageHandle?.dispose();pageHandle=null;$('surface').replaceChildren();
 if(doc.pdf){
  const p=await doc.pdf.getPage(index+1);if(token!==viewToken)return;
  const base=p.getViewport({scale:1});const factor=Math.min(2,2400/Math.max(base.width,base.height));const vp=p.getViewport({scale:factor});
  width=base.width;height=base.height;const c=canvas(Math.ceil(vp.width),Math.ceil(vp.height));c.style.width=width+'px';c.style.height=height+'px';
  await p.render({canvasContext:c.getContext('2d'),viewport:vp,background:'#ffffff'}).promise;if(token!==viewToken)return;$('surface').append(c);p.cleanup();
 }else if(doc.viewer){
  width=doc.viewer.slideWidth;height=doc.viewer.slideHeight;
  pageHandle=doc.viewer.renderSlideToContainer(index,$('surface'),1);await pageHandle?.ready;
  // Imported links and media are presentation content, never an interactive browser surface.
  $('surface').querySelectorAll('a,video,audio,iframe').forEach(el=>{el.removeAttribute('href');el.removeAttribute('src');el.setAttribute('tabindex','-1');});
 }
 current=index;contentChanged();navigation();layout();effect();
}
async function go(index){if(busy||!doc||!Number.isInteger(index)||index<0||index>=total){navigation();return;}busyState(true);try{await renderPage(index);message(note);}catch(e){message('페이지를 표시하지 못했습니다. 다른 페이지를 선택하거나 PDF로 다시 열어 주세요.',true);}finally{busyState(false);}}
async function loadFile(file){
 if(!file||busy)return;
 const ext=file.name.split('.').pop().toLowerCase();
 if(!['pdf','pptx','png','jpg','jpeg','webp','bmp'].includes(ext)){message('PDF, PPTX, PNG, JPG, WEBP, BMP 파일을 선택하세요.',true);return;}
 if(file.size>100*1024*1024){message('100MB 이하의 파일을 선택하세요.',true);return;}
 busyState(true);message('파일을 여는 중…');
 let candidate=null;
 try{
  const data=await file.arrayBuffer();
  if(ext==='pdf'){
   const task=pdfjs.getDocument({data,isEvalSupported:false,useWasm:false,useWorkerFetch:false,CMapReaderFactory:CMaps,StandardFontDataFactory:Fonts});
   try{candidate={pdf:await task.promise,pdfTask:task};}catch(e){await task.destroy();throw e;}
  }else if(ext==='pptx'){
   const files=await parseZip(data,RECOMMENDED_ZIP_LIMITS);const model=buildPresentation(files);
   const viewer=new PptxViewer(document.createElement('div'),{fitMode:'none',pdfjs:false});viewer.load(model);candidate={viewer};
   if(!viewer.slideCount)throw Error('Empty presentation');
  }else{
   const url=URL.createObjectURL(new Blob([data],{type:file.type||`image/${ext==='jpg'?'jpeg':ext}`}));const img=new Image();
   try{img.src=url;await img.decode();if(img.naturalWidth*img.naturalHeight>64000000)throw Error('Image too large');}catch(e){URL.revokeObjectURL(url);throw e;}
   candidate={img,url};
  }
  clearDocument();doc=candidate;candidate=null;current=0;fitMode=true;
  total=doc.pdf?.numPages||doc.viewer?.slideCount||1;$('filename').textContent=file.name;$('filename').title=file.name;
  note=ext==='pptx'?'PPTX의 일부 글꼴·도형·효과는 다를 수 있습니다. 정확한 배치는 PDF로 확인하세요.':'파일은 이 브라우저에서만 처리됩니다.';
  if(doc.img){imageURL=doc.url;width=doc.img.naturalWidth;height=doc.img.naturalHeight;$('surface').append(doc.img);contentChanged();layout();effect();}
  else await renderPage(0);
  message(note);
 }catch(e){console.error('File rendering failed:',e);candidate?.viewer?.destroy();if(candidate?.pdfTask)await candidate.pdfTask.destroy();
  message(e?.name==='PasswordException'?'암호로 보호된 PDF입니다. 암호를 해제한 파일을 열어 주세요.':'파일을 열지 못했습니다. 파일 손상 여부를 확인하거나 PDF로 변환해 다시 열어 주세요.',true);
 }finally{busyState(false);$('file').value='';}
}
$('open').onclick=()=>$('file').click();$('file').onchange=e=>loadFile(e.target.files[0]);
$('original').onclick=()=>setMode('original');$('projected').onclick=()=>setMode('projected');
$('preset').onchange=e=>preset(e.target.value);$('reset').onclick=()=>{definitions.forEach(d=>setValue(d[0],0));$('resolution').value='native';chooseResolution();};
$('resolution').onchange=()=>chooseResolution();
for(const id of ['resolutionWidth','resolutionHeight'])$(id).onchange=()=>chooseResolution(true);
$('sample').onclick=sample;$('minus').onclick=()=>zoomStep(-1);$('plus').onclick=()=>zoomStep(1);$('fit').onclick=()=>{fitMode=true;layout();};
$('prev').onclick=()=>go(current-1);$('next').onclick=()=>go(current+1);$('page').onchange=e=>go(Number(e.target.value)-1);
$('panelToggle').onclick=()=>{const hidden=!$('panel').hidden;$('panel').hidden=hidden;document.querySelector('main').classList.toggle('collapsed',hidden);$('panelToggle').setAttribute('aria-expanded',!hidden);$('panelToggle').title=hidden?'조절 패널 펼치기':'조절 패널 접기';layout();};
new ResizeObserver(layout).observe($('viewport'));
let dragDepth=0;
window.addEventListener('dragenter',e=>{if(e.dataTransfer?.types.includes('Files')){e.preventDefault();dragDepth++;document.body.classList.add('dragging');}});
window.addEventListener('dragover',e=>{e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect='copy';});
window.addEventListener('dragleave',()=>{if(--dragDepth<=0){dragDepth=0;document.body.classList.remove('dragging');}});
window.addEventListener('drop',e=>{e.preventDefault();dragDepth=0;document.body.classList.remove('dragging');loadFile(e.dataTransfer?.files[0]);});
window.addEventListener('keydown',e=>{
 if((e.ctrlKey||e.metaKey)&&e.key==='o'){e.preventDefault();if(!busy)$('file').click();return;}
 if(/INPUT|SELECT|TEXTAREA|BUTTON/.test(document.activeElement.tagName))return;
 if(e.key==='ArrowLeft'){e.preventDefault();go(current-1);}if(e.key==='ArrowRight'){e.preventDefault();go(current+1);}
 if(e.code==='Space'){e.preventDefault();setMode(mode==='projected'?'original':'projected');}
 if(e.key==='0'){fitMode=true;layout();}
});
preset('classroom');sample();
if(window.matchMedia('(max-width:720px)').matches)$('panelToggle').click();
// Optional WebMCP: the same preview state and validation as the visible controls.
try{const context=document.modelContext;if(context?.registerTool){const lifecycle=new AbortController();context.registerTool({name:'configure_projector_preview',description:'현재 미리보기의 원본/프로젝터 모드와 열화 값을 조절합니다.',inputSchema:{type:'object',properties:{mode:{type:'string',enum:['original','projected']},values:{type:'object',properties:Object.fromEntries(definitions.map(d=>[d[0],{type:'number',minimum:d[2],maximum:d[3]}])),additionalProperties:false}},additionalProperties:false},annotations:{readOnlyHint:false},execute(input){if(!input||typeof input!=='object')throw Error('Invalid input');if(input.mode&&!['original','projected'].includes(input.mode))throw Error('Invalid mode');for(const[k,v]of Object.entries(input.values||{})){const d=definitions.find(d=>d[0]===k);if(!d||typeof v!=='number'||!Number.isFinite(v)||v<d[2]||v>d[3])throw Error('Invalid value');}if(input.mode)setMode(input.mode);Object.entries(input.values||{}).forEach(([k,v])=>setValue(k,v));return {mode,values:{...values}};}},{signal:lifecycle.signal});window.addEventListener('pagehide',()=>lifecycle.abort(),{once:true});}}catch{}
