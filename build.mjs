import {build} from 'esbuild';
import {readFile,writeFile,readdir,mkdir,cp} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
const root=path.dirname(fileURLToPath(import.meta.url));process.chdir(root);
const assets={};
for(const dir of ['cmaps','standard_fonts'])for(const f of await readdir('node_modules/pdfjs-dist/'+dir))if(!f.toLowerCase().includes('license'))assets[f]=(await readFile('node_modules/pdfjs-dist/'+dir+'/'+f)).toString('base64');
const explicitFiles={name:'explicit-files',setup(b){
 b.onResolve({filter:/.*/},args=>{
  const from=args.importer||path.join(root,'package.json');
  const resolved=path.isAbsolute(args.path)?args.path:createRequire(from).resolve(args.path);
  return {path:resolved,namespace:'source'};
 });
 b.onLoad({filter:/.*/,namespace:'source'},async args=>({contents:await readFile(args.path,'utf8'),loader:'js'}));
}};
const result=await build({absWorkingDir:root,tsconfigRaw:{},plugins:[explicitFiles],entryPoints:[path.join(root,'src/app.js')],bundle:true,write:false,minify:true,format:'iife',target:'es2022',define:{__PDF_ASSETS__:JSON.stringify(assets),'process.env.NODE_ENV':'"production"'},legalComments:'inline'});
const js=result.outputFiles[0].text.replace(/<\/script/gi,'<\\/script');
const css=await readFile('src/style.css','utf8');
const html=(await readFile('src/index.html','utf8')).replace('/* STYLE */',()=>css).replace('/* SCRIPT */',()=>js);
await writeFile('index.html',html);await writeFile('.nojekyll','');
await mkdir('licenses',{recursive:true});
for(const name of ['pdfjs-dist','@aiden0z/pptx-renderer']){
 const folder='node_modules/'+name;
 for(const f of ['LICENSE','THIRD_PARTY_NOTICES.md'])try{await writeFile('licenses/'+name.replaceAll('/','-')+'-'+f,await readFile(folder+'/'+f));}catch{}
}
await cp('node_modules/@aiden0z/pptx-renderer/licenses','licenses/licenses',{recursive:true});
for(const dir of ['standard_fonts','cmaps'])for(const f of await readdir('node_modules/pdfjs-dist/'+dir))if(f.toLowerCase().includes('license'))await cp('node_modules/pdfjs-dist/'+dir+'/'+f,'licenses/pdfjs-'+dir+'-'+f);
for(const name of ['jszip','echarts','zrender','pako'])try{for(const f of await readdir('node_modules/'+name))if(/^(licen[cs]e|notice)/i.test(f))await cp('node_modules/'+name+'/'+f,'licenses/'+name+'-'+f);}catch{}
console.log(`Built standalone index.html (${(Buffer.byteLength(html)/1048576).toFixed(1)} MB)`);
