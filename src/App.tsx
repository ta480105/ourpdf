import { useEffect, useMemo, useRef, useState } from 'react'
import { PDFDocument, StandardFonts, rgb, degrees } from 'pdf-lib'
import * as pdfjsLib from 'pdfjs-dist'
import { Document, Packer, Paragraph, TextRun } from 'docx'
import * as XLSX from 'xlsx'
import pptxgen from 'pptxgenjs'
import {
  Upload, FileText, ShieldCheck, Zap, ArrowRight, Search, Moon, Sun, Menu, X,
  CheckCircle2, Download, Layers, Scissors, Image as ImageIcon, Minimize2,
  Sparkles, LockKeyhole, Heart, LoaderCircle, FileImage, FileSpreadsheet,
  Presentation, Brain, MessageCircle, Languages, PenTool, Lock, Unlock, Edit3,
  Files, FileType
} from 'lucide-react'
import './pdf-worker'
type ToolId='compress'|'merge'|'split'|'pdf-jpg'|'jpg-pdf'|'pdf-png'|'pdf-word'|'pdf-excel'|'pdf-ppt'|'summarize'|'chat'|'translate'|'esign'|'protect'|'unlock'|'edit'|'organize'|'pdf-text'
type ToolInfo={id:ToolId;title:string;desc:string;icon:any;color:string;badge?:string}
const tools:ToolInfo[]=[
{id:'compress',title:'Compress PDF',desc:'Reduce PDF file size with browser-based re-saving.',icon:Minimize2,color:'champagne',badge:'Popular'},
{id:'merge',title:'Merge PDF',desc:'Combine multiple PDF files into one.',icon:Layers,color:'emerald'},
{id:'split',title:'Split PDF',desc:'Extract selected pages or page ranges.',icon:Scissors,color:'violet'},
{id:'pdf-jpg',title:'PDF to JPG',desc:'Convert PDF pages into JPG images.',icon:ImageIcon,color:'sapphire'},
{id:'jpg-pdf',title:'JPG to PDF',desc:'Turn JPG and PNG images into a single PDF.',icon:FileText,color:'copper'},
{id:'pdf-png',title:'PDF to PNG',desc:'Convert every PDF page into a PNG image.',icon:FileImage,color:'sapphire'},
{id:'pdf-word',title:'PDF to Word',desc:'Extract PDF text into an editable DOCX document.',icon:FileText,color:'sapphire'},
{id:'pdf-excel',title:'PDF to Excel',desc:'Extract text and table-like rows into XLSX.',icon:FileSpreadsheet,color:'emerald'},
{id:'pdf-ppt',title:'PDF to PowerPoint',desc:'Turn each PDF page into a PowerPoint slide.',icon:Presentation,color:'copper'},
{id:'summarize',title:'AI PDF Summarizer',desc:'Get a summary and key points with Gemini AI.',icon:Brain,color:'violet'},
{id:'chat',title:'Chat with PDF',desc:'Ask Gemini questions about your PDF.',icon:MessageCircle,color:'emerald'},
{id:'translate',title:'PDF Translator',desc:'Translate extracted PDF text with Gemini AI.',icon:Languages,color:'champagne'},
{id:'esign',title:'E-Sign PDF',desc:'Add a typed signature to your PDF.',icon:PenTool,color:'copper'},
{id:'protect',title:'Add Password to PDF',desc:'Encrypt a PDF with a password using qpdf.',icon:Lock,color:'champagne'},
{id:'unlock',title:'Remove PDF Password',desc:'Remove encryption when you know the password.',icon:Unlock,color:'violet'},
{id:'edit',title:'Edit PDF',desc:'Add text to a PDF page.',icon:Edit3,color:'sapphire'},
{id:'organize',title:'Organize PDF',desc:'Reorder, rotate, or remove PDF pages.',icon:Files,color:'emerald'},
{id:'pdf-text',title:'PDF to Text',desc:'Extract selectable text into a TXT file.',icon:FileType,color:'copper'}
]
const seoPages=tools.map(t=>[({compress:'compress-pdf',merge:'merge-pdf',split:'split-pdf','pdf-jpg':'pdf-to-jpg','jpg-pdf':'jpg-to-pdf','pdf-png':'pdf-to-png','pdf-word':'pdf-to-word','pdf-excel':'pdf-to-excel','pdf-ppt':'pdf-to-powerpoint',summarize:'ai-pdf-summarizer',chat:'chat-with-pdf',translate:'pdf-translator',esign:'esign-pdf',protect:'password-protect-pdf',unlock:'unlock-pdf',edit:'edit-pdf',organize:'organize-pdf','pdf-text':'pdf-to-text'} as Record<ToolId,string>)[t.id],t.title,`${t.title} online for free`,t.desc] as const)
function formatBytes(n:number){if(n<1024)return `${n} B`;if(n<1024*1024)return `${(n/1024).toFixed(1)} KB`;return `${(n/1024/1024).toFixed(2)} MB`}
function saveBlob(blob:Blob,name:string){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500)}
async function extractPdfText(file:File,password?:string){const data=await file.arrayBuffer();const pdf=await pdfjsLib.getDocument({data,password}).promise;let full='';for(let i=1;i<=pdf.numPages;i++){const page=await pdf.getPage(i);const content=await page.getTextContent();const line=content.items.map((x:any)=>x.str||'').join(' ');full+=`\\n\\n--- Page ${i} ---\\n${line}`}return {text:full.trim(),pages:pdf.numPages,pdf}}
function App(){
 const [dark,setDark]=useState(true),[active,setActive]=useState<ToolId|null>(null),[files,setFiles]=useState<File[]>([]),[busy,setBusy]=useState(false),[progress,setProgress]=useState(0),[message,setMessage]=useState(''),[splitRange,setSplitRange]=useState('1'),[query,setQuery]=useState(''),[menu,setMenu]=useState(false),[route,setRoute]=useState(location.pathname)
 const [password,setPassword]=useState(''),[signature,setSignature]=useState(''),[editText,setEditText]=useState(''),[pageNumber,setPageNumber]=useState('1'),[pageOrder,setPageOrder]=useState(''),[rotate,setRotate]=useState('0'),[deletePages,setDeletePages]=useState(''),[aiQuestion,setAiQuestion]=useState(''),[language,setLanguage]=useState('English')
 const inputRef=useRef<HTMLInputElement>(null)
 useEffect(()=>{document.documentElement.dataset.theme=dark?'dark':'light'},[dark])
 useEffect(()=>{const onPop=()=>{setRoute(location.pathname);setActive(null)};window.addEventListener('popstate',onPop);return()=>window.removeEventListener('popstate',onPop)},[])
 const go=(path:string)=>{history.pushState({},'',path);setRoute(path);setActive(null);setFiles([]);setMessage('');window.scrollTo({top:0,behavior:'smooth'})}
 const currentSeo=seoPages.find(p=>`/${p[0]}`===route)
 useEffect(()=>{
  const page=currentSeo
  const title=page?`${page[1]} Online Free | OurPDF`:'OurPDF — Free Online PDF Tools'
  const description=page?`${page[3]} Use OurPDF online for free. No signup required.`:'Free online PDF tools to merge, split, compress, convert, organize, and work with PDF files. Try OurPDF for simple browser-based PDF tools.'
  document.title=title
  let meta=document.querySelector('meta[name="description"]') as HTMLMetaElement|null
  if(!meta){meta=document.createElement('meta');meta.name='description';document.head.appendChild(meta)}
  meta.content=description
  let canonical=document.querySelector('link[rel="canonical"]') as HTMLLinkElement|null
  if(!canonical){canonical=document.createElement('link');canonical.rel='canonical';document.head.appendChild(canonical)}
  canonical.href=`https://ourpdf.ta480105.workers.dev${route==='/'?'':route}`
 },[route,currentSeo])
 const filtered=useMemo(()=>tools.filter(t=>(t.title+' '+t.desc).toLowerCase().includes(query.toLowerCase())),[query])
 const openTool=(id:ToolId)=>{setActive(id);setFiles([]);setMessage('');setProgress(0);setPassword('');setSignature('');setEditText('');setAiQuestion('');window.scrollTo({top:0,behavior:'smooth'})}
 const addFiles=(incoming:FileList|null)=>{if(!incoming||!active)return;const arr=Array.from(incoming);const image=active==='jpg-pdf';const valid=arr.filter(f=>image?['image/jpeg','image/png'].includes(f.type):f.type==='application/pdf'||f.name.toLowerCase().endsWith('.pdf'));if(valid.length!==arr.length)setMessage(image?'Only JPG, JPEG, and PNG images are supported.':'Please select PDF files only.');else setMessage('');if(!valid.length)return;setFiles(prev=>active==='merge'||image?[...prev,...valid]:valid.slice(0,1))}
 const process=async()=>{
 if(!active)return
 const needsFile=!['chat','summarize','translate'].includes(active)
 if(needsFile&&!files.length)return
 setBusy(true);setProgress(5);setMessage('')
 try{
 const file=files[0]
 if(active==='protect'||active==='unlock'){
   if(!password)throw new Error('Enter the PDF password.')
   const fd=new FormData();fd.append('file',file);fd.append('password',password);fd.append('action',active==='protect'?'encrypt':'decrypt')
   const res=await fetch('/api/pdf/password',{method:'POST',body:fd});if(!res.ok){const j=await res.json().catch(()=>({}));throw new Error(j.error||'Password service unavailable. Ensure the backend and qpdf are installed.')}
   saveBlob(await res.blob(),active==='protect'?'ourpdf-protected.pdf':'ourpdf-unlocked.pdf')
 }else if(active==='summarize'||active==='chat'||active==='translate'){
   if(!file)throw new Error('Choose a PDF file first.')
   const {text}=await extractPdfText(file)
   if(!text)throw new Error('No selectable text was found in this PDF. Scanned PDFs need OCR before AI tools can read them.')
   const prompt=active==='summarize'?'Summarize this document. Provide a concise overview, key points, and important dates or figures.':active==='chat'?aiQuestion:`Translate the following document into ${language}. Preserve headings and meaning.`
   if(active==='chat'&&!aiQuestion.trim())throw new Error('Enter a question about your PDF.')
   const res=await fetch('/api/ai',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:active,prompt,text:text.slice(0,90000)})})
   const data=await res.json();if(!res.ok)throw new Error(data.error||'Gemini request failed.')
   setMessage(data.text||'No response was returned.')
 }else if(active==='merge'){
   if(files.length<2)throw new Error('Add at least two PDF files to merge.')
   const out=await PDFDocument.create();for(let i=0;i<files.length;i++){const src=await PDFDocument.load(await files[i].arrayBuffer());const pages=await out.copyPages(src,src.getPageIndices());pages.forEach(p=>out.addPage(p));setProgress(10+Math.round((i+1)/files.length*80))}
   saveBlob(new Blob([await out.save() as BlobPart],{type:'application/pdf'}),'ourpdf-merged.pdf')
 }else if(active==='split'){
   const src=await PDFDocument.load(await file.arrayBuffer());const total=src.getPageCount();const pages=splitRange.split(',').flatMap(part=>{const [a,b]=part.trim().split('-').map(Number);if(!a||a<1||a>total)return [];const end=b?Math.min(b,total):a;return Array.from({length:Math.max(0,end-a+1)},(_,i)=>a+i-1)});if(!pages.length)throw new Error(`Enter valid page numbers from 1 to ${total}, e.g. 1-3,5.`)
   const out=await PDFDocument.create();(await out.copyPages(src,[...new Set(pages)])).forEach(p=>out.addPage(p));saveBlob(new Blob([await out.save() as BlobPart],{type:'application/pdf'}),'ourpdf-split.pdf')
 }else if(active==='compress'){
   const src=await PDFDocument.load(await file.arrayBuffer());const bytes=await src.save({useObjectStreams:true,addDefaultPage:false});saveBlob(new Blob([bytes as BlobPart],{type:'application/pdf'}),'ourpdf-compressed.pdf');setMessage(`Finished re-saving. Original: ${formatBytes(file.size)}. New file: ${formatBytes(bytes.length)}. Re-saving does not guarantee a smaller file.`)
 }else if(active==='jpg-pdf'){
   const out=await PDFDocument.create();for(const f of files){const bytes=await f.arrayBuffer();const img=f.type==='image/png'?await out.embedPng(bytes):await out.embedJpg(bytes);const page=out.addPage([img.width,img.height]);page.drawImage(img,{x:0,y:0,width:img.width,height:img.height})}saveBlob(new Blob([await out.save() as BlobPart],{type:'application/pdf'}),'ourpdf-images.pdf')
 }else if(active==='pdf-jpg'||active==='pdf-png'){
   const {pdf,pages}=await extractPdfText(file);for(let i=1;i<=pages;i++){const page=await pdf.getPage(i);const viewport=page.getViewport({scale:1.8});const canvas=document.createElement('canvas');canvas.width=viewport.width;canvas.height=viewport.height;const ctx=canvas.getContext('2d');if(!ctx)throw new Error('Could not create image canvas.');await page.render({canvasContext:ctx,viewport}).promise;const type=active==='pdf-png'?'image/png':'image/jpeg';const blob=await new Promise<Blob>((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('Image conversion failed')),type,0.93));saveBlob(blob,`ourpdf-page-${i}.${active==='pdf-png'?'png':'jpg'}`);setProgress(Math.round(i/pages*90))}setMessage(`${pages} image(s) downloaded. Your browser may ask permission for multiple downloads.`)
 }else if(active==='pdf-text'||active==='pdf-word'||active==='pdf-excel'){
   const {text}=await extractPdfText(file);if(!text)throw new Error('No selectable text found. Scanned PDFs need OCR.')
   if(active==='pdf-text')saveBlob(new Blob([text],{type:'text/plain;charset=utf-8'}),'ourpdf-extracted.txt')
   if(active==='pdf-word'){const doc=new Document({sections:[{children:text.split(/\n+/).map(line=>new Paragraph({children:[new TextRun(line)]}))}]});saveBlob(await Packer.toBlob(doc),'ourpdf-converted.docx')}
   if(active==='pdf-excel'){const rows=text.split('\n').filter(x=>x.trim()).map(line=>line.split(/\s{2,}|\t/));const ws=XLSX.utils.aoa_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Extracted text');const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});saveBlob(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),'ourpdf-extracted.xlsx')}
 }else if(active==='pdf-ppt'){
   const {pdf,pages}=await extractPdfText(file);const ppt=new pptxgen();ppt.layout='LAYOUT_WIDE';for(let i=1;i<=pages;i++){const page=await pdf.getPage(i);const vp=page.getViewport({scale:1.5});const canvas=document.createElement('canvas');canvas.width=vp.width;canvas.height=vp.height;const ctx=canvas.getContext('2d');if(!ctx)throw new Error('Canvas unavailable');await page.render({canvasContext:ctx,viewport:vp}).promise;const slide=ppt.addSlide();slide.addImage({data:canvas.toDataURL('image/png'),x:0,y:0,w:13.333,h:7.5});setProgress(Math.round(i/pages*90))}const blob=await ppt.write({outputType:'blob'}) as Blob;saveBlob(blob,'ourpdf-slides.pptx')
 }else if(active==='esign'||active==='edit'){
   const doc=await PDFDocument.load(await file.arrayBuffer());const pages=doc.getPages();const idx=Math.min(Math.max(1,Number(pageNumber)||1),pages.length)-1;const page=pages[idx];const font=await doc.embedFont(StandardFonts.Helvetica);const value=active==='esign'?signature:editText;if(!value.trim())throw new Error(active==='esign'?'Enter the signature text.':'Enter text to add.')
   page.drawText(value,{x:55,y:65,size:active==='esign'?22:12,font,color:rgb(0.08,0.12,0.2),rotate:degrees(0),maxWidth:page.getWidth()-110});saveBlob(new Blob([await doc.save() as BlobPart],{type:'application/pdf'}),active==='esign'?'ourpdf-signed.pdf':'ourpdf-edited.pdf')
 }else if(active==='organize'){
   const src=await PDFDocument.load(await file.arrayBuffer());const total=src.getPageCount();let order=pageOrder.trim()?pageOrder.split(',').map(x=>Number(x.trim())-1):Array.from({length:total},(_,i)=>i);const removed=new Set(deletePages.split(',').map(x=>Number(x.trim())-1).filter(x=>Number.isInteger(x)&&x>=0&&x<total));order=order.filter((n,i)=>Number.isInteger(n)&&n>=0&&n<total&&!removed.has(n));if(!order.length)throw new Error('Your page order removed every page.');const out=await PDFDocument.create();const copied=await out.copyPages(src,order);copied.forEach(p=>{p.setRotation(degrees(Number(rotate)||0));out.addPage(p)});saveBlob(new Blob([await out.save() as BlobPart],{type:'application/pdf'}),'ourpdf-organized.pdf')
 }
 setProgress(100);if(!['compress','summarize','chat','translate'].includes(active))setMessage('Done! Your file is ready.')
 }catch(e){setMessage(e instanceof Error?e.message:'Something went wrong. Please try another file.')}finally{setBusy(false)}
 }
 const tool=tools.find(t=>t.id===active)
 const toolPath=(id:ToolId)=>'/'+(seoPages.find(p=>p[1]===tools.find(t=>t.id===id)?.title)?.[0]||'')
 const uploadArea=<div className="upload-zone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();addFiles(e.dataTransfer.files)}} onClick={()=>inputRef.current?.click()}>
 <input ref={inputRef} hidden type="file" multiple={active==='merge'||active==='jpg-pdf'} accept={active==='jpg-pdf'?'.jpg,.jpeg,.png,image/jpeg,image/png':'.pdf,application/pdf'} onChange={e=>addFiles(e.target.files)}/>
 <div className="upload-icon"><Upload size={27}/></div><h3>Drop your {active==='jpg-pdf'?'images':'PDF file'} here</h3><p>or <span className="browse">browse files</span> from your device</p><small>{active==='jpg-pdf'?'JPG, JPEG, PNG':'PDF'} · Files stay on your device</small></div>
 const aiTool=active==='summarize'||active==='chat'||active==='translate'
 return <div className="app-shell">
 <header className="header"><div className="header-inner"><button className="brand" onClick={()=>go('/')}><span className="brand-mark"><FileText size={22}/></span><span>Our<span className="brand-accent">PDF</span></span></button><nav className={menu?'nav open':'nav'}><button onClick={()=>{go('/');setMenu(false)}}>Home</button><button onClick={()=>{document.getElementById('tools')?.scrollIntoView({behavior:'smooth'});setMenu(false)}}>All tools</button><a href="#how-it-works" onClick={()=>setMenu(false)}>How it works</a><a href="#privacy" onClick={()=>setMenu(false)}>Privacy</a></nav><div className="header-actions"><button className="icon-btn" aria-label="Toggle theme" onClick={()=>setDark(!dark)}>{dark?<Sun size={19}/>:<Moon size={19}/>}</button><button className="mobile-menu icon-btn" onClick={()=>setMenu(!menu)} aria-label="Menu">{menu?<X/>:<Menu/>}</button><button className="header-cta" onClick={()=>document.getElementById('tools')?.scrollIntoView({behavior:'smooth'})}>Explore tools <ArrowRight size={16}/></button></div></div></header>
 {active&&tool?<main className="tool-page"><button className="back-link" onClick={()=>{setActive(null);setFiles([]);setMessage('')}}>← All tools</button><div className="tool-heading"><div className={`tool-icon ${tool.color}`}><tool.icon size={28}/></div><div><div className="eyebrow">OURPDF TOOL</div><h1>{tool.title} <span className="free-pill">Free</span></h1><p>{tool.desc}</p></div></div><div className="workspace"><div className="workspace-top"><div><strong>{aiTool?'Choose a PDF to work with':'Choose your files'}</strong><p>Secure processing in your browser wherever possible.</p></div><span className="privacy-mini"><ShieldCheck size={15}/> Private by design</span></div>{uploadArea}
 {active==='split'&&<div className="split-field"><label>Pages to extract</label><input value={splitRange} onChange={e=>setSplitRange(e.target.value)} placeholder="Example: 1-3, 5, 8-10"/><small>Use commas for individual pages and hyphens for ranges.</small></div>}
 {active==='protect'&&<div className="split-field"><label>New PDF password</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Create a strong password"/><small>Requires qpdf installed on the backend server.</small></div>}
 {active==='unlock'&&<div className="split-field"><label>Existing PDF password</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter the document password"/><small>Only unlock documents you are authorized to access.</small></div>}
 {active==='esign'&&<div className="split-field"><label>Signature (typed)</label><input value={signature} onChange={e=>setSignature(e.target.value)} placeholder="Type your name"/><label>Page number</label><input value={pageNumber} onChange={e=>setPageNumber(e.target.value)}/></div>}
 {active==='edit'&&<div className="split-field"><label>Text to add</label><input value={editText} onChange={e=>setEditText(e.target.value)} placeholder="Enter text to place near the bottom of the page"/><label>Page number</label><input value={pageNumber} onChange={e=>setPageNumber(e.target.value)}/></div>}
 {active==='organize'&&<div className="split-field"><label>New page order (optional)</label><input value={pageOrder} onChange={e=>setPageOrder(e.target.value)} placeholder="Example: 3,1,2,4"/><small>List original page numbers in the order you want.</small><label>Pages to remove (optional)</label><input value={deletePages} onChange={e=>setDeletePages(e.target.value)} placeholder="Example: 2,5"/><label>Rotate output pages</label><select value={rotate} onChange={e=>setRotate(e.target.value)}><option value="0">No rotation</option><option value="90">90° clockwise</option><option value="180">180°</option><option value="270">270° clockwise</option></select></div>}
 {active==='chat'&&<div className="split-field"><label>Ask a question about the PDF</label><textarea value={aiQuestion} onChange={e=>setAiQuestion(e.target.value)} placeholder="What are the main conclusions? Explain the key terms..." rows={3}/></div>}
 {active==='translate'&&<div className="split-field"><label>Translate into</label><select value={language} onChange={e=>setLanguage(e.target.value)}>{['English','Hindi','Tamil','Telugu','Bengali','Marathi','Gujarati','Kannada','Malayalam','Urdu','French','Spanish','German','Arabic','Japanese'].map(l=><option key={l}>{l}</option>)}</select></div>}
 {files.length>0&&<div className="file-list"><div className="file-list-head"><strong>{files.length} file{files.length>1?'s':''} selected</strong><button onClick={()=>setFiles([])}>Remove all</button></div>{files.map((f,i)=><div className="file-row" key={`${f.name}-${i}`}><span className="file-badge"><FileText size={18}/></span><div className="file-meta"><strong>{f.name}</strong><small>{formatBytes(f.size)}</small></div><button className="remove-file" onClick={()=>setFiles(prev=>prev.filter((_,idx)=>idx!==i))}>×</button></div>)}</div>}
 {busy&&<div className="progress-wrap"><div className="progress-label"><span>Processing in your browser…</span><b>{progress}%</b></div><div className="progress"><span style={{width:`${progress}%`}}/></div></div>}
 {message&&<div className={message.startsWith('Done')||message.startsWith('Finished')||message.includes('downloaded')?'notice success':'notice'}>{message}</div>}
 <button className="process-btn" disabled={busy||(!files.length)||((active==='protect'||active==='unlock')&&!password)} onClick={process}>{busy?<><LoaderCircle className="spin"/> Processing…</>:<>{aiTool?'Run with Gemini':tool.title} <ArrowRight size={18}/></>}</button>
 <p className="workspace-note"><LockKeyhole size={14}/> Files stay in your browser except password tools and Gemini requests. AI tools send extracted text to your configured Gemini backend.</p></div><div className="tool-help"><h2>How to {tool.title.toLowerCase()}</h2><ol><li>Choose your {active==='jpg-pdf'?'JPG or PNG images':'PDF file'} using the uploader above.</li><li>Configure any options and process your document.</li><li>Download your result. No signup required.</li></ol></div></main>
 :currentSeo?<main className="seo-page"><div className="breadcrumb"><button onClick={()=>go('/')}>Home</button> / {currentSeo[1]}</div><div className="seo-hero"><span className="eyebrow">FREE ONLINE PDF TOOL</span><h1>{currentSeo[2]}</h1><p>{currentSeo[3]}</p><button className="primary-btn" onClick={()=>openTool(tools.find(t=>t.title===currentSeo[1])!.id)}>Use this tool <ArrowRight size={17}/></button></div><section className="seo-content"><h2>How it works</h2><p>Choose your document, process it securely, and download the result. AI tools use Gemini through a server-side API key. Conversion quality depends on the source PDF.</p><h2>Frequently asked questions</h2><details><summary>Is OurPDF free to use?</summary><p>Yes. OurPDF tools are free to use. AI requests are subject to your Gemini API quota and any applicable provider terms.</p></details><details><summary>Are my documents uploaded?</summary><p>Most tools process documents locally. AI features send extracted text to the configured Gemini backend; password operations send the PDF to your own backend.</p></details><h2>Explore more tools</h2><div className="related-grid">{tools.map(t=><button key={t.id} onClick={()=>go(toolPath(t.id))}>{t.title} <ArrowRight size={15}/></button>)}</div></section></main>
 :<><main><section className="hero"><div className="hero-copy"><div className="hero-tag"><Sparkles size={14}/> YOUR EVERYDAY PDF TOOLKIT</div><h1>PDF work made<br/><span>ridiculously simple.</span></h1><p>Everything you need to work with PDFs. Free, easy to use, and designed with your privacy in mind.</p><div className="hero-buttons"><button className="primary-btn" onClick={()=>document.getElementById('tools')?.scrollIntoView({behavior:'smooth'})}>Explore PDF tools <ArrowRight size={18}/></button><button className="secondary-btn" onClick={()=>document.getElementById('how-it-works')?.scrollIntoView({behavior:'smooth'})}>How it works</button></div><div className="hero-trust"><span><ShieldCheck size={16}/> Files stay on your device</span><span><CheckCircle2 size={16}/> No signup required</span></div></div><div className="hero-art"><div className="art-glow"/><div className="floating-card card-back"><div className="mini-lines"><i/><i/><i/></div><span>PDF · 2.4 MB</span></div><div className="pdf-card"><div className="pdf-ribbon">PDF</div><div className="pdf-art-icon"><FileText size={58}/></div><div className="pdf-lines"><i/><i/><i/><i/></div><div className="pdf-art-footer"><span>OURPDF</span><span>01 / 12</span></div></div><div className="floating-card card-front"><div className="check-circle"><CheckCircle2 size={19}/></div><div><strong>Ready to go!</strong><small>Your file is processed</small></div></div><div className="sparkle s1">✦</div><div className="sparkle s2">✧</div></div></section>
 <section className="trust-strip"><div><ShieldCheck/><strong>Privacy first</strong><span>Files stay on your device</span></div><div><Zap/><strong>Fast & simple</strong><span>No complicated steps</span></div><div><Heart/><strong>Free to use</strong><span>No account required</span></div></section>
 <section className="tools-section" id="tools"><div className="section-heading"><div><div className="eyebrow">THE PDF TOOLKIT</div><h2>One place. All your PDF needs.</h2><p>Pick a tool and get things done in a few clicks.</p></div><div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Find a tool…"/></div></div><div className="tools-grid">{filtered.map(t=><button className="tool-card" key={t.id} onClick={()=>openTool(t.id)}><div className={`tool-icon ${t.color}`}><t.icon size={23}/></div>{t.badge&&<span className="popular-badge">{t.badge}</span>}<h3>{t.title}</h3><p>{t.desc}</p><span className="tool-card-bottom">Use tool <ArrowRight size={16}/></span></button>)}</div>{filtered.length===0&&<div className="empty-search">No tools found. Try a different search.</div>}</section>
 <section className="how-section" id="how-it-works"><div className="how-intro"><div className="eyebrow">NO LEARNING CURVE</div><h2>Three steps.<br/><span>That’s all it takes.</span></h2><p>We believe PDF tools should be simple. No clutter, no confusing menus, just get it done.</p></div><div className="steps"><div className="step"><span>01</span><div className="step-icon"><Upload/></div><h3>Choose your file</h3><p>Drag and drop or browse your device.</p></div><div className="step"><span>02</span><div className="step-icon"><Zap/></div><h3>Let OurPDF work</h3><p>We process your file right in your browser.</p></div><div className="step"><span>03</span><div className="step-icon"><Download/></div><h3>Download & done</h3><p>Get your file and get on with your day.</p></div></div></section>
 <section className="privacy-section" id="privacy"><div className="privacy-icon"><ShieldCheck size={32}/></div><div><div className="eyebrow">YOUR FILES ARE YOURS</div><h2>Privacy isn’t a feature.<br/>It’s the default.</h2><p>Most OurPDF tools run directly in your browser. Gemini AI tools send extracted text to your configured backend. Password operations use your own backend. Files are not stored by OurPDF by default.</p></div><div className="privacy-checks"><span><CheckCircle2/> Local processing where possible</span><span><CheckCircle2/> No document storage by default</span><span><CheckCircle2/> Secure server-side AI key</span></div></section>
 <section className="seo-tools"><div className="section-heading"><div><div className="eyebrow">THE COMPLETE TOOLKIT</div><h2>More tools. Same simplicity.</h2><p>All 18 PDF tools are now available in the toolkit.</p></div></div><div className="roadmap-grid">{tools.map(t=><button key={t.id} onClick={()=>openTool(t.id)}><span>{t.title}</span><span className="coming-small">Available</span><ArrowRight size={16}/></button>)}</div></section>
 </main><footer className="footer"><div className="footer-top"><button className="brand" onClick={()=>go('/')}><span className="brand-mark"><FileText size={21}/></span><span>Our<span className="brand-accent">PDF</span></span></button><p>Your PDFs. Your Tools. Your Privacy.</p></div><div className="footer-bottom"><span>© {new Date().getFullYear()} OurPDF. Made for simpler PDF days.</span><div><a href="#privacy">Privacy</a><a href="#tools">Tools</a><a href="mailto:hello@ourpdf.example">Contact</a></div></div></footer></>}
 </div>
}
export default App
