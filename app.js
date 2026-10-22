const CFG={model:'gemini-2.5-flash',api:'https://generativelanguage.googleapis.com/v1beta/models/'}; // verifica modelo y cuota gratuita vigentes en ai.google.dev
const $=s=>document.querySelector(s);let DATA=[],RES=null,HIST=[],CTX='',CH=null;
const pct=v=>{if(typeof v==='number')return v;const n=parseFloat(String(v).replace('%','').replace(/\./g,'').replace(',','.'));return isNaN(n)?null:n};
const save=(b,n)=>{const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=n;a.click()};
const say=(t,e)=>{$('#st').textContent=t;$('#st').className=e?'err':''};
$('#key').value=localStorage.getItem('gk')||'';
async function loadCtx(){CTX=(await Promise.all(['prompts/system_prompt.md','knowledge/normas.md','knowledge/innovacion.md','knowledge/kpis.md'].map(f=>fetch(f).then(r=>r.text())))).join('\n\n---\n\n')}
function parse(rows){let p='',o=[];for(const r of rows){const b=String(r[1]||'');if(b.startsWith('Proceso:')){p=b.slice(8).trim();continue}
 if(b&&b!=='Indicador'&&typeof r[8]==='string'&&r[8].includes('%'))o.push({proceso:p,indicador:b,tendencia:r[2],planeado:r[3],ejecutado:r[4],fecha:String(r[5]).slice(0,10),resultado:pct(r[6]),meta:pct(r[7]),cumplimiento:pct(r[8])})}return o}
function summ(d){const m={};d.forEach(x=>{const v=m[x.proceso]??={n:0,s:0,i:0};v.n++;v.s+=x.cumplimiento;if(x.cumplimiento<100)v.i++});
 return Object.entries(m).map(([p,v])=>({proceso:p,indicadores:v.n,cumplimiento_prom:+(v.s/v.n).toFixed(1),incumplidos:v.i}))}
async function gem(contents,json){const k=$('#key').value.trim();if(!k)throw Error('Ingresa tu API key de Gemini');localStorage.setItem('gk',k);
 const r=await fetch(`${CFG.api}${CFG.model}:generateContent?key=${k}`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({system_instruction:{parts:[{text:CTX}]},contents,generationConfig:{temperature:.2,...(json?{responseMimeType:'application/json'}:{})}})});
 const j=await r.json();if(!r.ok)throw Error(j.error?.message||r.status);return j.candidates[0].content.parts.map(p=>p.text).join('')}
$('#file').onchange=async e=>{try{say('Leyendo…');await loadCtx();const wb=XLSX.read(await e.target.files[0].arrayBuffer());
 DATA=parse(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''}));if(!DATA.length)throw Error('No encontré indicadores con el formato esperado.');
 const s=summ(DATA);CH?.destroy();CH=new Chart($('#ch'),{type:'bar',data:{labels:s.map(x=>x.proceso.slice(0,28)),datasets:[{label:'% cumplimiento promedio',data:s.map(x=>x.cumplimiento_prom),backgroundColor:'#0f766e'}]},options:{animation:false}});
 HIST=[{role:'user',parts:[{text:'DATOS DEL INFORME (JSON):'+JSON.stringify(DATA)}]},{role:'model',parts:[{text:'Datos recibidos.'}]}];
 $('#go').disabled=$('#ask').disabled=false;say(`${DATA.length} indicadores en ${s.length} procesos.`)}catch(x){say(x.message,1)}};
const SCHEMA='Analiza TODO el informe. Responde SOLO JSON: {"resumen":str,"hallazgos":[{"proceso":str,"indicador":str,"dato":str,"lectura":str}],"tendencias":[str],"efectividad":[{"indicador":str,"evaluacion":"efectiva|parcial|no efectiva|sin evidencia","evidencia":str}],"acciones":[{"accion":str,"indicador":str,"norma_relacionada":str,"plazo_sugerido":str,"indicador_de_eficacia":str}],"riesgos_sgi":[str],"decisiones":[str]}';
$('#go').onclick=async()=>{try{say('Analizando…');RES=JSON.parse(await gem([...HIST,{role:'user',parts:[{text:SCHEMA}]}],1));
 const L=(t,a)=>`<h3>${t}</h3><ul>${a.map(x=>`<li>${x}</li>`).join('')}</ul>`;
 $('#res').innerHTML=`<p>${RES.resumen}</p>`+L('Hallazgos',RES.hallazgos.map(h=>`<b>${h.indicador}</b>: ${h.dato}. ${h.lectura}`))+L('Tendencias',RES.tendencias)+L('Efectividad de acciones',RES.efectividad.map(e=>`<b>${e.indicador}</b>: ${e.evaluacion}. ${e.evidencia}`))+L('Acciones propuestas',RES.acciones.map(a=>`${a.accion} <i>(${a.norma_relacionada}; ${a.plazo_sugerido})</i>`))+L('Riesgos al SGI',RES.riesgos_sgi)+L('Decisiones',RES.decisiones);
 $('#out').hidden=false;say('Listo. Revisa antes de usar: son propuestas sujetas a aprobación.')}catch(x){say(x.message,1)}};
$('#ask').onclick=async()=>{const q=$('#q').value.trim();if(!q)return;const add=(w,t)=>{const d=document.createElement('div');d.className='m';d.textContent=w+': '+t;$('#log').append(d)};
 add('Tú',q);HIST.push({role:'user',parts:[{text:q}]});try{const a=await gem(HIST);HIST.push({role:'model',parts:[{text:a}]});add('Agente',a)}catch(x){add('Error',x.message)}$('#q').value=''};
$('#bx').onclick=()=>{const wb=XLSX.utils.book_new();[['Resumen',summ(DATA)],['Detalle',DATA],['Hallazgos',RES.hallazgos],['Efectividad',RES.efectividad],['Acciones',RES.acciones]].forEach(([n,a])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(a),n));XLSX.writeFile(wb,'Informe_SGI.xlsx')};
$('#bw').onclick=async()=>{const{Document,Packer,Paragraph,HeadingLevel,ImageRun}=docx,H=t=>new Paragraph({text:t,heading:HeadingLevel.HEADING_1}),B=a=>a.map(t=>new Paragraph({text:t,bullet:{level:0}}));
 const img=Uint8Array.from(atob(CH.toBase64Image().split(',')[1]),c=>c.charCodeAt(0));
 const d=new Document({sections:[{children:[new Paragraph({text:'Informe ejecutivo SGI',heading:HeadingLevel.TITLE}),new Paragraph(RES.resumen),new Paragraph({children:[new ImageRun({data:img,transformation:{width:560,height:260}})]}),
 H('Hallazgos'),...B(RES.hallazgos.map(h=>`${h.indicador}: ${h.dato}. ${h.lectura}`)),H('Tendencias'),...B(RES.tendencias),H('Efectividad de acciones'),...B(RES.efectividad.map(e=>`${e.indicador}: ${e.evaluacion}. ${e.evidencia}`)),
 H('Acciones propuestas'),...B(RES.acciones.map(a=>`${a.accion} (${a.indicador}; ${a.norma_relacionada}; ${a.plazo_sugerido})`)),H('Riesgos al SGI'),...B(RES.riesgos_sgi),H('Decisiones requeridas'),...B(RES.decisiones)]}]});
 save(await Packer.toBlob(d),'Informe_Ejecutivo_SGI.docx')};
$('#bp').onclick=()=>{const p=new PptxGenJS(),s=(t,a)=>{const l=p.addSlide();l.addText(t,{x:.5,y:.3,w:9,h:.6,fontSize:24,bold:true,color:'14213D'});if(a)l.addText(a.slice(0,6).map(x=>({text:x,options:{bullet:true,breakLine:true}})),{x:.5,y:1.1,w:9,h:4.2,fontSize:14,valign:'top'});return l};
 s('Informe de gestión SGI',[RES.resumen]);s('Cumplimiento por proceso').addImage({data:CH.toBase64Image(),x:.5,y:1.1,w:9,h:4});
 s('Hallazgos',RES.hallazgos.map(h=>`${h.indicador}: ${h.dato}`));s('Tendencias',RES.tendencias);s('Efectividad de acciones',RES.efectividad.map(e=>`${e.indicador}: ${e.evaluacion}`));
 s('Acciones propuestas',RES.acciones.map(a=>`${a.accion} (${a.plazo_sugerido})`));s('Decisiones',RES.decisiones);p.writeFile({fileName:'Presentacion_SGI.pptx'})};
