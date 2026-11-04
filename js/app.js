const $=s=>document.querySelector(s);
const MON={JAN:0,FEB:1,MAR:2,APR:3,MEI:4,MAY:4,JUN:5,JUL:6,AGU:7,AUG:7,SEP:8,OKT:9,OCT:9,NOV:10,DES:11,DEC:11};
const DONE=["done","tolak / batal"],TOLAK="tolak / batal";
const DEF={"pejabat toko":"Proxy approval Pejabat Toko di Asset Online","am":"Follow-up AM untuk approval","atasan penyerah":"Follow-up Atasan Penyerah","atasan langsung":"Follow-up Atasan langsung","konfirmasi pejabat toko":"Proxy konfirmasi Pejabat Toko","validasi accounting":"Follow-up validasi Accounting","checker fa":"Follow-up tim FA (Checker FA)","gsm":"Follow-up GSM untuk approval","user penerima":"Follow-up user penerima untuk konfirmasi"};
let VIEW=false,D=null,tab="tk",shareSel="",shMode="all",selK="",F={q:"",p:"",a:"",c:"",j:""};const badge=a=>a==null?"-":`<span class="bd ${a>=+$("#kr").value?"K":a>=+$("#ti").value?"T":""}">${a} hari</span>`;
const norm=v=>String(v==null?"":v).trim();
const lc=v=>norm(v).toLowerCase();
const WH={"pejabat toko":"Toko","konfirmasi pejabat toko":"Toko","user penerima":"Toko","atasan penyerah":"Toko","am":"AM","checker fa":"Tim FA","gsm":"GSM","validasi accounting":"Accounting","atasan langsung":"Atasan langsung"},who=r=>r.aksi?"<b>"+(WH[r.s]||"Toko")+"</b>":(WH[r.s]||"Cek Asset Online"),COL={bs:false,hs:false},H=(id,t)=>`<h2 class="tg"><button class="tgb" data-id="${id}" title="Tampilkan / sembunyikan">${COL[id]?"+":"−"}</button> ${t}</h2>`,pmn=r=>hp(r)?norm(r.nopm):"-",pmt=r=>hp(r)?(fd(r.tsel)||"-"):"-",hp=r=>r.mod==="Mutasi"&&norm(r.nopm),nd=r=>hp(r)?`<b>PM: ${r.nopm}</b><br><span class="sm">Dok: ${norm(r.nttb)}</span>`:norm(r.nttb),td=r=>hp(r)?(fd(r.tsel)||"-"):(fd(r.tgl)||"-"),pmv=r=>r.mod==="Mutasi"?[r.nopm||"-",fd(r.tsel)||"-"]:["-","-"],mk=d=>d.toISOString().slice(0,7),ML=k=>["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"][+k.slice(5)-1]+" "+k.slice(0,4),fd=v=>{const d=toDate(v);return d?d.toISOString().slice(0,10):""};
function toDate(v){
 if(v instanceof Date)return v;
 if(typeof v==="number"&&v>20000)return new Date(Math.round((v-25569)*864e5));
 const s=norm(v).toUpperCase();let m=s.match(/^(\d{1,2})[-\/ ]([A-Z]{3})[-\/ ](\d{2,4})$/);
 if(m&&m[2] in MON)return new Date(Date.UTC(m[3].length==2?2000+ +m[3]:+m[3],MON[m[2]],+m[1]));
 m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);return m?new Date(Date.UTC(+m[1],m[2]-1,+m[3])):null;
}
function findHeader(rows,key){for(let i=0;i<Math.min(rows.length,30);i++)if(rows[i].some(c=>lc(c)===key))return i;return -1}
function colIdx(h,names){for(const n of names){const i=h.findIndex(c=>lc(c)===n);if(i>=0)return i}return -1}
function readSheet(wb,name,key,map){
 const ws=wb.Sheets[name];if(!ws)throw new Error("Sheet tidak ada: "+name);
 const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
 const hi=findHeader(rows,key);if(hi<0)throw new Error("Header '"+key+"' tidak ditemukan di "+name);
 const h=rows[hi],ix={};for(const k in map){ix[k]=colIdx(h,map[k]);}
 const out=[];
 for(let i=hi+1;i<rows.length;i++){const r=rows[i];const o={_row:i+1};let any=false;
  for(const k in ix){o[k]=ix[k]>=0?r[ix[k]]:"";if(norm(o[k])!=="")any=true}
  if(any&&norm(o.status)!=="")out.push(o)}
 return out;
}
function load(buf){
 const wb=XLSX.read(buf,{type:"array",cellDates:true});
 const akt=readSheet(wb,"UPDATE AKTIVASI","no nttb",{tgl:["tgl nttb","tanggal","tgl"],nttb:["no nttb"],toko:["uu tujuan","kode toko","uu"],tsel:["tanggal terima"],tag:["no. tag","no tag"],barang:["nama barang","nama asset","nama aset"],status:["status"]});
 const mut=readSheet(wb,"UPDATE MUTASI","no pengajuan",{tgl:["tgl pengajuan","tanggal pengajuan","tgl"],nttb:["no pengajuan"],asal:["uu awal"],toko:["uu tujuan"],tsel:["tgl. pm"],nopm:["no pm"],tag:["no tag","no. tag"],barang:["nama barang","nama asset","nama aset"],status:["status"]});
 const str=[];const ws=wb.Sheets["UPDATE STRUKTUR"];
 if(ws){const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});const hi=findHeader(rows,"kode toko");
  if(hi>=0){const h=rows[hi],ik=colIdx(h,["kode toko"]),inm=colIdx(h,["nama toko"]),iam=colIdx(h,["nama am","am"]),iac=colIdx(h,["nama ac","ac"]);
   for(let i=hi+1;i<rows.length;i++)if(norm(rows[i][ik]))str.push({kode:norm(rows[i][ik]),nama:norm(rows[i][inm]),am:iam>=0?norm(rows[i][iam]):"",ac:iac>=0?norm(rows[i][iac]):"",pj:""})}}
 const par={};const wp=wb.Sheets["PARAMETER"];
 if(wp)XLSX.utils.sheet_to_json(wp,{header:1,defval:""}).forEach(r=>{if(norm(r[0])&&norm(r[1]))par[lc(r[0])]=norm(r[1])});
 let rf=null;const wr=wb.Sheets["RINGKASAN"];
 if(wr){const rows=XLSX.utils.sheet_to_json(wr,{header:1,raw:true,defval:""});
  rows.slice(0,10).forEach(r=>{const l=lc(r[0]);if(l.includes("tanggal acuan aging")){const d=toDate(r[1]);if(d)rf=d}
   if(l.startsWith("batas kritis")&&+r[1])$("#kr").value=+r[1];if(l.startsWith("batas tinggi")&&+r[1])$("#ti").value=+r[1]})}
 const pj={},wj=wb.Sheets["PEJABAT TOKO"];
 if(wj){const rows=XLSX.utils.sheet_to_json(wj,{header:1,defval:""}),hi=findHeader(rows,"kode toko");if(hi>=0){const hd=rows[hi],ik=colIdx(hd,["kode toko"]),ip=hd.findIndex(c=>lc(c).includes("pejabat"));
  if(ip>=0)for(let i=hi+1;i<rows.length;i++)if(norm(rows[i][ik]))pj[norm(rows[i][ik]).toUpperCase()]=norm(rows[i][ip])}}
 str.forEach(s=>s.pj=pj[s.kode.toUpperCase()]||"");
 start({akt,mut,str,par,rf});
}
function start(d){D=d;const{akt,mut,str}=d,rf=d.rf?new Date(d.rf):null;if(d.kr)$("#kr").value=d.kr;if(d.ti)$("#ti").value=d.ti;
 if(VIEW){const e=document.querySelector("header .hd>div>div");if(e)e.textContent="Data per "+GIT_DATA.label+" · mode lihat saja"}
 const all=akt.concat(mut).map(r=>toDate(r.tgl)).filter(Boolean);
 const mx=all.length?new Date(Math.max(...all)):new Date();
 $("#ref").value=(rf||mx).toISOString().slice(0,10);
 const ms=[...new Set(akt.concat(mut).map(r=>{const d=toDate(r.tgl);return d?mk(d):""}).filter(Boolean))].sort().reverse(),cur=(rf||mx).toISOString().slice(0,7);
 $("#bln").innerHTML='<option value="">Semua bulan</option>'+ms.map(m=>`<option value="${m}">${ML(m)}</option>`).join("");$("#bln").value=ms.includes(cur)?cur:(ms[0]||"");
 $("#info").textContent=`Aktivasi ${akt.length.toLocaleString("id")} baris, mutasi ${mut.length.toLocaleString("id")} baris, toko ${str.length}.`;
 render();
}
function analyse(){
 const ref=new Date($("#ref").value+"T00:00:00Z"),kr=+$("#kr").value,ti=+$("#ti").value;
 const kode=new Set(D.str.map(s=>s.kode.toUpperCase())),info=Object.fromEntries(D.str.map(s=>[s.kode,s]));
 const proc=(arr,mod)=>arr.map(r=>{
  const st=norm(r.status),s=lc(st),d=toDate(r.tgl);
  const ds=toDate(r.tsel);return {...r,mod,st,s,bln:d?mk(d):"",bs:ds&&s==="done"?mk(ds):"",pend:!DONE.includes(s),tolak:s===TOLAK,age:d?Math.max(0,Math.floor((ref-d)/864e5)):null,dkey:lc(r.tag)}});
 const A=proc(D.akt,"Aktivasi"),M=proc(D.mut,"Mutasi");
 for(const L of[A,M]){
  const g={};L.forEach(r=>{if(r.dkey)(g[r.dkey]=g[r.dkey]||[]).push(r)});
  L.forEach(r=>{r.jenis="";const G=g[r.dkey]||[];if(G.length<2)return;
   const nt=G.filter(x=>!x.tolak).length,pd=G.filter(x=>x.pend).length;
   if(r.pend&&pd>1)r.jenis="GANDA AKTIF";
   else if(r.tolak)r.jenis=nt>0?"TOLAK DIAJUKAN ULANG":"TAG GANDA";
   else r.jenis=G.some(x=>x.tolak)?"PENGAJUAN ULANG":"TAG GANDA"});
 }
const TK=new Set(["pejabat toko","konfirmasi pejabat toko","user penerima","atasan penyerah"]),U=v=>norm(v).toUpperCase();
 for(const r of A.concat(M)){r.aksi=TK.has(r.s);r.tp=kode.has(U(r.toko))?U(r.toko):kode.has(U(r.asal))?U(r.asal):U(r.toko)}
 return{A,M,kr,ti,kode,info};
}
const ageCls=(a,kr,ti)=>a==null?"":a>=kr?"K":a>=ti?"T":"";
const fmt=n=>n.toLocaleString("id");
function table(head,rows,numCols=[]){return `<div class="w"><table><thead><tr>${head.map((h,i)=>`<th class="${numCols.includes(i)?"n":""}">${h}</th>`).join("")}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map((c,i)=>`<td class="${numCols.includes(i)?"n":""}" data-l="${head[i]}">${c}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`}
function render(){
 if(!D)return;$("#out").style.display="";
 const X=analyse(),{kr,ti,kode,info}=X,mf=$("#bln").value,inM=r=>!mf||r.bln===mf,A=X.A.filter(inM),M=X.M.filter(inM),AA=X.A,MM=X.M,lb=mf?ML(mf):"Semua bulan";
 const pa=A.filter(r=>r.pend),pm=M.filter(r=>r.pend);
 const cnt=(L,j)=>L.filter(r=>r.jenis===j).length;
 const unk=[...new Set(pa.concat(pm).map(r=>r.s))].filter(s=>!(s in DEF)&&!(s in D.par));
 $("#warn").innerHTML=unk.length?`<div class="warn">Status baru belum ada di PARAMETER: <b>${unk.join(", ")}</b>. Tetap dihitung sebagai belum selesai.</div>`:"";
 const sa={};pa.concat(pm).forEach(r=>{if(r.tp)sa[r.tp]=Math.max(sa[r.tp]||0,r.age||0)});const sv=Object.values(sa);
 const kp=[["Toko dengan item belum selesai",sv.length,""],["Aktivasi belum selesai",pa.length,""],["Mutasi belum selesai",pm.length,""],["Toko KRITIS (>= "+kr+" hari)",sv.filter(a=>a>=kr).length,"k"],["Toko TINGGI (>= "+ti+" hari)",sv.filter(a=>a>=ti&&a<kr).length,"t"],["Item butuh aksi toko",pa.concat(pm).filter(r=>r.aksi).length,""],["Aging tertua · "+lb,Math.max(0,...pa.concat(pm).map(r=>r.age||0)),"t"," hari"],["Aging tertua · semua bulan",Math.max(0,...AA.concat(MM).filter(r=>r.pend).map(r=>r.age||0)),""," hari"]];
 $("#kpi").innerHTML=kp.map(k=>`<div class="${k[2]}"><b>${fmt(k[1])}${k[3]||""}</b><span>${k[0]}</span></div>`).join("");
 const pg=(L,n)=>{const t=L.length,d=L.filter(r=>r.s==="done").length,tk=L.filter(r=>r.tolak).length,b=t-d-tk,pc=v=>t?v*100/t:0,v2=t-tk,px=v=>(v2?v*100/v2:0).toFixed(1);
  return `<div class="card"><div class="bar"><h2 style="margin:0">${n} · ${lb}</h2><span style="flex:1"></span><span class="big">${pc(d).toFixed(1)}%</span><span class="sm">selesai</span></div>
  <div class="pb"><i style="width:${pc(d)}%;background:var(--g)"></i><i style="width:${pc(b)}%;background:var(--o)"></i><i style="width:${pc(tk)}%;background:#9aa6b8"></i></div>
  <div class="lg"><span><s style="background:var(--g)"></s>Selesai (Done) <b>${fmt(d)}</b> · ${pc(d).toFixed(1)}%</span><span><s style="background:var(--o)"></s>Belum selesai <b>${fmt(b)}</b> · ${pc(b).toFixed(1)}%</span><span><s style="background:#9aa6b8"></s>Tolak/Batal <b>${fmt(tk)}</b> · ${pc(tk).toFixed(1)}%</span></div>
  <div class="sm">Dari ${fmt(t)} pengajuan. Tanpa Tolak/Batal: ${px(d)}% selesai, ${px(b)}% belum selesai.</div></div>`};
 $("#prog").innerHTML=`<div class="g2">${pg(A,"Aktivasi")}${pg(M,"Mutasi")}</div>`;
 const p=$("#pane");
 if(tab==="st"){
  const sts={};A.concat(M).forEach(r=>{const o=sts[r.s]=sts[r.s]||{n:r.st,a:0,m:0,aa:null,ma:null};
   if(r.mod==="Aktivasi"){o.a++;if(r.pend&&r.age!=null)o.aa=Math.max(o.aa??0,r.age)}else{o.m++;if(r.pend&&r.age!=null)o.ma=Math.max(o.ma??0,r.age)}});
  const rows=Object.values(sts).sort((x,y)=>(DONE.includes(lc(x.n))-DONE.includes(lc(y.n)))||(y.a+y.m)-(x.a+x.m)).map(o=>{const k=lc(o.n);
   return[o.n+(DONE.includes(k)?' <span class="tag">selesai</span>':""),fmt(o.a),o.aa==null?"-":`<span class="${ageCls(o.aa,kr,ti)}">${o.aa}</span>`,fmt(o.m),o.ma==null?"-":`<span class="${ageCls(o.ma,kr,ti)}">${o.ma}</span>`,DONE.includes(k)?"-":(D.par[k]||DEF[k]||"Cek status di Asset Online")]});
  p.innerHTML="<h2>Posisi proses semua status</h2>"+table(["Status","Aktivasi","Aging tertua (hari)","Mutasi","Aging tertua (hari)","Tindak lanjut"],rows,[1,2,3,4]);
 }else if(tab==="dp"){
  const js=["TOLAK DIAJUKAN ULANG","PENGAJUAN ULANG","GANDA AKTIF","TAG GANDA"];
  const sum=js.map(j=>[j,fmt(cnt(A,j)),fmt(cnt(M,j))]);
  sum.push(["Total Tolak/Batal",fmt(A.filter(r=>r.tolak).length),fmt(M.filter(r=>r.tolak).length)]);
  sum.push(["Tolak/Batal belum diajukan ulang",fmt(A.filter(r=>r.tolak).length-cnt(A,"TOLAK DIAJUKAN ULANG")),fmt(M.filter(r=>r.tolak).length-cnt(M,"TOLAK DIAJUKAN ULANG"))]);
  p.innerHTML=`<h2>Duplikat &amp; Tolak</h2>${table(["Jenis","Aktivasi","Mutasi"],sum,[1,2])}<br>
  <div class="bar"><select id="fm"><option>Aktivasi</option><option>Mutasi</option></select><select id="fj"><option value="">Selain TAG GANDA</option>${js.map(j=>`<option>${j}</option>`).join("")}</select><input id="fq" placeholder="Cari tag / toko / barang" style="flex:1"><button id="ex" class="alt">Unduh CSV</button></div><div id="dl"></div>`;
  const draw=()=>{const L=$("#fm").value==="Aktivasi"?A:M,j=$("#fj").value,q=lc($("#fq").value);
   const rs=L.filter(r=>r.jenis&&(j?r.jenis===j:r.jenis!=="TAG GANDA")&&(!q||lc([r.tag,r.toko,r.barang,r.nttb].join(" ")).includes(q)));
   X.cur=rs;$("#dl").innerHTML=`<div class="sub">${fmt(rs.length)} baris${rs.length>500?" (tampil 500 pertama, unduh CSV untuk semua)":""}</div>`+
   table(["Jenis","Tgl","No dokumen","Toko","No tag","Barang","Status","Baris"],rs.slice(0,500).map(r=>[r.jenis,(toDate(r.tgl)||"").toString()?(toDate(r.tgl)?toDate(r.tgl).toISOString().slice(0,10):""):"",r.nttb,r.toko,r.tag,r.barang,r.st,r._row]))};
  ["fm","fj"].forEach(i=>$("#"+i).onchange=draw);$("#fq").oninput=draw;draw();
  $("#ex").onclick=()=>{const h="jenis;modul;no_dokumen;toko;no_tag;barang;status;baris\n",b=X.cur.map(r=>[r.jenis,r.mod,r.nttb,r.toko,r.tag,r.barang,r.st,r._row].map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(";")).join("\n");
   saveFile("duplikat_tolak.csv","\ufeff"+h+b)};
 }else if(tab==="tk"){
  const all=AA.concat(MM),bk={};all.filter(r=>r.pend&&r.tp).forEach(r=>{const o=bk[r.tp]=bk[r.tp]||{n:0,age:0};o.n++;if(r.age>o.age)o.age=r.age});
  const fin=all.filter(r=>r.s==="done"&&inM(r)),pp=pa.concat(pm);
  const ks=[...new Set(Object.keys(info).concat(pp.map(r=>r.tp).filter(Boolean)))];
  const S=ks.map(k=>{const s=info[k]||{},x=pp.filter(r=>r.tp===k),a=x.filter(r=>r.mod==="Aktivasi").length,f=fin.filter(r=>r.tp===k),age=x.length?Math.max(...x.map(r=>r.age||0)):-1;
   return{k,nama:s.nama||"Branch Palu",am:s.am||"-",ac:s.ac||"-",pj:s.pj||"-",x,a,m:x.length-a,age,bk:(bk[k]||{}).age,bn:(bk[k]||{}).n||0,f,fa:f.filter(r=>r.mod==="Aktivasi").length,fm:f.filter(r=>r.mod==="Mutasi").length,
    pr:!x.length?((bk[k]||{}).n?"BACKLOG":"BERSIH"):age>=kr?"KRITIS":age>=ti?"TINGGI":"NORMAL"}}).sort((u,v)=>v.age-u.age||u.k.localeCompare(v.k));
  const opt=f=>[...new Set(S.map(s=>s[f]))].sort().map(v=>`<option>${v}</option>`).join("");let lastR=[];
  p.innerHTML=`<h2>Cari toko · ${lb}</h2><div class="sub">Ketik kode atau nama toko: saran muncul, pilih atau tekan Enter, profil toko langsung tampil dan tetap tampil walau kotak pencarian dikosongkan (klik Tutup untuk menyembunyikan). Klik judul kolom tabel untuk mengurutkan naik/turun.</div><div class="bar"><div class="qw"><input id="q" placeholder="Ketik kode atau nama toko..." autocomplete="off"><div id="sg" class="sg"></div></div><select id="fp"><option value="">Semua prioritas</option><option>KRITIS</option><option>TINGGI</option><option>NORMAL</option><option>BACKLOG</option><option>BERSIH</option></select><select id="fa"><option value="">Semua AM</option>${opt("am")}</select><select id="fc"><option value="">Semua AC</option>${opt("ac")}</select><select id="fj"><option value="">Aktivasi &amp; Mutasi</option><option value="a">Ada aktivasi belum selesai</option><option value="m">Ada mutasi belum selesai</option></select></div><div id="det"></div><div id="res"></div>`;
  $("#q").value=F.q;$("#fp").value=F.p;$("#fa").value=F.a;$("#fc").value=F.c;$("#fj").value=F.j;
  const det=k=>{selK=k;const s=S.find(z=>z.k===k);if(!s){$("#det").innerHTML="";return}
   const x=s.x.slice().sort((u,v)=>(v.age||0)-(u.age||0)),f=s.f.slice().sort((u,v)=>(v.bs||v.bln||"").localeCompare(u.bs||u.bln||""));
   $("#det").innerHTML=`<div class="card" style="margin-top:12px"><div class="bar"><h2 style="margin:0">${s.k} · ${s.nama}</h2><span style="flex:1"></span>${x.length&&!VIEW?'<button id="sb">Share ke toko</button>':""}<button id="xs" class="alt">Unduh Excel toko ini</button><button id="cl" class="alt">Tutup</button></div>
   <div class="bar" style="margin:8px 0"><span class="chip">AC: <b>${s.ac}</b></span><span class="chip">AM: <b>${s.am}</b></span><span class="chip">Pejabat Toko: <b>${s.pj}</b></span></div>
   <div class="kpi"><div class="t"><b>${s.a}</b><span>Aktivasi belum selesai</span></div><div class="t"><b>${s.m}</b><span>Mutasi belum selesai</span></div><div><b>${s.fa}</b><span>Aktivasi selesai</span></div><div><b>${s.fm}</b><span>Mutasi selesai</span></div></div>
   <div class="sub">Periode: ${lb}. Backlog semua bulan: ${s.bn} item belum selesai${s.bn?", tertua "+s.bk+" hari":""}.</div>${H("bs","Belum selesai ("+x.length+")")}<div class="sb" ${COL.bs?"hidden":""}>`+
   (x.length?table(["Modul","Menunggu","Status","Tgl input","No dokumen","No PM","Tgl PM","No tag","Barang","Aging"],x.map(r=>[r.mod,who(r),r.st,fd(r.tgl),norm(r.nttb),pmn(r),pmt(r),r.tag,r.barang,badge(r.age)])):'<div class="sub ok">Tidak ada item belum selesai.</div>')+"</div>"+
   `${H("hs","Histori selesai ("+s.f.length+(s.f.length>100?", tampil 100 terbaru":"")+")")}<div class="sb" ${COL.hs?"hidden":""}>`+(f.length?table(["Modul","Tgl input","Tgl selesai","No dokumen","No PM","Tgl PM","No tag","Barang"],f.slice(0,100).map(r=>[r.mod,fd(r.tgl),fd(r.tsel),norm(r.nttb),pmn(r),pmt(r),r.tag,r.barang])):'<div class="sub">Belum ada histori selesai pada periode ini.</div>')+"</div></div>";
   if($("#sb"))$("#sb").onclick=()=>{shareSel=k;go("sh")};$("#xs").onclick=()=>xl(new Set([k]));$("#cl").onclick=()=>{selK="";$("#det").innerHTML=""};$("#det").querySelectorAll(".tg").forEach(g=>g.onclick=()=>{const b=g.querySelector(".tgb"),id=b.dataset.id;COL[id]=!COL[id];b.textContent=COL[id]?"+":"\u2212";g.nextElementSibling.hidden=COL[id]})};
  const draw=()=>{F={q:$("#q").value,p:$("#fp").value,a:$("#fa").value,c:$("#fc").value,j:$("#fj").value};const q=lc(F.q).replace(/\s+-\s+/g," - ");
   lastR=S.filter(s=>(!F.p||s.pr===F.p)&&(!F.a||s.am===F.a)&&(!F.c||s.ac===F.c)&&(!F.j||(F.j==="a"?s.a:s.m))&&(!q||(lc(s.k+" - "+s.nama+" "+s.am+" "+s.ac+" "+s.pj)+" "+s.x.map(r=>lc(r.tag+" "+r.barang)).join(" ")).includes(q)));
   $("#res").innerHTML=`<div class="bar" style="margin:8px 0"><div class="sub" style="flex:1;margin:0">${lastR.length} dari ${S.length} toko</div><button id="xa" class="alt">Unduh Excel (belum selesai, ${lastR.length} toko)</button></div>`+table(["Prioritas","Kode","Toko","AC","AM","Pejabat Toko","Aktivasi belum","Mutasi belum","Aging tertua · "+lb,"Backlog semua bulan","Selesai",""],
    lastR.map(s=>[`<span class="bd ${s.pr==="KRITIS"?"K":s.pr==="TINGGI"||s.pr==="BACKLOG"?"T":""}">${s.pr}</span>`,"<b>"+s.k+"</b>",s.nama,s.ac,s.am,s.pj,s.a,s.m,s.age<0?"-":badge(s.age),s.bk==null?"-":badge(s.bk),s.fa+s.fm,`<button class="alt" data-k="${s.k}">Lihat</button>`]),[6,7,10]);
   $("#res").querySelectorAll("button[data-k]").forEach(b=>b.onclick=()=>{det(b.dataset.k);$("#det").scrollIntoView&&$("#det").scrollIntoView({behavior:"smooth"})});$("#xa").onclick=()=>xl(new Set(lastR.map(s=>s.k)));
  };
  ["fp","fa","fc","fj"].forEach(i=>$("#"+i).onchange=draw);let hi=-1;const mt=q=>S.filter(s=>lc(s.k+" - "+s.nama+" "+s.am+" "+s.ac+" "+s.pj).includes(q));
  const nm=s=>s.k+" - "+s.nama,hide=()=>{$("#sg").innerHTML="";hi=-1};
  const jump=()=>{$("#det").scrollIntoView&&$("#det").scrollIntoView({behavior:"smooth",block:"nearest"})};
  const pick=k=>{const s=S.find(z=>z.k===k);if(!s)return;$("#q").value=nm(s);hide();draw();det(k);jump()};
  const sug=()=>{const q=lc($("#q").value).replace(/\s+-\s+/g," - ");if(!q){hide();return}const L=mt(q).slice(0,8);
   $("#sg").innerHTML=L.map((s,i)=>`<div data-k="${s.k}" class="${i===hi?"on":""}"><b>${s.k}</b> · ${s.nama}<br><span class="sm">AC: ${s.ac} · AM: ${s.am} · ${s.x.length?s.x.length+" belum selesai":s.bn?"bulan ini bersih, backlog "+s.bn+" item":"bersih"}</span></div>`).join("")||'<div class="sm">Toko tidak ditemukan</div>';
   $("#sg").querySelectorAll("div[data-k]").forEach(d=>d.onmousedown=e=>{e.preventDefault();pick(d.dataset.k)});
   const ex=S.find(s=>lc(s.k)===q||lc(s.nama)===q||lc(nm(s))===q)||(L.length===1&&q.length>=3?L[0]:null);
   if(ex&&ex.k!==selK){det(ex.k);jump()}};
  $("#q").oninput=()=>{hi=-1;draw();sug()};
  $("#q").onkeydown=e=>{const q=lc($("#q").value).replace(/\s+-\s+/g," - "),L=q?mt(q).slice(0,8):[];
   if(e.key==="ArrowDown"||e.key==="ArrowUp"){e.preventDefault();hi=L.length?(hi+(e.key==="ArrowDown"?1:-1)+L.length)%L.length:-1;sug()}
   else if(e.key==="Enter"){e.preventDefault();const m=L[hi>=0?hi:0]||lastR[0];if(m)pick(m.k)}else if(e.key==="Escape")hide()};
  document.onclick=e=>{if($("#sg")&&e.target&&e.target.closest&&!e.target.closest(".qw"))hide()};
  draw();if(selK&&S.some(s=>s.k===selK))det(selK);
 }else if(tab==="hs"){
  const mkg=L=>{const g={},o=m=>g[m]=g[m]||{t:0,d:0,b:0,k:0,sel:0};L.forEach(r=>{if(r.bln){const z=o(r.bln);z.t++;r.s==="done"?z.d++:r.tolak?z.k++:z.b++}if(r.bs)o(r.bs).sel++});return g};
  const tb=(L,n)=>{const g=mkg(L);return `<h2>${n}</h2>`+table(["Bulan","Diinput","Selesai","Belum selesai","Tolak/Batal","% selesai (dari input)","Diselesaikan di bulan ini",""],
   Object.keys(g).sort().reverse().map(m=>{const o=g[m],pc=o.t?o.d*100/o.t:0;return["<b>"+ML(m)+"</b>",fmt(o.t),fmt(o.d),fmt(o.b),fmt(o.k),`<span class="bd ${pc<90?"T":""}">${pc.toFixed(1)}%</span>`,fmt(o.sel),`<button class="alt" data-m="${m}">Pilih bulan</button>`]}),[1,2,3,4,5,6])};
  p.innerHTML=`<div class="sub">Histori semua bulan. "% selesai" = bagian yang Done dari pengajuan yang diinput pada bulan itu. "Diselesaikan di bulan ini" menghitung berdasarkan tanggal terima (aktivasi) atau tanggal PM (mutasi). Klik Pilih bulan untuk membuka bulan tersebut di Cari Toko.</div>`+tb(AA,"Aktivasi")+"<br>"+tb(MM,"Mutasi");
  p.querySelectorAll("button[data-m]").forEach(b=>b.onclick=()=>{$("#bln").value=b.dataset.m;go("tk")});
 }else if(tab==="sh"){
  const pass=(r,m)=>m==="all"||(m==="toko"&&r.aksi)||(m==="lain"&&!r.aksi)||m==="s:"+r.s;
  const by={};pa.concat(pm).forEach(r=>{if(r.tp&&pass(r,shMode))(by[r.tp]=by[r.tp]||[]).push(r)});
  const mx=k=>Math.max(0,...by[k].map(r=>r.age||0)),keys=Object.keys(by).sort((a,b)=>mx(b)-mx(a));
  const sts={};pa.concat(pm).forEach(r=>{const o=sts[r.s]=sts[r.s]||{n:r.st,c:0};o.c++});
  const modeOpt='<option value="all">Semua status</option><option value="toko">Hanya menunggu toko (Pejabat Toko, Konfirmasi, User Penerima, Atasan Penyerah)</option><option value="lain">Hanya menunggu pihak lain (AM, Checker FA, GSM, Accounting)</option>'+Object.entries(sts).sort((a,b)=>b[1].c-a[1].c).map(([s,o])=>`<option value="s:${s}">Hanya status ${o.n} (${o.c})</option>`).join("");
  const setMode=()=>{$("#sm").value=shMode;$("#sm").onchange=()=>{shMode=$("#sm").value;render()}};
  if(!keys.length){p.innerHTML=`<h2>Share ke Toko</h2><div class="bar"><select id="sm">${modeOpt}</select></div><div class="sub">Tidak ada item untuk pilihan ini pada periode ${lb}.</div>`;setMode();return}
  if(!keys.includes(shareSel))shareSel=keys[0];
  const msg=k=>{const x=by[k],s=info[k]||{},g={};x.forEach(r=>(g[r.s]=g[r.s]||[]).push(r));
   const gs=Object.values(g).sort((a,b)=>(b[0].aksi-a[0].aksi)||(Math.max(...b.map(r=>r.age||0))-Math.max(...a.map(r=>r.age||0))));
   let t=`*FOLLOW-UP APPROVAL ASSET ONLINE*\nToko: ${k} - ${s.nama||"Branch Palu"}\nAC: ${s.ac||"-"}\nPejabat Toko: ${s.pj||"-"}\nAM: ${s.am||"-"}\nPer ${$("#ref").value}\n\nAda *${x.length} item* belum selesai, tertua *${mx(k)} hari*:`;
   gs.forEach(y=>{const tl=D.par[y[0].s]||DEF[y[0].s]||"Cek status di Asset Online";t+=`\n\n*${y[0].st} (${y.length})* - ${tl}`;
    ["Aktivasi","Mutasi"].forEach(m=>{const z=y.filter(r=>r.mod===m).sort((a,b)=>(b.age||0)-(a.age||0));if(!z.length)return;t+=`\n_${m}_`;z.slice(0,10).forEach((r,i)=>t+=`\n${i+1}. ${r.tag} | ${norm(r.barang).slice(0,45)} | ${r.age} hari`);if(z.length>10)t+=`\n...dan ${z.length-10} item lainnya`})});
   return t+"\n\n"+(x.some(r=>r.aksi)?"Mohon item yang menunggu toko segera diselesaikan di Asset Online (menu approval sesuai status). ":"")+(x.some(r=>!r.aksi)?"Item yang menunggu AM/FA/GSM/Accounting sedang kami follow-up; mohon dibantu dikawal bila ada kendala. ":"")+"Jika sudah diproses, mohon info ke Staff FA Branch Palu. Terima kasih."};
  p.innerHTML=`<h2>Share ke Toko (${keys.length} toko)</h2><div class="sub">Pilih status yang ingin dishare. Pesan dikelompokkan per status beserta tindak lanjutnya (termasuk proxy Pejabat Toko).</div><div class="bar"><select id="sm">${modeOpt}</select><select id="ss">${keys.map(k=>`<option value="${k}">${k} - ${(info[k]||{}).nama||"Branch Palu"} (${by[k].length} item, ${mx(k)} hari)</option>`).join("")}</select></div><br><textarea id="mt" rows="18" style="width:100%;font:12.5px monospace;background:var(--c);color:var(--t);border:1px solid var(--b);border-radius:6px;padding:8px"></textarea>
  <div class="bar" style="margin-top:8px"><button id="cp">Salin pesan</button><button id="wa" class="alt">Buka WhatsApp</button><button id="all" class="alt">Unduh semua toko (.txt)</button><span id="ok" class="ok"></span></div>`;
  setMode();$("#ss").value=shareSel;const sh=()=>{shareSel=$("#ss").value;$("#mt").value=msg(shareSel)};$("#ss").onchange=sh;sh();
  $("#cp").onclick=()=>{const v=$("#mt").value;$("#mt").select();(navigator.clipboard?navigator.clipboard.writeText(v):Promise.reject()).catch(()=>document.execCommand("copy")).then(()=>{$("#ok").textContent="Tersalin"})};
  $("#wa").onclick=()=>window.open("https://wa.me/?text="+encodeURIComponent($("#mt").value),"_blank");
  $("#all").onclick=()=>{saveFile("pesan_wa_toko.txt",keys.map(msg).join("\n\n"+"=".repeat(40)+"\n\n"))};
 }else if(tab==="lk"){
  p.innerHTML=`<h2>Link untuk anak toko (hanya lihat)</h2><div class="sub">Data aktivasi, mutasi, toko, AC, AM dan Pejabat Toko dienkripsi dengan kata sandi. Anak toko membuka link, memasukkan sandi, lalu bisa mencari toko dan mengunduh Excel. Tidak ada tombol untuk mengubah data.</div>
  <div class="bar"><input type="password" id="p1" placeholder="Kata sandi (min. 10 karakter)" autocomplete="new-password"><input type="password" id="p2" placeholder="Ulangi kata sandi" autocomplete="new-password"><button id="mk">Buat data.js</button></div><div id="mm" class="sub"></div>
  <ol class="sub"><li>Klik <b>Buat data.js</b>, lalu simpan filenya.</li><li>Unggah <code>data.js</code> ke repositori GitHub (timpa file lama). Link yang sama langsung menampilkan data terbaru.</li><li>Bagikan link dan kata sandi lewat jalur berbeda (mis. link di grup, sandi lewat pesan pribadi).</li><li>Halaman admin: tambahkan <code>?admin</code> di akhir link.</li></ol>`;
  $("#mk").onclick=async()=>{const a=$("#p1").value;if(a.length<10){$("#mm").textContent="Kata sandi minimal 10 karakter.";return}if(a!==$("#p2").value){$("#mm").textContent="Pengulangan sandi tidak sama.";return}
   $("#mm").textContent="Mengenkripsi (beberapa detik)...";const v=$("#ref").value,lab=+v.slice(8,10)+" "+ML(v.slice(0,7)).replace(" "," ");
   try{const o=await enc({akt:D.akt.map(({_row,...r})=>r),mut:D.mut.map(({_row,...r})=>r),str:D.str,par:D.par,rf:v,kr:+$("#kr").value,ti:+$("#ti").value},a);
    await saveFile("data.js","window.GIT_DATA="+JSON.stringify({...o,label:lab})+";");$("#mm").textContent="Selesai: data.js dibuat (data per "+lab+")."}catch(e){$("#mm").textContent="Gagal: "+e.message}}
 }else{
  const U=v=>norm(v).toUpperCase();
  const bad=(L)=>L.filter(r=>r.mod==="Aktivasi"?(U(r.toko)&&U(r.toko)!=="SOZ1"&&!kode.has(U(r.toko))):!(U(r.asal)==="SOZ1"&&U(r.toko)==="SOZ1")&&!kode.has(U(r.asal))&&!kode.has(U(r.toko)));
  const noDate=A.concat(M).filter(r=>r.age==null).length;
  const rows=[["Aktivasi: kode toko tidak ada di struktur",fmt(bad(A).length)],["Mutasi: kode toko tidak ada di struktur",fmt(bad(M).length)],["Baris dengan tanggal tidak terbaca",fmt(noDate)],["Baris tanpa nomor tag",fmt(A.concat(M).filter(r=>!r.dkey).length)],["Status tidak dikenal",unk.join(", ")||"-"]];
  p.innerHTML="<h2>Cek kualitas data</h2>"+table(["Pemeriksaan","Hasil"],rows,[1]);
 }
}
$("#f").onchange=e=>{const f=e.target.files[0];if(!f)return;$("#info").textContent="Membaca file...";
 const rd=new FileReader();rd.onload=()=>{try{load(rd.result)}catch(err){$("#info").textContent="Gagal: "+err.message}};rd.readAsArrayBuffer(f)};
["ref","kr","ti","bln"].forEach(i=>$("#"+i).onchange=render);
const b64=u=>{let s="";for(let i=0;i<u.length;i+=8192)s+=String.fromCharCode.apply(null,u.subarray(i,i+8192));return btoa(s)},ub64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
const gz=async b=>new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(new CompressionStream("gzip"))).arrayBuffer()),gunz=async b=>new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(new DecompressionStream("gzip"))).arrayBuffer());
async function dkey(pw,salt){const m=await crypto.subtle.importKey("raw",new TextEncoder().encode(pw),"PBKDF2",false,["deriveKey"]);return crypto.subtle.deriveKey({name:"PBKDF2",salt,iterations:310000,hash:"SHA-256"},m,{name:"AES-GCM",length:256},false,["encrypt","decrypt"])}
async function enc(obj,pw){const salt=crypto.getRandomValues(new Uint8Array(16)),iv=crypto.getRandomValues(new Uint8Array(12)),k=await dkey(pw,salt);
 const ct=new Uint8Array(await crypto.subtle.encrypt({name:"AES-GCM",iv},k,await gz(new TextEncoder().encode(JSON.stringify(obj)))));return{v:1,s:b64(salt),i:b64(iv),c:b64(ct)}}
async function dec(p,pw){const k=await dkey(pw,ub64(p.s));return JSON.parse(new TextDecoder().decode(await gunz(new Uint8Array(await crypto.subtle.decrypt({name:"AES-GCM",iv:ub64(p.i)},k,ub64(p.c))))))}
function xl(stores){const X=analyse(),mf=$("#bln").value,lb=mf?ML(mf):"Semua bulan",sel=r=>r.pend&&(!mf||r.bln===mf)&&(!stores||stores.has(r.tp)),srt=(a,b)=>(b.age||0)-(a.age||0);
 const bs=r=>{const s=X.info[r.tp]||{};return[r.tp,s.nama||"Branch Palu",s.ac||"",s.am||"",s.pj||""]},wh=r=>WH[r.s]||"Cek Asset Online";
 const A=X.A.filter(sel).sort(srt).map(r=>[...bs(r),fd(r.tgl),norm(r.nttb),r.st,wh(r),norm(r.tag),norm(r.barang),r.age,r.bln?ML(r.bln):""]);
 const M=X.M.filter(sel).sort(srt).map(r=>[...bs(r),fd(r.tgl),norm(r.nttb),norm(r.asal),norm(r.toko),r.st,wh(r),norm(r.nopm),r.nopm?fd(r.tsel):"",norm(r.tag),norm(r.barang),r.age,r.bln?ML(r.bln):""]);
 const h0=["KODE TOKO","NAMA TOKO","AC","AM","PEJABAT TOKO"],wb=XLSX.utils.book_new();
 const add=(n,hd,rows,w)=>{const ws=XLSX.utils.aoa_to_sheet([hd,...rows]);ws["!cols"]=w.map(x=>({wch:x}));ws["!autofilter"]={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(rows.length,1),c:hd.length-1}})};XLSX.utils.book_append_sheet(wb,ws,n)};
 add("Aktivasi",[...h0,"TGL NTTB","NO NTTB","STATUS","MENUNGGU","NO TAG","NAMA BARANG","AGING (HARI)","BULAN INPUT"],A,[10,28,20,22,22,12,26,22,12,22,46,10,12]);
 add("Mutasi",[...h0,"TGL PENGAJUAN","NO PENGAJUAN","UU AWAL","UU TUJUAN","STATUS","MENUNGGU","NO PM","TGL PM","NO TAG","NAMA BARANG","AGING (HARI)","BULAN INPUT"],M,[10,28,20,22,22,14,26,9,10,22,12,24,12,22,46,10,12]);
 const nm=stores&&stores.size===1?[...stores][0]:"semua-toko";saveFile("GIT_Belum_Selesai_"+nm+"_"+lb.replace(" ","-")+".xlsx",new Uint8Array(XLSX.write(wb,{bookType:"xlsx",type:"array"})))}
function enterViewer(){VIEW=true;document.body.classList.add("view");$("#login").style.display="";const e=document.querySelector("header .hd>div>div");if(e)e.textContent="Mode lihat saja · masukkan kata sandi";
 $("#pwi").textContent="Data per "+GIT_DATA.label+".";
 const go1=async()=>{const pw=$("#pw").value;if(!pw)return;$("#pwm").textContent="Membuka data...";try{const d=await dec(GIT_DATA,pw);$("#login").style.display="none";start(d)}catch(err){$("#pwm").textContent="Kata sandi salah atau data tidak bisa dibuka."}};
 $("#pwb").onclick=go1;$("#pw").onkeydown=e=>{if(e.key==="Enter")go1()}}
function boot(){if(/[?&]admin/.test(location.search))return;const s=document.createElement("script");s.src="data.js?v="+Date.now();s.onload=()=>{if(window.GIT_DATA)enterViewer()};s.onerror=()=>{};document.head.appendChild(s)}
async function saveFile(name,text){const bytes=typeof text==="string"?new TextEncoder().encode(text):text;
 try{if(window.claude&&claude.use){const d=await claude.use("downloads");if(d){await d.save({filename:name,data:bytes});return}}}catch(e){if(e&&e.code==="declined")return}
 const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([bytes]));a.download=name;a.click()}
const SORT={},RK={KRITIS:0,TINGGI:1,BACKLOG:2,NORMAL:3,BERSIH:4};
function sv(t){t=t.trim();if(t===""||t==="-")return null;const u=t.toUpperCase();if(u in RK)return RK[u];
 const m=t.match(/^(-?[\d.]+)\s*(hari|%)?$/);if(m){const s=m[1];return m[2]==="%"||!/^\d{1,3}(\.\d{3})+$/.test(s)?parseFloat(s):parseFloat(s.replace(/\./g,""))}return t.toLowerCase()}
function cmp(a,b,d){if(a===null&&b===null)return 0;if(a===null)return 1;if(b===null)return -1;
 return typeof a==="number"&&typeof b==="number"?(a-b)*d:String(a).localeCompare(String(b),"id",{numeric:true})*d}
function sortTable(t,i,d){const rows=[...t.tBodies[0].rows],v=new Map(rows.map(r=>[r,sv(r.cells[i].textContent)]));
 rows.sort((x,y)=>cmp(v.get(x),v.get(y),d));rows.forEach(r=>t.tBodies[0].appendChild(r));t.querySelectorAll("th").forEach((h,j)=>h.dataset.d=j===i?d:"")}
function initSort(){document.querySelectorAll(".w table:not([data-s])").forEach(t=>{t.dataset.s=1;const hs=[...t.querySelectorAll("th")],key=hs.map(h=>h.textContent).join("|"),w=t.parentNode;
  const flip=i=>{const c=SORT[key],d=c&&c.i===i?-c.d:1;SORT[key]={i,d};sortTable(t,i,d)};
  hs.forEach((h,i)=>{if(h.textContent.trim())h.onclick=()=>flip(i)});
  const bar=document.createElement("div");bar.className="srt";bar.innerHTML=`<select>${hs.map((h,i)=>h.textContent.trim()?`<option value="${i}">Urutkan: ${h.textContent}</option>`:"").join("")}</select><button class="alt" title="Naik / turun">&#8645;</button>`;
  w.parentNode.insertBefore(bar,w);const sel=bar.querySelector("select");
  sel.onchange=()=>{const i=+sel.value;SORT[key]={i,d:1};sortTable(t,i,1)};bar.querySelector("button").onclick=()=>flip(+sel.value);
  if(SORT[key]){sel.value=SORT[key].i;sortTable(t,SORT[key].i,SORT[key].d)}})}
if(typeof MutationObserver!=="undefined")new MutationObserver(()=>requestAnimationFrame(initSort)).observe(document.body,{childList:true,subtree:true});
function go(t){tab=t;document.querySelectorAll(".tabs button").forEach(x=>x.classList.toggle("alt",x.dataset.t!==t));render()}
document.querySelectorAll(".tabs button").forEach(b=>b.onclick=()=>go(b.dataset.t));

if("serviceWorker" in navigator&&location.protocol.startsWith("http"))window.addEventListener("load",()=>navigator.serviceWorker.register("sw.js").catch(()=>{}));
boot();
