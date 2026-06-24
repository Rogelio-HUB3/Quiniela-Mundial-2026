<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
  <title>Quiniela Mundial 2026 · SAGA</title>
  <!-- Microsoft Auth Library -->
  <script src="https://alcdn.msauth.net/browser/3.14.0/js/msal-browser.min.js"></script>
  <!-- React + Babel -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/react/18.3.1/umd/react.production.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.3.1/umd/react-dom.production.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.24.7/babel.min.js"></script>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',system-ui,sans-serif;background:#030a03;color:#fff;min-height:100vh}
    input[type=number]::-webkit-inner-spin-button{opacity:1}
    .hs{scrollbar-width:none}.hs::-webkit-scrollbar{display:none}
    @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
    .pulse{animation:pulse 1.5s infinite}
    @keyframes spin{to{transform:rotate(360deg)}}
    .spin{animation:spin .8s linear infinite;display:inline-block}
    select option{background:#111;color:#fff}
  </style>
</head>
<body>
<div id="root"></div>
<script type="text/babel" data-presets="react">
// ═══════════════════════════════════════════════════════════════════════════
//  ▼▼▼  CONFIGURACIÓN — EL ADMIN LLENA ESTO UNA VEZ Y COMPARTE EL ARCHIVO  ▼▼▼
// ═══════════════════════════════════════════════════════════════════════════
const CONFIG = {
  clientId:  "PASTE_CLIENT_ID_HERE",       // ID de aplicación de Azure AD
  tenantId:  "PASTE_TENANT_ID_HERE",       // ID del directorio (tenant)
  siteUrl:   "https://EMPRESA.sharepoint.com/sites/NOMBRE_SITIO",
  listPartidos:  "QuinielaPartidos",       // nombre de la lista de partidos
  listQuinielas: "QuinielaQuinielas",      // nombre de la lista de quinielas
  adminPass: "saga2026",                   // contraseña del panel admin
};
// ═══════════════════════════════════════════════════════════════════════════

const IS_CONFIGURED = !CONFIG.clientId.includes("PASTE") && !CONFIG.tenantId.includes("PASTE") && !CONFIG.siteUrl.includes("EMPRESA");

// ─── Fases / Scoring ──────────────────────────────────────────────────────
const PHASES=[
  {id:'groups_j3',label:'Jornada 3',sub:'Fase de Grupos',ko:false},
  {id:'round_of_32',label:'Ronda de 32',sub:'Eliminatorias',ko:true},
  {id:'round_of_16',label:'Octavos',sub:'Eliminatorias',ko:true},
  {id:'quarterfinal',label:'Cuartos',sub:'Eliminatorias',ko:true},
  {id:'semifinal',label:'Semifinal',sub:'Eliminatorias',ko:true},
  {id:'third_place',label:'3er Lugar',sub:'Eliminatorias',ko:true},
  {id:'final',label:'Gran Final',sub:'Eliminatorias',ko:true},
];
const PM=Object.fromEntries(PHASES.map(p=>[p.id,p]));

// ─── Seed — Claude actualiza SEED_VERSION + results aquí ─────────────────
const SEED_VERSION=2;
const SEED_MATCHES=[
  {id:'j3_b1',phase:'groups_j3',group:'B',team1:'Suiza',team2:'Canadá',matchDatetime:'2026-06-24T14:00',result:null},
  {id:'j3_b2',phase:'groups_j3',group:'B',team1:'Bosnia',team2:'Qatar',matchDatetime:'2026-06-24T14:00',result:null},
  {id:'j3_c1',phase:'groups_j3',group:'C',team1:'Escocia',team2:'Brasil',matchDatetime:'2026-06-24T17:00',result:null},
  {id:'j3_c2',phase:'groups_j3',group:'C',team1:'Marruecos',team2:'Haití',matchDatetime:'2026-06-24T17:00',result:null},
  {id:'j3_a1',phase:'groups_j3',group:'A',team1:'Chequia',team2:'México',matchDatetime:'2026-06-24T21:00',result:null},
  {id:'j3_a2',phase:'groups_j3',group:'A',team1:'Sudáfrica',team2:'Corea del Sur',matchDatetime:'2026-06-24T21:00',result:null},
  {id:'j3_e1',phase:'groups_j3',group:'E',team1:'Ecuador',team2:'Alemania',matchDatetime:'2026-06-25T15:00',result:null},
  {id:'j3_e2',phase:'groups_j3',group:'E',team1:'Curazao',team2:'C. de Marfil',matchDatetime:'2026-06-25T15:00',result:null},
  {id:'j3_f1',phase:'groups_j3',group:'F',team1:'Japón',team2:'Suecia',matchDatetime:'2026-06-25T18:00',result:null},
  {id:'j3_f2',phase:'groups_j3',group:'F',team1:'Túnez',team2:'Países Bajos',matchDatetime:'2026-06-25T18:00',result:null},
  {id:'j3_d1',phase:'groups_j3',group:'D',team1:'Turquía',team2:'EE.UU.',matchDatetime:'2026-06-25T21:00',result:null},
  {id:'j3_d2',phase:'groups_j3',group:'D',team1:'Paraguay',team2:'Australia',matchDatetime:'2026-06-25T21:00',result:null},
  {id:'j3_i1',phase:'groups_j3',group:'I',team1:'Noruega',team2:'Francia',matchDatetime:'2026-06-26T14:00',result:null},
  {id:'j3_i2',phase:'groups_j3',group:'I',team1:'Senegal',team2:'Irak',matchDatetime:'2026-06-26T14:00',result:null},
  {id:'j3_h1',phase:'groups_j3',group:'H',team1:'Cabo Verde',team2:'Arabia Saudita',matchDatetime:'2026-06-26T19:00',result:null},
  {id:'j3_h2',phase:'groups_j3',group:'H',team1:'Uruguay',team2:'España',matchDatetime:'2026-06-26T19:00',result:null},
  {id:'j3_g1',phase:'groups_j3',group:'G',team1:'Egipto',team2:'Irán',matchDatetime:'2026-06-26T22:00',result:null},
  {id:'j3_g2',phase:'groups_j3',group:'G',team1:'Nueva Zelanda',team2:'Bélgica',matchDatetime:'2026-06-26T22:00',result:null},
  {id:'j3_l1',phase:'groups_j3',group:'L',team1:'Panamá',team2:'Inglaterra',matchDatetime:'2026-06-27T16:00',result:null},
  {id:'j3_l2',phase:'groups_j3',group:'L',team1:'Croacia',team2:'Ghana',matchDatetime:'2026-06-27T16:00',result:null},
  {id:'j3_k1',phase:'groups_j3',group:'K',team1:'Colombia',team2:'Portugal',matchDatetime:'2026-06-27T18:30',result:null},
  {id:'j3_k2',phase:'groups_j3',group:'K',team1:'DR Congo',team2:'Uzbekistán',matchDatetime:'2026-06-27T18:30',result:null},
  {id:'j3_j1',phase:'groups_j3',group:'J',team1:'Argelia',team2:'Austria',matchDatetime:'2026-06-27T21:00',result:null},
  {id:'j3_j2',phase:'groups_j3',group:'J',team1:'Jordania',team2:'Argentina',matchDatetime:'2026-06-27T21:00',result:null},
];

// ─── Utilerías ────────────────────────────────────────────────────────────
const {useState,useEffect,useCallback,useRef}=React;
function calcPts(pred,res,ko){
  if(!pred||!res||res.s1==null||res.s2==null) return null;
  if(+pred.s1===+res.s1&&+pred.s2===+res.s2) return 5;
  const pw=ko?pred.winner:(+pred.s1>+pred.s2?'h':+pred.s2>+pred.s1?'a':'d');
  const rw=ko?res.winner:(+res.s1>+res.s2?'h':+res.s2>+res.s1?'a':'d');
  if(pw===rw){if(ko&&pred.et!=null&&pred.et===res.et) return 4;return 3;}
  return 0;
}
function getStatus(m,now){
  if(m.result) return 'played';if(!m.matchDatetime) return 'open';
  const s=new Date(m.matchDatetime),l=new Date(s-3600000);
  if(now<l) return 'open';if(now<s) return 'locked';return 'live';
}
function lockCD(m,now){
  if(!m.matchDatetime) return null;
  const l=new Date(new Date(m.matchDatetime)-3600000),d=l-now;
  if(d<=0) return null;
  const h=Math.floor(d/3600000),mi=Math.floor((d%3600000)/60000);
  if(d<3600000) return{label:`Cierra en ${mi}m`,urgent:true};
  if(d<86400000) return{label:`Cierra en ${h}h ${String(mi).padStart(2,'0')}m`,urgent:false};
  return null;
}
function fmtDt(dt){
  if(!dt) return null;
  return new Date(dt).toLocaleString('es-MX',{weekday:'short',day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'});
}

// ─── MSAL ─────────────────────────────────────────────────────────────────
let msalApp=null;
function getMSAL(){
  if(!msalApp&&IS_CONFIGURED){
    msalApp=new msal.PublicClientApplication({
      auth:{
        clientId:CONFIG.clientId,
        authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri:window.location.href.replace(/[?#].*/,''),
      },
      cache:{cacheLocation:'sessionStorage',storeAuthStateInCookie:false},
    });
  }
  return msalApp;
}
const SCOPES=["User.Read","Sites.ReadWrite.All","Sites.Manage.All"];

async function getToken(){
  const app=getMSAL();const accounts=app.getAllAccounts();
  if(!accounts.length) throw new Error('NO_ACCOUNT');
  try{
    const r=await app.acquireTokenSilent({scopes:SCOPES,account:accounts[0]});
    return r.accessToken;
  }catch{
    const r=await app.acquireTokenPopup({scopes:SCOPES});
    return r.accessToken;
  }
}
async function signIn(){
  const app=getMSAL();
  const r=await app.loginPopup({scopes:SCOPES});
  return r;
}
function getAccount(){
  const app=getMSAL();if(!app) return null;
  const acc=app.getAllAccounts();return acc.length?acc[0]:null;
}
async function signOut(){
  const app=getMSAL();
  await app.logoutPopup({account:getAccount()});
}

// ─── Graph / SharePoint helpers ───────────────────────────────────────────
async function gFetch(path,method='GET',body=null){
  const token=await getToken();
  const r=await fetch(`https://graph.microsoft.com/v1.0${path}`,{
    method,
    headers:{Authorization:`Bearer ${token}`,...(body?{'Content-Type':'application/json'}:{})},
    body:body?JSON.stringify(body):null,
  });
  if(r.status===204) return null;
  const data=await r.json();
  if(!r.ok) throw new Error(data?.error?.message||`Graph ${r.status}: ${path}`);
  return data;
}

// Resolve site ID from siteUrl
let _siteId=null,_partListId=null,_quinListId=null;
async function getSiteId(){
  if(_siteId) return _siteId;
  const url=new URL(CONFIG.siteUrl);
  const host=url.hostname,path=url.pathname;
  const site=await gFetch(`/sites/${host}:${path}`);
  _siteId=site.id;return _siteId;
}
async function getListId(listName){
  const sid=await getSiteId();
  const lists=await gFetch(`/sites/${sid}/lists?$select=id,displayName&$filter=displayName eq '${encodeURIComponent(listName)}'`);
  if(lists.value.length===0) return null;
  return lists.value[0].id;
}

// ─── SharePoint List Setup (first run) ───────────────────────────────────
async function ensureLists(){
  const sid=await getSiteId();
  // Create QuinielaPartidos if missing
  let pid=await getListId(CONFIG.listPartidos);
  if(!pid){
    const list=await gFetch(`/sites/${sid}/lists`,'POST',{
      displayName:CONFIG.listPartidos,
      list:{template:'genericList'},
      columns:[
        {name:'MatchId',text:{}},
        {name:'DataJSON',text:{maxLength:2000}},
        {name:'ResultJSON',text:{maxLength:500}},
        {name:'SeedVersion',number:{}},
      ],
    });
    pid=list.id;
  }
  _partListId=pid;
  // Create QuinielaQuinielas if missing
  let qid=await getListId(CONFIG.listQuinielas);
  if(!qid){
    const list=await gFetch(`/sites/${sid}/lists`,'POST',{
      displayName:CONFIG.listQuinielas,
      list:{template:'genericList'},
      columns:[
        {name:'MatchId',text:{}},
        {name:'UserEmail',text:{}},
        {name:'UserName',text:{}},
        {name:'PredJSON',text:{maxLength:500}},
      ],
    });
    qid=list.id;
  }
  _quinListId=qid;
  return{partidosId:pid,quinielasId:qid};
}

// ─── Data Layer ───────────────────────────────────────────────────────────
async function spGet(listId,filter='',select=''){
  const sid=await getSiteId();
  let url=`/sites/${sid}/lists/${listId}/items?$expand=fields&$top=500`;
  if(filter) url+=`&$filter=${filter}`;
  if(select) url+=`&$select=${select}`;
  const data=await gFetch(url);
  return data.value;
}
async function spPost(listId,fields){
  const sid=await getSiteId();
  return gFetch(`/sites/${sid}/lists/${listId}/items`,'POST',{fields});
}
async function spPatch(listId,itemId,fields){
  const sid=await getSiteId();
  return gFetch(`/sites/${sid}/lists/${listId}/items/${itemId}`,'PATCH',{fields});
}

async function ensureListIds(){
  if(!_partListId) _partListId=await getListId(CONFIG.listPartidos);
  if(!_quinListId) _quinListId=await getListId(CONFIG.listQuinielas);
  if(!_partListId||!_quinListId) throw new Error('LISTS_NOT_FOUND');
}

async function loadMatches(){
  await ensureListIds();
  const items=await spGet(_partListId);
  return items.map(i=>({
    id:i.fields.MatchId||i.fields.Title,
    spId:i.id,
    ...JSON.parse(i.fields.DataJSON||'{}'),
    result:i.fields.ResultJSON?JSON.parse(i.fields.ResultJSON):null,
  })).sort((a,b)=>(a.matchDatetime||'')<(b.matchDatetime||'')?-1:1);
}

async function loadAllPredictions(){
  await ensureListIds();
  const items=await spGet(_quinListId);
  const preds={};
  items.forEach(i=>{
    const{MatchId,UserName,PredJSON}=i.fields;
    if(!MatchId||!UserName) return;
    if(!preds[MatchId]) preds[MatchId]={};
    try{preds[MatchId][UserName]={...JSON.parse(PredJSON),spId:i.id};}catch{}
  });
  return preds;
}

async function savePrediction(matchId,pred,userEmail,userName){
  await ensureListIds();
  const titleKey=`${matchId}|${userEmail}`;
  // Check if already exists
  const existing=await spGet(_quinListId,`fields/Title eq '${titleKey}'`);
  const fields={Title:titleKey,MatchId:matchId,UserEmail:userEmail,UserName:userName,PredJSON:JSON.stringify(pred)};
  if(existing.length>0) await spPatch(_quinListId,existing[0].id,fields);
  else await spPost(_quinListId,fields);
}

async function seedMatchesToSP(matches){
  await ensureListIds();
  const existing=await spGet(_partListId);
  const existingById=Object.fromEntries(existing.map(i=>[i.fields.MatchId||i.fields.Title,i]));
  for(const m of matches){
    const{result,...data}=m;
    const ex=existingById[m.id];
    if(!ex){
      await spPost(_partListId,{Title:m.id,MatchId:m.id,DataJSON:JSON.stringify(data),ResultJSON:result?JSON.stringify(result):'',SeedVersion:SEED_VERSION});
    }else if(!ex.fields.ResultJSON&&result){
      await spPatch(_partListId,ex.id,{ResultJSON:JSON.stringify(result),SeedVersion:SEED_VERSION});
    }
  }
}

async function updateMatchResult(matchSpId,result){
  await ensureListIds();
  await spPatch(_partListId,matchSpId,{ResultJSON:result?JSON.stringify(result):''});
}

async function syncSeed(){
  // Force merge SEED_MATCHES (update results if missing)
  await seedMatchesToSP(SEED_MATCHES);
}

// ─── App ──────────────────────────────────────────────────────────────────
export default function App(){
  const[auth,setAuth]=useState(null);      // {name, email}
  const[authLoading,setAuthLoading]=useState(true);
  const[matches,setMatches]=useState(null);
  const[preds,setPreds]=useState(null);
  const[loading,setLoading]=useState(false);
  const[tab,setTab]=useState('partidos');
  const[now,setNow]=useState(new Date());
  const[toast,setToast]=useState(null);

  const showToast=(msg,type='ok')=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};

  // Check if already signed in
  useEffect(()=>{
    if(!IS_CONFIGURED){setAuthLoading(false);return;}
    const app=getMSAL();
    app.initialize().then(()=>{
      app.handleRedirectPromise().then(()=>{
        const acc=getAccount();
        if(acc) setAuth({name:acc.name,email:acc.username});
        setAuthLoading(false);
      }).catch(()=>setAuthLoading(false));
    }).catch(()=>setAuthLoading(false));
    const c=setInterval(()=>setNow(new Date()),30000);
    return()=>clearInterval(c);
  },[]);

  // Load data after auth
  const loadData=useCallback(async()=>{
    setLoading(true);
    try{
      const [m,p]=await Promise.all([loadMatches(),loadAllPredictions()]);
      setMatches(m);setPreds(p);
    }catch(e){
      if(e.message==='LISTS_NOT_FOUND'){setMatches([]);setPreds({});}
      else showToast('Error cargando datos: '+e.message,'err');
    }
    setLoading(false);
  },[]);

  useEffect(()=>{if(auth) loadData();},[auth,loadData]);

  const handleSignIn=async()=>{
    try{const r=await signIn();setAuth({name:r.account.name,email:r.account.username});}
    catch(e){if(e.errorCode!=='user_cancelled') showToast('Error al iniciar sesión','err');}
  };

  const handleSignOut=async()=>{try{await signOut();setAuth(null);}catch{}};

  const handleSavePred=async(matchId,pred)=>{
    try{
      await savePrediction(matchId,pred,auth.email,auth.name);
      const updated={...preds,[matchId]:{...(preds[matchId]||{}),[auth.name]:{...pred}}};
      setPreds(updated);
      showToast('✅ Quiniela guardada');
    }catch(e){showToast('Error guardando: '+e.message,'err');}
  };

  const handleSaveResult=async(match,result)=>{
    try{
      await updateMatchResult(match.spId,result);
      setMatches(m=>m.map(x=>x.id===match.id?{...x,result}:x));
      showToast('✅ Resultado guardado');
    }catch(e){showToast('Error: '+e.message,'err');}
  };

  const handleSync=async()=>{
    setLoading(true);
    try{await syncSeed();await loadData();showToast('✅ Sincronizado');}
    catch(e){showToast('Error sync: '+e.message,'err');setLoading(false);}
  };

  const handleSetup=async()=>{
    setLoading(true);
    try{
      await ensureLists();
      await seedMatchesToSP(SEED_MATCHES);
      await loadData();
      showToast('✅ Plataforma configurada. ¡Listo!');
    }catch(e){showToast('Error de configuración: '+e.message,'err');setLoading(false);}
  };

  // Not configured yet → show setup guide
  if(!IS_CONFIGURED) return <SetupGuide/>;

  if(authLoading) return<Loader msg="Iniciando…"/>;

  if(!auth) return(
    <LoginScreen onLogin={handleSignIn}/>
  );

  const board=(matches||[]).reduce((acc,m)=>{
    if(!m.result) return acc;
    const ko=PM[m.phase]?.ko;
    Object.entries(preds?.[m.id]||{}).forEach(([name,pr])=>{
      if(!acc[name]) acc[name]={pts:0,ex:0,w:0};
      const p=calcPts(pr,m.result,ko);
      if(p===5){acc[name].pts+=5;acc[name].ex++;}
      else if(p===4){acc[name].pts+=4;acc[name].w++;}
      else if(p===3){acc[name].pts+=3;acc[name].w++;}
    });
    return acc;
  },{});
  const boardArr=Object.entries(board).map(([n,s])=>({n,...s})).sort((a,b)=>b.pts-a.pts||b.ex-a.ex||b.w-a.w);

  return(
    <div style={{minHeight:'100vh',background:'#030a03'}}>
      {/* Toast */}
      {toast&&<div style={{position:'fixed',top:16,left:'50%',transform:'translateX(-50%)',zIndex:999,background:toast.type==='ok'?'#15803d':'#b91c1c',color:'#fff',fontWeight:700,fontSize:13,padding:'10px 20px',borderRadius:24,boxShadow:'0 4px 20px rgba(0,0,0,.5)',whiteSpace:'nowrap'}}>{toast.msg}</div>}

      {/* Header */}
      <div style={{background:'linear-gradient(135deg,#021502,#06360a,#021502)',borderBottom:'2px solid rgba(234,179,8,.35)'}}>
        <div style={{maxWidth:560,margin:'0 auto',padding:'13px 16px',display:'flex',alignItems:'center',gap:12}}>
          <div style={{width:46,height:46,background:'rgba(234,179,8,.15)',border:'2px solid rgba(234,179,8,.4)',borderRadius:'50%',display:'flex',alignItems:'center',justifyContent:'center',fontSize:22,flexShrink:0}}>🏆</div>
          <div style={{flex:1}}>
            <div style={{color:'#facc15',fontWeight:900,fontSize:16,letterSpacing:'.1em'}}>QUINIELA MUNDIAL 2026</div>
            <div style={{color:'rgba(234,179,8,.5)',fontSize:8,letterSpacing:'.5em',fontWeight:700}}>◆ SAGA ◆</div>
          </div>
          <div style={{textAlign:'right'}}>
            <div style={{display:'flex',alignItems:'center',gap:7,background:'rgba(22,101,52,.4)',border:'1px solid rgba(34,197,94,.25)',borderRadius:20,padding:'4px 10px',cursor:'default'}}>
              <span style={{fontSize:16}}>👤</span>
              <div>
                <div style={{color:'#86efac',fontSize:11,fontWeight:700,maxWidth:120,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{auth.name}</div>
                <div style={{color:'#4b5563',fontSize:9,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:120}}>{auth.email}</div>
              </div>
            </div>
            <button onClick={handleSignOut} style={{background:'none',border:'none',color:'#374151',fontSize:10,cursor:'pointer',marginTop:3,textDecoration:'underline'}}>Cerrar sesión</button>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{position:'sticky',top:0,zIndex:20,background:'rgba(3,10,3,.96)',backdropFilter:'blur(12px)',borderBottom:'1px solid rgba(255,255,255,.07)'}}>
        <div style={{maxWidth:560,margin:'0 auto',display:'flex'}}>
          {[{id:'tabla',e:'🥇',l:'Tabla'},{id:'partidos',e:'⚽',l:'Partidos'},{id:'admin',e:'⚙️',l:'Admin'}].map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{flex:1,padding:'13px 0',fontSize:12,fontWeight:700,border:'none',background:'none',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',gap:5,color:tab===t.id?'#facc15':'#6b7280',borderBottom:tab===t.id?'2px solid #facc15':'2px solid transparent'}}>
              <span>{t.e}</span><span>{t.l}</span>
            </button>
          ))}
        </div>
      </div>

      {loading&&<div style={{background:'rgba(234,179,8,.07)',borderBottom:'1px solid rgba(234,179,8,.15)',padding:'7px 16px',textAlign:'center',fontSize:11,color:'#ca8a04',fontWeight:600}}><span className="spin">⚙️</span> Conectando con SharePoint…</div>}

      <div style={{maxWidth:560,margin:'0 auto',padding:'20px 12px 80px'}}>
        {tab==='tabla'    &&<TablaView board={boardArr} matches={matches||[]}/>}
        {tab==='partidos' &&<PartidosView matches={matches} preds={preds} me={auth.name} now={now} onSavePred={handleSavePred} onRefresh={loadData}/>}
        {tab==='admin'    &&<AdminView matches={matches} preds={preds} onSaveResult={handleSaveResult} onSync={handleSync} onSetup={handleSetup} loading={loading}/>}
      </div>
    </div>
  );
}

// ── Tabla ─────────────────────────────────────────────────────────────────
function TablaView({board,matches}){
  const done=matches.filter(m=>m.result).length,total=matches.length;
  const M=['🥇','🥈','🥉'];
  return(
    <div>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
        <h2 style={{fontWeight:900,fontSize:20,margin:0}}>Clasificación</h2>
        <span style={{color:'#4b5563',fontSize:11}}>{done}/{total} resultados</span>
      </div>
      <div style={{background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.08)',borderRadius:14,padding:'10px 14px',marginBottom:18,display:'flex',gap:16,flexWrap:'wrap'}}>
        {[['#facc15','5','Marcador exacto'],['#4ade80','4','Ganador+T.Extra'],['#60a5fa','3','Ganador/empate']].map(([c,n,l])=>(
          <div key={n} style={{display:'flex',alignItems:'center',gap:5}}><span style={{color:c,fontWeight:900,fontSize:15}}>{n}</span><span style={{color:'#6b7280',fontSize:11}}>{l}</span></div>
        ))}
      </div>
      {board.length===0?<Empty icon="📋" msg="Sin datos" sub="Los resultados irán apareciendo conforme avancen los partidos"/>
        :board.map((p,i)=>{
          const cc=i===0?'#facc15':i===1?'#d1d5db':i===2?'#fb923c':'#fff';
          const bc=i===0?'rgba(234,179,8,.3)':i===1?'rgba(156,163,175,.2)':i===2?'rgba(249,115,22,.25)':'rgba(255,255,255,.06)';
          const bg=i===0?'rgba(234,179,8,.06)':i===1?'rgba(156,163,175,.06)':i===2?'rgba(249,115,22,.06)':'rgba(255,255,255,.02)';
          return<div key={p.n} style={{display:'flex',alignItems:'center',gap:12,padding:'12px 14px',borderRadius:16,border:`1px solid ${bc}`,background:bg,marginBottom:8}}>
            <div style={{width:28,textAlign:'center',fontSize:i<3?20:13,fontWeight:700,color:'#6b7280'}}>{M[i]||i+1}</div>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:14,color:'#fff'}}>{p.n}</div>
              <div style={{display:'flex',gap:10,marginTop:2}}>
                <span style={{fontSize:10,color:'rgba(250,204,21,.7)'}}>⭐ {p.ex} exactos</span>
                <span style={{fontSize:10,color:'rgba(96,165,250,.7)'}}>✓ {p.w} ganador</span>
              </div>
            </div>
            <div style={{fontWeight:900,fontSize:28,color:cc}}>{p.pts}</div>
            <div style={{color:'#374151',fontSize:10}}>pts</div>
          </div>;
        })
      }
    </div>
  );
}

// ── Partidos ──────────────────────────────────────────────────────────────
function PartidosView({matches,preds,me,now,onSavePred,onRefresh}){
  const[phase,setPhase]=useState('groups_j3');
  const[modal,setModal]=useState(null);
  const[showAll,setShowAll]=useState({});
  const[saving,setSaving]=useState(false);

  if(!matches) return<Loader msg="Cargando partidos…"/>;

  const activePhases=PHASES.filter(p=>p.id==='groups_j3'||matches.some(m=>m.phase===p.id));
  const phaseMatches=[...matches.filter(m=>m.phase===phase)].sort((a,b)=>(a.matchDatetime||'zzz')<(b.matchDatetime||'zzz')?-1:1);
  const byDate=phaseMatches.reduce((acc,m)=>{const d=m.matchDatetime?m.matchDatetime.slice(0,10):'tbd';if(!acc[d])acc[d]=[];acc[d].push(m);return acc;},{});
  const dates=Object.keys(byDate).sort();
  const fmtDH=d=>{if(d==='tbd') return 'Fecha TBD';const dt=new Date(d+'T12:00');return dt.toLocaleDateString('es-MX',{weekday:'long',day:'numeric',month:'long'});};

  return(
    <div>
      {/* Info banner */}
      <div style={{background:'rgba(22,101,52,.15)',border:'1px solid rgba(34,197,94,.2)',borderRadius:14,padding:'10px 14px',marginBottom:16,display:'flex',alignItems:'center',gap:10}}>
        <span style={{fontSize:20}}>👤</span>
        <div>
          <div style={{color:'#4ade80',fontWeight:700,fontSize:13}}>{me}</div>
          <div style={{color:'#6b7280',fontSize:11}}>Tu sesión de Microsoft. Tus quinielas se guardan directamente en SharePoint.</div>
        </div>
        <button onClick={onRefresh} style={{marginLeft:'auto',background:'rgba(255,255,255,.06)',border:'1px solid rgba(255,255,255,.12)',color:'#9ca3af',fontSize:11,fontWeight:700,padding:'5px 10px',borderRadius:20,cursor:'pointer',flexShrink:0}}>↻ Recargar</button>
      </div>

      {/* Phase tabs */}
      <div className="hs" style={{display:'flex',gap:6,overflowX:'auto',paddingBottom:10,marginBottom:16}}>
        {activePhases.map(p=><button key={p.id} onClick={()=>setPhase(p.id)} style={{flexShrink:0,fontSize:12,padding:'6px 14px',borderRadius:24,fontWeight:700,cursor:'pointer',border:'none',background:phase===p.id?'#eab308':'rgba(255,255,255,.05)',color:phase===p.id?'#000':'#9ca3af'}}>{p.label}</button>)}
      </div>

      {phaseMatches.length===0?<Empty icon="📅" msg="Sin partidos en esta fase" sub="Vuelve pronto"/>
        :dates.map(d=>(
          <div key={d}>
            <div style={{fontSize:11,color:'#4ade80',fontWeight:700,textTransform:'capitalize',margin:'16px 0 8px 2px'}}>📅 {fmtDH(d)}</div>
            {byDate[d].map(m=>{
              const myPred=preds?.[m.id]?.[me];const ko=PM[m.phase]?.ko;
              const myPts=myPred&&m.result?calcPts(myPred,m.result,ko):null;
              const status=getStatus(m,now);const cd=lockCD(m,now);const open=showAll[m.id];
              return(
                <div key={m.id} style={{background:'rgba(255,255,255,.03)',border:`1px solid ${status==='live'?'rgba(239,68,68,.3)':status==='locked'?'rgba(255,165,0,.15)':'rgba(255,255,255,.07)'}`,borderRadius:20,overflow:'hidden',marginBottom:10}}>
                  <div style={{padding:'14px 16px 0'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8}}>
                      <div>
                        <div style={{fontSize:10,color:'#4b5563',textTransform:'uppercase',letterSpacing:'.06em'}}>Grupo {m.group}</div>
                        {m.matchDatetime&&<div style={{fontSize:11,color:status==='live'?'#fca5a5':status==='locked'?'#fcd34d':'#6b7280',marginTop:2,fontWeight:status!=='open'?700:400}}>🕐 {new Date(m.matchDatetime).toLocaleTimeString('es-MX',{hour:'2-digit',minute:'2-digit'})} h · CDMX</div>}
                      </div>
                      <SBadge status={status}/>
                    </div>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8,paddingBottom:14}}>
                      <div style={{flex:1,textAlign:'right',fontWeight:900,fontSize:15,color:'#fff'}}>{m.team1}</div>
                      <div style={{textAlign:'center',padding:'0 8px'}}>
                        {m.result?<div><div style={{background:'rgba(255,255,255,.08)',borderRadius:10,padding:'6px 14px',fontWeight:900,fontSize:22,color:'#fff',letterSpacing:2}}>{m.result.s1} – {m.result.s2}</div>{ko&&<div style={{fontSize:10,color:'#6b7280',marginTop:3}}>{m.result.et?'⏱ T.Extra · ':'90 min · '}<span style={{color:'#86efac'}}>{m.result.winner==='h'?m.team1:m.team2} ✓</span></div>}</div>
                          :<div style={{color:'#374151',fontWeight:900,fontSize:15}}>VS</div>}
                      </div>
                      <div style={{flex:1,fontWeight:900,fontSize:15,color:'#fff'}}>{m.team2}</div>
                    </div>
                  </div>
                  <div style={{borderTop:'1px solid rgba(255,255,255,.06)',padding:'10px 16px',display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:6}}>
                    <div style={{display:'flex',alignItems:'center',gap:7,flexWrap:'wrap'}}>
                      {myPred?(<>
                        <span style={{fontSize:11,color:'#6b7280'}}>Tu quiniela:</span>
                        <span style={{fontSize:11,fontWeight:700,background:'rgba(255,255,255,.08)',borderRadius:8,padding:'2px 9px',color:'#e5e7eb',fontFamily:'monospace'}}>{myPred.s1}–{myPred.s2}{ko?` · ${myPred.winner==='h'?m.team1:m.team2}${myPred.et?' +T':''}`:''}</span>
                        {myPts!==null&&<span style={{fontWeight:900,fontSize:13,color:myPts===5?'#facc15':myPts===4?'#4ade80':myPts===3?'#60a5fa':'#ef4444'}}>{myPts===5?'⭐':myPts>=3?'✓':'✗'} {myPts} pts</span>}
                        {cd&&<span style={{fontSize:10,color:cd.urgent?'#fbbf24':'#d97706',fontWeight:700}}>{cd.label}</span>}
                      </>):(<span style={{fontSize:11,color:'#374151'}}>{cd&&<span style={{color:cd.urgent?'#fbbf24':'#d97706',fontWeight:700,marginRight:6}}>{cd.label}</span>}{status==='open'?'Sin quiniela aún':'—'}</span>)}
                    </div>
                    <div style={{display:'flex',gap:8,alignItems:'center'}}>
                      {m.result&&<button onClick={()=>setShowAll(p=>({...p,[m.id]:!p[m.id]}))} style={S.link}>{open?'Ocultar':'Ver todos'}</button>}
                      {status==='open'&&<button onClick={()=>setModal(m)} style={S.bet}>{myPred?'Editar ✏️':'Apostar 🎯'}</button>}
                      {status==='locked'&&!myPred&&<span style={{fontSize:11,color:'#92400e',fontWeight:700}}>🔒 Cerrado</span>}
                    </div>
                  </div>
                  {open&&(
                    <div style={{borderTop:'1px solid rgba(255,255,255,.05)',padding:'12px 16px 14px',background:'rgba(0,0,0,.2)'}}>
                      <div style={{fontSize:10,textTransform:'uppercase',letterSpacing:'.1em',color:'#4b5563',marginBottom:10,fontWeight:700}}>Quinielas de todos</div>
                      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'2px 12px'}}>
                        {Object.entries(preds?.[m.id]||{}).map(([n,pr])=>{
                          const pts=pr&&m.result?calcPts(pr,m.result,ko):null;
                          return<div key={n} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'2px 0'}}>
                            <span style={{fontSize:11,color:'#9ca3af',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:80}}>{n}</span>
                            <span style={{fontSize:10,color:'#d1d5db',fontFamily:'monospace'}}>{pr.s1}–{pr.s2}{ko?` ${pr.winner==='h'?m.team1:m.team2}${pr.et?' +T':''}`:''}</span>
                            {pts!==null&&<span style={{fontSize:12,fontWeight:900,color:pts===5?'#facc15':pts===4?'#4ade80':pts===3?'#60a5fa':'#ef4444',minWidth:14,textAlign:'right'}}>{pts}</span>}
                          </div>;
                        })}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ))
      }
      {modal&&<PredictModal match={modal} ko={PM[modal.phase]?.ko} existing={preds?.[modal.id]?.[me]} saving={saving}
        onClose={()=>setModal(null)}
        onSave={async pred=>{
          if(getStatus(modal,new Date())!=='open'){alert('⏱ Las predicciones ya están cerradas.');setModal(null);return;}
          setSaving(true);await onSavePred(modal.id,pred);setSaving(false);setModal(null);
        }}/>}
    </div>
  );
}

function SBadge({status}){
  if(status==='played') return<span style={{fontSize:10,fontWeight:700,background:'rgba(22,101,52,.5)',color:'#4ade80',borderRadius:20,padding:'2px 10px',whiteSpace:'nowrap'}}>✓ Jugado</span>;
  if(status==='live')   return<span style={{fontSize:10,fontWeight:700,background:'rgba(239,68,68,.2)',color:'#fca5a5',borderRadius:20,padding:'2px 10px',whiteSpace:'nowrap',display:'flex',alignItems:'center',gap:4}}><span className="pulse" style={{width:6,height:6,background:'#ef4444',borderRadius:'50%',display:'inline-block'}}/>En curso</span>;
  if(status==='locked') return<span style={{fontSize:10,fontWeight:700,background:'rgba(180,83,9,.3)',color:'#fcd34d',borderRadius:20,padding:'2px 10px',whiteSpace:'nowrap'}}>🔒 Cerrado</span>;
  return null;
}

function PredictModal({match,ko,existing,saving,onClose,onSave}){
  const[s1,setS1]=useState(existing?.s1??'');const[s2,setS2]=useState(existing?.s2??'');
  const[winner,setWinner]=useState(existing?.winner??'');const[et,setEt]=useState(ko?(existing?.et??null):false);
  const valid=s1!==''&&s2!==''&&(!ko||(winner!==''&&et!==null));
  return(
    <div onClick={onClose} style={{position:'fixed',inset:0,zIndex:50,background:'rgba(0,0,0,.85)',display:'flex',alignItems:'flex-end',justifyContent:'center'}}>
      <div onClick={e=>e.stopPropagation()} style={{background:'#0d1a0d',border:'1px solid rgba(255,255,255,.12)',borderRadius:'24px 24px 0 0',width:'100%',maxWidth:520,padding:'22px 20px 40px'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:14}}>
          <h3 style={{margin:0,fontWeight:900,fontSize:18,color:'#fff'}}>Tu quiniela ⚽</h3>
          <button onClick={onClose} style={{background:'rgba(255,255,255,.08)',border:'none',color:'#9ca3af',width:32,height:32,borderRadius:'50%',cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{textAlign:'center',color:'#9ca3af',fontSize:14,marginBottom:16}}>
          <span style={{fontWeight:900,color:'#fff'}}>{match.team1}</span><span style={{color:'#374151',margin:'0 10px'}}>vs</span><span style={{fontWeight:900,color:'#fff'}}>{match.team2}</span>
          {match.matchDatetime&&<div style={{fontSize:11,color:'#4b5563',marginTop:4}}>📅 {fmtDt(match.matchDatetime)}</div>}
        </div>
        <p style={{fontSize:11,color:'#6b7280',textAlign:'center',margin:'0 0 10px'}}>Marcador al finalizar los 90 min</p>
        <div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:14,marginBottom:ko?20:14}}>
          {[{v:s1,set:setS1,l:match.team1},{v:s2,set:setS2,l:match.team2}].reduce((a,inp,i)=>[...a,
            i===1&&<span key="d" style={{color:'#374151',fontSize:26,fontWeight:900,marginTop:22}}>–</span>,
            <div key={inp.l} style={{textAlign:'center'}}><div style={{fontSize:11,color:'#6b7280',marginBottom:8}}>{inp.l}</div>
              <input type="number" inputMode="numeric" min="0" max="20" value={inp.v} onChange={e=>inp.set(e.target.value)}
                style={{width:76,height:76,textAlign:'center',fontSize:30,fontWeight:900,background:'rgba(255,255,255,.08)',border:'2px solid rgba(255,255,255,.15)',borderRadius:18,color:'#fff',outline:'none',display:'block'}}
                onFocus={e=>e.target.style.borderColor='#eab308'} onBlur={e=>e.target.style.borderColor='rgba(255,255,255,.15)'}/>
            </div>
          ],[])}
        </div>
        {ko&&<div style={{display:'flex',flexDirection:'column',gap:14,marginBottom:14}}>
          <div><p style={{fontSize:11,color:'#6b7280',textAlign:'center',margin:'0 0 8px'}}>¿Quién gana la llave?</p>
            <div style={{display:'flex',gap:8}}>{[{v:'h',l:match.team1},{v:'a',l:match.team2}].map(o=><button key={o.v} onClick={()=>setWinner(o.v)} style={{flex:1,padding:'10px 4px',fontSize:12,fontWeight:700,cursor:'pointer',borderRadius:14,border:'none',background:winner===o.v?'#15803d':'rgba(255,255,255,.05)',color:winner===o.v?'#fff':'#6b7280'}}>{o.l}</button>)}</div>
          </div>
          <div><p style={{fontSize:11,color:'#6b7280',textAlign:'center',margin:'0 0 8px'}}>¿Habrá tiempo extra / penales?</p>
            <div style={{display:'flex',gap:8}}>{[{v:true,l:'Sí ⏱️'},{v:false,l:'No 💨'}].map(o=><button key={String(o.v)} onClick={()=>setEt(o.v)} style={{flex:1,padding:'10px 4px',fontSize:12,fontWeight:700,cursor:'pointer',borderRadius:14,border:'none',background:et===o.v?'#1d4ed8':'rgba(255,255,255,.05)',color:et===o.v?'#fff':'#6b7280'}}>{o.l}</button>)}</div>
          </div>
        </div>}
        <div style={{display:'flex',justifyContent:'center',gap:14,fontSize:11,color:'#4b5563',marginBottom:16}}>
          <span><span style={{color:'#facc15',fontWeight:900}}>5</span> exacto</span>
          {ko&&<span><span style={{color:'#4ade80',fontWeight:900}}>4</span> ganador+T.extra</span>}
          <span><span style={{color:'#60a5fa',fontWeight:900}}>3</span> ganador</span>
        </div>
        <button onClick={()=>valid&&!saving&&onSave({s1:+s1,s2:+s2,...(ko?{winner,et}:{})})} disabled={!valid||saving}
          style={{width:'100%',padding:'15px 0',fontWeight:900,fontSize:15,borderRadius:18,border:'none',cursor:valid&&!saving?'pointer':'default',background:valid&&!saving?'#eab308':'rgba(255,255,255,.06)',color:valid&&!saving?'#000':'#374151'}}>
          {saving?'Guardando…':'Guardar quiniela 🎯'}
        </button>
      </div>
    </div>
  );
}

// ── Admin ─────────────────────────────────────────────────────────────────
function AdminView({matches,preds,onSaveResult,onSync,onSetup,loading}){
  const[pass,setPass]=useState('');const[authed,setAuthed]=useState(false);const[sel,setSel]=useState(null);const[form,setForm]=useState({});
  if(!authed) return(
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',padding:'60px 0',gap:14}}>
      <div style={{fontSize:52}}>🔐</div>
      <p style={{fontWeight:900,fontSize:17,color:'#fff',margin:0}}>Panel Admin</p>
      <input type="password" value={pass} onChange={e=>setPass(e.target.value)} placeholder="Contraseña"
        onKeyDown={e=>e.key==='Enter'&&(pass===CONFIG.adminPass?setAuthed(true):alert('❌ Contraseña incorrecta'))}
        style={{...S.inp,textAlign:'center',width:200,marginBottom:12}} onFocus={e=>e.target.style.borderColor='#eab308'} onBlur={e=>e.target.style.borderColor='rgba(255,255,255,.12)'}/>
      <button onClick={()=>pass===CONFIG.adminPass?setAuthed(true):alert('❌ Contraseña incorrecta')}
        style={{background:'#eab308',border:'none',color:'#000',fontWeight:900,padding:'12px 36px',borderRadius:14,cursor:'pointer'}}>Entrar</button>
    </div>
  );
  const pick=m=>{setSel(m);setForm(m.result?{s1:m.result.s1,s2:m.result.s2,et:m.result.et??false,winner:m.result.winner??'h'}:{s1:'',s2:'',et:false,winner:'h'});};
  const save=async()=>{
    if(form.s1===''||form.s2==='') return alert('Ingresa el marcador');
    const ko=PM[sel.phase]?.ko;
    const result={s1:+form.s1,s2:+form.s2};
    if(ko){result.et=form.et;result.winner=form.winner;}
    else result.winner=+form.s1>+form.s2?'h':+form.s2>+form.s1?'a':'d';
    await onSaveResult(sel,result);setSel(null);
  };
  const PB=({active,onClick,label,ac='#15803d'})=><button onClick={onClick} style={{flex:1,padding:'9px 4px',fontSize:13,fontWeight:700,cursor:'pointer',borderRadius:13,border:'none',background:active?ac:'rgba(255,255,255,.05)',color:active?'#fff':'#6b7280'}}>{label}</button>;

  if(sel){const ko=PM[sel.phase]?.ko;return(
    <div>
      <button onClick={()=>setSel(null)} style={{background:'none',border:'none',color:'#6b7280',fontSize:12,cursor:'pointer',textDecoration:'underline',marginBottom:16}}>← Volver</button>
      <div style={{background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.1)',borderRadius:20,padding:18,display:'flex',flexDirection:'column',gap:16}}>
        <div><p style={{fontSize:11,color:'#6b7280',margin:'0 0 3px'}}>{PM[sel.phase]?.label}</p><p style={{fontSize:17,fontWeight:900,color:'#fff',margin:0}}>{sel.team1} vs {sel.team2}</p>{sel.matchDatetime&&<p style={{fontSize:11,color:'#4b5563',margin:'3px 0 0'}}>📅 {fmtDt(sel.matchDatetime)}</p>}</div>
        <div><p style={{fontSize:11,color:'#9ca3af',margin:'0 0 10px'}}>Marcador (90 min)</p>
          <div style={{display:'flex',alignItems:'center',gap:12}}>
            {[{v:form.s1,set:v=>setForm({...form,s1:v}),l:sel.team1},{v:form.s2,set:v=>setForm({...form,s2:v}),l:sel.team2}].reduce((a,inp,i)=>[...a,
              i===1&&<span key="d" style={{color:'#374151',fontSize:22,fontWeight:900,marginTop:20}}>–</span>,
              <div key={inp.l} style={{textAlign:'center'}}><div style={{fontSize:10,color:'#4b5563',marginBottom:5}}>{inp.l}</div>
                <input type="number" min="0" max="20" value={inp.v} onChange={e=>inp.set(e.target.value)} style={{width:68,height:68,textAlign:'center',fontSize:26,fontWeight:900,background:'rgba(255,255,255,.08)',border:'2px solid rgba(255,255,255,.15)',borderRadius:14,color:'#fff',outline:'none',display:'block'}}/>
              </div>
            ],[])}
          </div>
        </div>
        {ko&&<><div><p style={{fontSize:11,color:'#9ca3af',margin:'0 0 8px'}}>¿Hubo tiempo extra?</p><div style={{display:'flex',gap:8}}><PB active={form.et===true} onClick={()=>setForm({...form,et:true})} label="Sí ⏱️" ac="#1d4ed8"/><PB active={form.et===false} onClick={()=>setForm({...form,et:false})} label="No, 90 min" ac="#1d4ed8"/></div></div>
        <div><p style={{fontSize:11,color:'#9ca3af',margin:'0 0 8px'}}>Ganador final</p><div style={{display:'flex',gap:8}}><PB active={form.winner==='h'} onClick={()=>setForm({...form,winner:'h'})} label={sel.team1}/><PB active={form.winner==='a'} onClick={()=>setForm({...form,winner:'a'})} label={sel.team2}/></div></div></>}
        <button onClick={save} style={{background:'#16a34a',border:'none',color:'#fff',fontWeight:900,padding:'14px 0',borderRadius:16,cursor:'pointer',fontSize:15}}>✅ Guardar resultado</button>
      </div>
    </div>
  );}

  const pending=(matches||[]).filter(m=>!m.result),played=(matches||[]).filter(m=>m.result);
  const listsReady=(matches||[]).length>0;
  return(
    <div>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
        <h2 style={{fontWeight:900,fontSize:19,margin:0}}>⚙️ Admin</h2>
        <div style={{display:'flex',gap:8}}>
          <button onClick={onSync} disabled={loading} style={{background:'rgba(234,179,8,.15)',border:'1px solid rgba(234,179,8,.3)',color:'#facc15',fontSize:11,fontWeight:700,padding:'5px 10px',borderRadius:24,cursor:'pointer',opacity:loading?.6:1}}>⚡ Sync resultados</button>
          <button onClick={()=>setAuthed(false)} style={{background:'rgba(239,68,68,.15)',border:'1px solid rgba(239,68,68,.3)',color:'#f87171',fontSize:11,fontWeight:700,padding:'5px 10px',borderRadius:24,cursor:'pointer'}}>Salir</button>
        </div>
      </div>

      <div style={{background:'rgba(234,179,8,.06)',border:'1px solid rgba(234,179,8,.15)',borderRadius:14,padding:'11px 14px',marginBottom:16,fontSize:11,color:'#ca8a04',lineHeight:1.6}}>
        <strong style={{color:'#facc15'}}>⚡ Sync resultados</strong> — cuando Claude actualice marcadores o agregue 16vos, presiona este botón y los resultados se actualizan en SharePoint para todos automáticamente.
      </div>

      {!listsReady&&(
        <div style={{background:'rgba(37,99,235,.12)',border:'1px solid rgba(37,99,235,.3)',borderRadius:16,padding:18,marginBottom:20}}>
          <p style={{fontWeight:700,color:'#93c5fd',margin:'0 0 8px'}}>🚀 Primera vez — Configurar plataforma</p>
          <p style={{fontSize:12,color:'#6b7280',margin:'0 0 12px'}}>Creará las listas en SharePoint y cargará los 24 partidos de Jornada 3. Solo se hace una vez.</p>
          <button onClick={onSetup} disabled={loading} style={{background:'#2563eb',border:'none',color:'#fff',fontWeight:900,padding:'11px 24px',borderRadius:14,cursor:'pointer',fontSize:14,opacity:loading?.6:1}}>
            {loading?<span><span className="spin" style={{marginRight:6}}>⚙️</span>Configurando…</span>:'🚀 Configurar ahora'}
          </button>
        </div>
      )}

      {pending.length>0&&<div style={{marginBottom:20}}>
        <p style={{fontSize:10,textTransform:'uppercase',letterSpacing:'.1em',color:'#6b7280',fontWeight:700,margin:'0 0 10px'}}>Pendientes de resultado</p>
        {pending.map(m=><button key={m.id} onClick={()=>pick(m)} style={{width:'100%',textAlign:'left',background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.08)',borderRadius:16,padding:'11px 16px',cursor:'pointer',marginBottom:7,display:'block'}}
          onMouseEnter={e=>e.currentTarget.style.borderColor='rgba(234,179,8,.4)'} onMouseLeave={e=>e.currentTarget.style.borderColor='rgba(255,255,255,.08)'}>
          <div style={{fontSize:10,color:'#4b5563'}}>{PM[m.phase]?.label}{m.group?` · Grupo ${m.group}`:''}{m.matchDatetime?` · ${fmtDt(m.matchDatetime)}`:''}</div>
          <div style={{fontSize:14,fontWeight:700,color:'#fff'}}>{m.team1} vs {m.team2}</div>
        </button>)}
      </div>}

      {played.length>0&&<div>
        <p style={{fontSize:10,textTransform:'uppercase',letterSpacing:'.1em',color:'#6b7280',fontWeight:700,margin:'0 0 10px'}}>Con resultado</p>
        {played.map(m=><div key={m.id} style={{display:'flex',alignItems:'center',background:'rgba(22,101,52,.1)',border:'1px solid rgba(22,101,52,.25)',borderRadius:16,padding:'10px 14px',gap:8,marginBottom:7}}>
          <div style={{flex:1,cursor:'pointer'}} onClick={()=>pick(m)}>
            <div style={{fontSize:10,color:'#4b5563'}}>{PM[m.phase]?.label}{m.group?` · Grupo ${m.group}`:''}</div>
            <div style={{fontSize:13,fontWeight:700,color:'#e5e7eb',display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>{m.team1} vs {m.team2}<span style={{color:'#4ade80',fontFamily:'monospace'}}>{m.result.s1}–{m.result.s2}</span>{PM[m.phase]?.ko&&m.result.winner&&<span style={{fontSize:10,color:'#86efac'}}>{m.result.winner==='h'?m.team1:m.team2} ✓{m.result.et?' +T':''}</span>}</div>
          </div>
          <button onClick={()=>onSaveResult(m,null)} style={{background:'none',border:'none',color:'#374151',fontSize:18,cursor:'pointer'}}>✕</button>
        </div>)}
      </div>}
    </div>
  );
}

// ── Setup Guide ───────────────────────────────────────────────────────────
function SetupGuide(){
  const url=window.location.href.replace(/[?#].*/,'');
  return(
    <div style={{minHeight:'100vh',background:'#030a03',padding:'0 0 60px'}}>
      <div style={{background:'linear-gradient(135deg,#021502,#06360a)',borderBottom:'2px solid rgba(234,179,8,.35)',padding:'20px 20px'}}>
        <div style={{maxWidth:560,margin:'0 auto',display:'flex',alignItems:'center',gap:14}}>
          <div style={{width:50,height:50,background:'rgba(234,179,8,.15)',border:'2px solid rgba(234,179,8,.4)',borderRadius:'50%',display:'flex',alignItems:'center',justifyContent:'center',fontSize:24}}>🏆</div>
          <div><div style={{color:'#facc15',fontWeight:900,fontSize:17,letterSpacing:'.1em'}}>QUINIELA MUNDIAL 2026</div><div style={{color:'rgba(234,179,8,.5)',fontSize:9,letterSpacing:'.5em',fontWeight:700}}>◆ SAGA ◆</div></div>
        </div>
      </div>
      <div style={{maxWidth:560,margin:'0 auto',padding:'30px 20px'}}>
        <div style={{background:'rgba(37,99,235,.1)',border:'1px solid rgba(37,99,235,.3)',borderRadius:20,padding:'20px 22px',marginBottom:24}}>
          <p style={{fontWeight:900,color:'#93c5fd',fontSize:16,margin:'0 0 6px'}}>⚙️ Configuración inicial requerida</p>
          <p style={{color:'#6b7280',fontSize:13,lineHeight:1.6,margin:0}}>Hay que registrar la app en Azure AD y llenar 3 campos en el HTML. Solo una vez. Toma ~10 minutos.</p>
        </div>

        {[
          {n:1,title:'Registrar app en Azure AD',color:'#2563eb',steps:[
            <>Ve a <a href="https://portal.azure.com" target="_blank" style={{color:'#60a5fa'}}>portal.azure.com</a> → <strong>Azure Active Directory → Registros de aplicaciones → Nueva</strong></>,
            <>Nombre: <strong>Quiniela SAGA</strong> · Tipos de cuenta: <em>Solo esta organización</em></>,
            <>URI de redirección: tipo <strong>SPA</strong>, valor: <code style={{background:'rgba(255,255,255,.08)',padding:'2px 6px',borderRadius:6,fontSize:11}}>{url}</code></>,
            'Clic en "Registrar"',
          ]},
          {n:2,title:'Copiar Client ID y Tenant ID',color:'#7c3aed',steps:[
            'En la página de la app registrada, copia el "Application (client) ID" → pégalo en clientId del HTML',
            'Copia el "Directory (tenant) ID" → pégalo en tenantId del HTML',
          ]},
          {n:3,title:'Permisos de API',color:'#065f46',steps:[
            <>Ve a <strong>Permisos de API → Agregar permiso → Microsoft Graph → Delegados</strong></>,
            <><strong>Sites.ReadWrite.All</strong> + <strong>Sites.Manage.All</strong> → Agregar</>,
            'Clic en "Conceder consentimiento de administrador" (requiere ser admin de Azure AD)',
          ]},
          {n:4,title:'URL de SharePoint y subir HTML',color:'#92400e',steps:[
            'Copia la URL de tu sitio SharePoint (ej: https://empresa.sharepoint.com/sites/misitio)',
            'Pégala en siteUrl del HTML',
            <>Sube el HTML a GitHub Pages (o SharePoint) y asegúrate de que la URL coincida con la URI de redirección del paso 1</>,
          ]},
          {n:5,title:'Primera vez: configurar plataforma',color:'#166534',steps:[
            'Abre el HTML configurado en el navegador',
            'Inicia sesión con tu cuenta de Microsoft',
            'Ve a Admin (contraseña: saga2026) y presiona "🚀 Configurar ahora"',
            'Se crearán las listas en SharePoint y se cargarán los 24 partidos automáticamente',
          ]},
        ].map(s=>(
          <div key={s.n} style={{marginBottom:18,background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.08)',borderRadius:16,overflow:'hidden'}}>
            <div style={{background:`rgba(0,0,0,.3)`,borderBottom:'1px solid rgba(255,255,255,.07)',padding:'11px 16px',display:'flex',alignItems:'center',gap:10}}>
              <div style={{width:26,height:26,background:s.color,borderRadius:'50%',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:900,fontSize:13,flexShrink:0}}>{s.n}</div>
              <span style={{fontWeight:700,color:'#e5e7eb',fontSize:14}}>{s.title}</span>
            </div>
            <div style={{padding:'12px 16px'}}>
              {s.steps.map((st,i)=><div key={i} style={{display:'flex',gap:10,marginBottom:i<s.steps.length-1?8:0}}>
                <span style={{color:'#374151',fontWeight:700,fontSize:12,flexShrink:0,marginTop:1}}>→</span>
                <span style={{color:'#9ca3af',fontSize:12,lineHeight:1.5}}>{st}</span>
              </div>)}
            </div>
          </div>
        ))}

        <div style={{background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.08)',borderRadius:14,padding:'14px 16px',marginTop:8}}>
          <p style={{fontSize:12,color:'#6b7280',lineHeight:1.6,margin:0}}>
            <strong style={{color:'#e5e7eb'}}>Campos a modificar en el HTML</strong> (al inicio del script, sección CONFIG):<br/>
            <code style={{background:'rgba(255,255,255,.06)',display:'block',padding:'10px 12px',borderRadius:10,marginTop:8,fontSize:11,lineHeight:1.8,color:'#86efac'}}>
              clientId: "tu-client-id-aquí",<br/>
              tenantId: "tu-tenant-id-aquí",<br/>
              siteUrl: "https://empresa.sharepoint.com/sites/sitio",
            </code>
          </p>
        </div>
      </div>
    </div>
  );
}

function LoginScreen({onLogin}){
  return(
    <div style={{minHeight:'100vh',background:'#030a03',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:24,padding:20}}>
      <div style={{width:72,height:72,background:'rgba(234,179,8,.15)',border:'2px solid rgba(234,179,8,.4)',borderRadius:'50%',display:'flex',alignItems:'center',justifyContent:'center',fontSize:36}}>🏆</div>
      <div style={{textAlign:'center'}}>
        <div style={{color:'#facc15',fontWeight:900,fontSize:22,letterSpacing:'.1em'}}>QUINIELA MUNDIAL 2026</div>
        <div style={{color:'rgba(234,179,8,.5)',fontSize:10,letterSpacing:'.55em',fontWeight:700,marginTop:4}}>◆ SAGA ◆</div>
      </div>
      <div style={{textAlign:'center',color:'#6b7280',fontSize:14,maxWidth:280,lineHeight:1.6}}>Inicia sesión con tu cuenta de Microsoft para ver y registrar tus quinielas.</div>
      <button onClick={onLogin} style={{display:'flex',alignItems:'center',gap:12,background:'#fff',border:'none',borderRadius:14,padding:'14px 28px',cursor:'pointer',fontWeight:700,fontSize:15,color:'#111',boxShadow:'0 4px 20px rgba(0,0,0,.4)'}}>
        <svg width="20" height="20" viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>
        Iniciar sesión con Microsoft
      </button>
      <p style={{color:'#1f2937',fontSize:11}}>Tus quinielas se guardan directamente en SharePoint</p>
    </div>
  );
}

function Loader({msg}){return<div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'50vh',gap:16}}><span className="spin" style={{fontSize:40}}>⚽</span><p style={{color:'#4ade80',fontWeight:600}}>{msg||'Cargando…'}</p></div>;}
function Empty({icon,msg,sub}){return<div style={{textAlign:'center',padding:'50px 0',color:'#374151'}}><div style={{fontSize:40,marginBottom:10}}>{icon}</div><p style={{fontWeight:600,margin:0,color:'#4b5563'}}>{msg}</p>{sub&&<p style={{fontSize:12,margin:'5px 0 0',color:'#1f2937'}}>{sub}</p>}</div>;}

const S={
  inp:{background:'rgba(255,255,255,.06)',border:'1px solid rgba(255,255,255,.1)',borderRadius:12,padding:'9px 12px',color:'#fff',fontSize:13,outline:'none'},
  link:{background:'none',border:'none',color:'#6b7280',fontSize:11,cursor:'pointer',textDecoration:'underline'},
  bet:{background:'#eab308',border:'none',color:'#000',fontWeight:900,fontSize:11,padding:'5px 14px',borderRadius:24,cursor:'pointer'},
};

ReactDOM.createRoot(document.getElementById('root')).render(<App/>);
</script>
</body>
</html>
