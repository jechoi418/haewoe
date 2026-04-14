"""
해외법인 손익 대시보드 웹 서버
- Render.com 무료 배포용
- 실행: gunicorn server:app --bind 0.0.0.0:$PORT
- 로컬: python server.py → http://localhost:5000
"""

import os, io
from flask import Flask, request, jsonify, render_template_string

try:
    import pandas as pd
except ImportError:
    raise SystemExit("pip install flask pandas openpyxl gunicorn")

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

CORPS_ALL = [
    '미국','멕시코','독일','이태리','일본','중국','싱가포르',
    '인도','베트남','태국','말레이시아','인도네시아','우크라이나','미얀마','합 계'
]
CORP_START = {
    '미국':2,'멕시코':11,'독일':20,'이태리':29,'일본':38,'중국':47,
    '싱가포르':56,'인도':65,'베트남':74,'태국':83,'말레이시아':92,
    '인도네시아':101,'우크라이나':110,'미얀마':119,'합 계':128
}


def get_val(src, corp, ind, col):
    s = CORP_START[corp]
    for i in range(s, s + 18):
        if i < len(src) and src.iloc[i][1] == ind:
            v = src.iloc[i][col]
            if pd.isna(v) or float(v) == 0.0:
                return None
            return float(v)
    return None


def extract(fileobj):
    src = pd.read_excel(fileobj, header=None)
    result = {}
    for corp in CORPS_ALL:
        key = 'total' if corp == '합 계' else corp
        result[key] = {}
        for ind in ['매출액', '영업이익']:
            mp, mf, ma = [], [], []
            for m in range(1, 13):
                base = 3 + (m - 1) * 4
                mp.append(get_val(src, corp, ind, base))
                mf.append(get_val(src, corp, ind, base + 1))
                ma.append(get_val(src, corp, ind, base + 2))
            short = 'sl' if ind == '매출액' else 'op'
            yp   = get_val(src, corp, ind, 52)
            ya   = get_val(src, corp, ind, 53)
            yach = get_val(src, corp, ind, 54)
            result[key][short] = {
                'mp':   [round(v, 1) if v is not None else None for v in mp],
                'mf':   [round(v, 1) if v is not None else None for v in mf],
                'ma':   [round(v, 1) if v is not None else None for v in ma],
                'yp':   round(yp,   1) if yp   is not None else None,
                'ya':   round(ya,   1) if ya   is not None else None,
                'yach': round(yach, 4) if yach is not None else None,
            }
    return result


PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>해외법인 손익 대시보드</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F0F2F5;color:#1a1a2e;font-size:13px}
.page{max-width:1400px;margin:0 auto;padding:20px}
.hdr{background:#1B2A4A;border-radius:12px;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:10px}
.hdr h1{font-size:15px;font-weight:600;color:#fff}
.hdr-meta{font-size:11px;color:#7fa8c8;text-align:right;line-height:1.7}
.upload-zone{background:#fff;border:2px dashed #cbd5e1;border-radius:12px;padding:20px;text-align:center;margin-bottom:12px;transition:border-color .2s,background .2s}
.upload-zone:hover,.upload-zone.drag{border-color:#185FA5;background:#EFF6FF}
.upload-zone input{display:none}
.upload-zone label{cursor:pointer;display:block}
.up-icon{width:40px;height:40px;margin:0 auto 8px;background:#EFF6FF;border-radius:50%;display:flex;align-items:center;justify-content:center}
.up-icon svg{width:20px;height:20px;color:#185FA5}
.up-title{font-size:14px;font-weight:600;color:#1e293b;margin-bottom:3px}
.up-sub{font-size:12px;color:#64748b}
.up-btn{display:inline-block;margin-top:10px;padding:8px 20px;background:#185FA5;color:#fff;border-radius:8px;font-size:13px;font-weight:600;border:none;cursor:pointer;font-family:inherit}
.up-btn:hover{background:#0C447C}
.up-status{margin-top:8px;font-size:12px;color:#64748b;min-height:18px}
.up-status.ok{color:#166534;font-weight:600}
.up-status.err{color:#991b1b}
.slicer{background:#fff;border-radius:10px;padding:10px 18px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px;border:1px solid #e2e8f0}
.sl-lbl{font-size:10px;font-weight:600;color:#94a3b8;letter-spacing:.6px;text-transform:uppercase;white-space:nowrap}
.sl-sep{width:1px;height:20px;background:#e2e8f0;margin:0 4px}
.sl{padding:5px 13px;font-size:12px;border:1px solid #e2e8f0;border-radius:999px;cursor:pointer;background:transparent;color:#64748b;transition:all .15s;white-space:nowrap;font-family:inherit}
.sl:hover{background:#f1f5f9;color:#1e293b}
.sl.on{background:#1B2A4A;color:#fff;border-color:#1B2A4A;font-weight:600}
.notice{background:#FAEEDA;border-radius:8px;padding:8px 14px;font-size:11px;color:#633806;margin-bottom:12px;line-height:1.6}
.krow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px}
.kpi{background:#fff;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0}
.kl{font-size:11px;color:#64748b;margin-bottom:5px}
.kv{font-size:20px;font-weight:600;line-height:1;letter-spacing:-.5px}
.ka{display:inline-block;font-size:12px;font-weight:600;padding:2px 8px;border-radius:999px;margin-top:6px}
.crow{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.cc{background:#fff;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0}
.cct{font-size:10px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.leg{display:flex;gap:12px;font-size:11px;color:#64748b;margin-bottom:8px;flex-wrap:wrap}
.leg span{display:flex;align-items:center;gap:4px}
.lsq{width:9px;height:9px;border-radius:2px;flex-shrink:0}
.tw{overflow-x:auto;background:#fff;border-radius:10px;border:1px solid #e2e8f0}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#1B2A4A;color:#fff;padding:8px 10px;font-weight:600;font-size:11px;text-align:right;white-space:nowrap}
th:first-child{text-align:left}
td{padding:7px 10px;border-bottom:1px solid #f1f5f9;text-align:right;white-space:nowrap}
td:first-child{text-align:left;font-weight:600}
tr.tt td{background:#EFF6FF;color:#1e40af;font-weight:600}
tr:not(.tt):hover td{background:#f8fafc}
.ap{display:inline-block;padding:2px 7px;border-radius:999px;font-size:11px;font-weight:600}
.g{background:#dcfce7;color:#166534}.a{background:#fef9c3;color:#854d0e}.r{background:#fee2e2;color:#991b1b}
@media(max-width:900px){.krow{grid-template-columns:repeat(2,1fr)}.crow{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="page">
<div class="hdr">
  <h1>해외법인 매출액 · 영업이익 분석</h1>
  <div class="hdr-meta" id="hdr-meta">엑셀 파일을 업로드하면 대시보드가 자동 생성됩니다</div>
</div>
<div class="upload-zone" id="upzone">
  <label for="finput">
    <div class="up-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/>
      </svg>
    </div>
    <div class="up-title">SAP 엑셀 파일 업로드</div>
    <div class="up-sub">해외법인_월별_손익.xlsx · 최대 20MB</div>
    <button class="up-btn" onclick="document.getElementById('finput').click();event.preventDefault()">파일 선택</button>
  </label>
  <input type="file" id="finput" accept=".xlsx,.xls">
  <div class="up-status" id="upstat"></div>
</div>
<div id="dash" style="display:none">
  <div class="slicer" id="slbar">
    <span class="sl-lbl">보기</span>
    <button class="sl on" data-k="total">종합</button>
    <div class="sl-sep"></div>
    <span class="sl-lbl">법인</span>
    <button class="sl" data-k="미국">미국</button><button class="sl" data-k="멕시코">멕시코</button><button class="sl" data-k="독일">독일</button><button class="sl" data-k="이태리">이태리</button><button class="sl" data-k="일본">일본</button><button class="sl" data-k="중국">중국</button><button class="sl" data-k="싱가포르">싱가포르</button><button class="sl" data-k="인도">인도</button><button class="sl" data-k="베트남">베트남</button><button class="sl" data-k="태국">태국</button><button class="sl" data-k="말레이시아">말레이시아</button><button class="sl" data-k="인도네시아">인도네시아</button><button class="sl" data-k="우크라이나">우크라이나</button><button class="sl" data-k="미얀마">미얀마</button>
  </div>
  <div class="notice"><b>범례</b> &nbsp;|&nbsp; 연한 바 = 계획 &nbsp;|&nbsp; <span style="color:#185FA5">■</span> 진한 파랑 = 실적 &nbsp;<span style="color:#EF9F27">■</span> 주황 = 전망 &nbsp;|&nbsp; 달성률: <span style="color:#166534">●</span>100%↑ <span style="color:#854d0e">●</span>80~100% <span style="color:#991b1b">●</span>80%미만</div>
  <div class="krow" id="krow"></div>
  <div class="crow">
    <div class="cc">
      <div class="cct" id="t1">매출액 — 계획 vs 실적/전망</div>
      <div class="leg"><span><span class="lsq" style="background:#B5D4F4"></span>계획</span><span><span class="lsq" style="background:#185FA5"></span>실적</span><span><span class="lsq" style="background:#EF9F27"></span>전망</span></div>
      <div style="position:relative;height:200px"><canvas id="c1"></canvas></div>
    </div>
    <div class="cc">
      <div class="cct" id="t2">영업이익 — 계획 vs 실적/전망</div>
      <div class="leg"><span><span class="lsq" style="background:#9FE1CB"></span>계획</span><span><span class="lsq" style="background:#0F6E56"></span>실적</span><span><span class="lsq" style="background:#EF9F27"></span>전망</span></div>
      <div style="position:relative;height:200px"><canvas id="c2"></canvas></div>
    </div>
  </div>
  <div class="cc" id="ytd-sl" style="margin-bottom:12px;display:none">
    <div class="cct">법인별 매출액 YTD — 계획(연) vs 실적(진)</div>
    <div class="leg"><span><span class="lsq" style="background:#B5D4F4"></span>계획</span><span><span class="lsq" style="background:#185FA5"></span>실적</span></div>
    <div style="position:relative;height:220px"><canvas id="c3"></canvas></div>
  </div>
  <div class="cc" id="ytd-op" style="margin-bottom:12px;display:none">
    <div class="cct">법인별 영업이익 YTD — 계획(연) vs 실적(진)</div>
    <div class="leg"><span><span class="lsq" style="background:#9FE1CB"></span>계획</span><span><span class="lsq" style="background:#0F6E56"></span>실적</span></div>
    <div style="position:relative;height:220px"><canvas id="c4"></canvas></div>
  </div>
  <div class="tw"><table id="tbl"></table></div>
</div>
</div>
<script>
const MO=['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
const CORPS=['미국','멕시코','독일','이태리','일본','중국','싱가포르','인도','베트남','태국','말레이시아','인도네시아','우크라이나','미얀마'];
let D=null,cur='total';const ch={};
const fN=v=>v===null||v===undefined?'-':(v<0?'-':'')+Math.abs(Math.round(v)).toLocaleString('ko-KR');
const fP=v=>v===null?'-':(v*100).toFixed(1)+'%';
const aC=v=>v===null?'':v>=1?'#166534':v>=.8?'#854d0e':'#991b1b';
const aB=v=>v===null?'':v>=1?'#dcfce7':v>=.8?'#fef9c3':'#fee2e2';
const aCl=v=>v===null?'':v>=1?'g':v>=.8?'a':'r';
const ax=v=>{const a=Math.abs(v);if(a>=1e6)return(v<0?'-':'')+(a/1e6).toFixed(1)+'M';if(a>=1e3)return(v<0?'-':'')+(a/1e3).toFixed(0)+'K';return String(Math.round(v));};
const kill=id=>{if(ch[id]){ch[id].destroy();delete ch[id];}};
function fg(ma,mf,i){if(ma[i]!==null)return{v:ma[i],t:'act'};if(mf[i]!==null)return{v:mf[i],t:'fct'};return{v:null,t:null};}
function overlayChart(id,ma,mf,mp,pBg,pBd,aC2,fC){
  kill(id);const fgD=mp.map((_,i)=>fg(ma,mf,i));
  ch[id]=new Chart(document.getElementById(id),{type:'bar',data:{labels:MO,datasets:[
    {label:'계획',data:mp,backgroundColor:pBg,borderColor:pBd,borderWidth:1,borderRadius:4,order:2,barPercentage:0.7,categoryPercentage:0.8},
    {label:'실적/전망',data:fgD.map(f=>f.v),backgroundColor:fgD.map(f=>f.t==='act'?(f.v>=0?aC2+'ee':'#E24B4Aee'):f.t==='fct'?fC+'dd':'transparent'),borderColor:fgD.map(f=>f.t==='act'?aC2:f.t==='fct'?'#BA7517':'transparent'),borderWidth:1.5,borderRadius:4,order:1,barPercentage:0.42,categoryPercentage:0.8}
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{if(ctx.datasetIndex===0)return` 계획: ${fN(ctx.raw)}`;const f=fgD[ctx.dataIndex];return` ${f.t==='act'?'실적':'전망'}: ${fN(ctx.raw)}`;}}}},scales:{x:{ticks:{font:{size:10},maxRotation:0,autoSkip:false},grid:{display:false}},y:{ticks:{font:{size:10},callback:ax},grid:{color:'rgba(0,0,0,0.05)'}}}}});
}
function corpChart(id,planD,actualD,pBg,pBd,aC2){
  kill(id);
  ch[id]=new Chart(document.getElementById(id),{type:'bar',data:{labels:CORPS,datasets:[
    {label:'계획',data:planD,backgroundColor:pBg,borderColor:pBd,borderWidth:1,borderRadius:4,order:2,barPercentage:0.72,categoryPercentage:0.72},
    {label:'실적',data:actualD,backgroundColor:actualD.map(v=>v===null?'transparent':v>=0?aC2+'dd':'#E24B4Add'),borderColor:actualD.map(v=>v===null?'transparent':v>=0?aC2:'#E24B4A'),borderWidth:1.5,borderRadius:4,order:1,barPercentage:0.42,categoryPercentage:0.72}
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.dataset.label}: ${fN(ctx.raw)}`}}},scales:{x:{ticks:{font:{size:10},maxRotation:35,autoSkip:false},grid:{display:false}},y:{ticks:{font:{size:10},callback:ax},grid:{color:'rgba(0,0,0,0.05)'}}}}});
}
function render(){
  if(!D)return;
  const isT=cur==='total',d=D[cur],sl=d.sl,op=d.op,lbl=isT?'합계':cur;
  document.getElementById('t1').textContent=`매출액 — 계획 vs 실적/전망 (${lbl})`;
  document.getElementById('t2').textContent=`영업이익 — 계획 vs 실적/전망 (${lbl})`;
  document.getElementById('krow').innerHTML=[{l:'매출액 YTD 계획',v:fN(sl.yp),a:null},{l:'매출액 YTD 실적',v:fN(sl.ya),a:sl.yach},{l:'영업이익 YTD 계획',v:fN(op.yp),a:null},{l:'영업이익 YTD 실적',v:fN(op.ya),a:op.yach}].map(k=>`<div class="kpi"><div class="kl">${k.l}</div><div class="kv" style="${k.a!==null?'color:'+aC(k.a):''}">${k.v}</div>${k.a!==null?`<span class="ka" style="background:${aB(k.a)};color:${aC(k.a)}">${fP(k.a)}</span>`:''}</div>`).join('');
  overlayChart('c1',sl.ma,sl.mf,sl.mp,'#B5D4F455','#85B7EB','#185FA5','#EF9F27');
  overlayChart('c2',op.ma,op.mf,op.mp,'#9FE1CB44','#5DCAA5','#0F6E56','#EF9F27');
  const esl=document.getElementById('ytd-sl'),eop=document.getElementById('ytd-op');
  if(isT){esl.style.display='';eop.style.display='';corpChart('c3',CORPS.map(c=>D[c].sl.yp),CORPS.map(c=>D[c].sl.ya),'#B5D4F455','#85B7EB','#185FA5');corpChart('c4',CORPS.map(c=>D[c].op.yp),CORPS.map(c=>D[c].op.ya),'#9FE1CB44','#5DCAA5','#0F6E56');}
  else{esl.style.display='none';eop.style.display='none';kill('c3');kill('c4');}
  const tbl=document.getElementById('tbl');
  if(isT){
    const rows=CORPS.map(c=>{const s=D[c].sl,o=D[c].op;return`<tr><td>${c}</td><td>${fN(s.yp)}</td><td>${fN(s.ya)}</td><td><span class="ap ${aCl(s.yach)}">${fP(s.yach)}</span></td><td style="color:${(o.ya||0)<0?'#991b1b':''}">${fN(o.yp)}</td><td style="color:${(o.ya||0)<0?'#991b1b':''}">${fN(o.ya)}</td><td><span class="ap ${aCl(o.yach)}">${fP(o.yach)}</span></td></tr>`;});
    rows.push(`<tr class="tt"><td>합 계</td><td>${fN(D.total.sl.yp)}</td><td>${fN(D.total.sl.ya)}</td><td>${fP(D.total.sl.yach)}</td><td>${fN(D.total.op.yp)}</td><td>${fN(D.total.op.ya)}</td><td>${fP(D.total.op.yach)}</td></tr>`);
    tbl.innerHTML=`<thead><tr><th rowspan="2" style="text-align:left">법인</th><th colspan="3" style="background:#0C447C;text-align:center">매출액 YTD</th><th colspan="3" style="background:#085041;text-align:center">영업이익 YTD</th></tr><tr><th style="background:#185FA5">계획</th><th style="background:#185FA5">실적</th><th style="background:#185FA5">달성률</th><th style="background:#0F6E56">계획</th><th style="background:#0F6E56">실적</th><th style="background:#0F6E56">달성률</th></tr></thead><tbody>${rows.join('')}</tbody>`;
  } else {
    const mkR=(ind,d2)=>['계획','실적/전망','달성률'].map(sub=>{
      const cells=MO.map((_,i)=>{const f=fg(d2.ma,d2.mf,i);if(sub==='계획')return`<td style="color:#64748b">${fN(d2.mp[i])}</td>`;if(sub==='실적/전망'){if(f.t==='act')return`<td style="font-weight:600${f.v<0?';color:#991b1b':''}">${fN(f.v)}</td>`;if(f.t==='fct')return`<td style="color:#854F0B;font-style:italic">${fN(f.v)}</td>`;return`<td style="color:#cbd5e1">-</td>`;}if(f.t&&d2.mp[i]&&d2.mp[i]!==0){const r=f.v/d2.mp[i];return`<td><span class="ap ${aCl(r)}" style="${f.t==='fct'?'font-style:italic':''}">${fP(r)}</span></td>`;}return`<td style="color:#cbd5e1">-</td>`;}).join('');
      const yV=sub==='계획'?fN(d2.yp):sub==='실적/전망'?`<strong>${fN(d2.ya)}</strong>`:`<span class="ap ${aCl(d2.yach)}">${fP(d2.yach)}</span>`;
      return`<tr><td style="color:${ind==='매출액'?'#0C447C':'#085041'};font-weight:600">${ind}</td><td style="color:#64748b">${sub}</td>${cells}<td style="background:#EFF6FF;color:#1e40af;font-weight:600">${yV}</td></tr>`;
    }).join('');
    tbl.innerHTML=`<thead><tr><th style="text-align:left">지표</th><th style="text-align:left">구분</th>${MO.map(m=>`<th>${m}</th>`).join('')}<th style="background:#0C447C">YTD</th></tr></thead><tbody>${mkR('매출액',sl)}<tr><td colspan="${MO.length+3}" style="height:5px;background:#f8fafc;padding:0"></td></tr>${mkR('영업이익',op)}</tbody>`;
  }
}
document.getElementById('slbar').addEventListener('click',e=>{const b=e.target.closest('.sl');if(!b)return;cur=b.dataset.k;document.querySelectorAll('.sl').forEach(x=>x.classList.remove('on'));b.classList.add('on');render();});
const upzone=document.getElementById('upzone'),finput=document.getElementById('finput'),upstat=document.getElementById('upstat');
upzone.addEventListener('dragover',e=>{e.preventDefault();upzone.classList.add('drag');});
upzone.addEventListener('dragleave',()=>upzone.classList.remove('drag'));
upzone.addEventListener('drop',e=>{e.preventDefault();upzone.classList.remove('drag');if(e.dataTransfer.files[0])upload(e.dataTransfer.files[0]);});
finput.addEventListener('change',e=>{if(e.target.files[0])upload(e.target.files[0]);});
async function upload(file){
  upstat.textContent='업로드 중...';upstat.className='up-status';
  const fd=new FormData();fd.append('file',file);
  try{
    const res=await fetch('/upload',{method:'POST',body:fd});
    const json=await res.json();
    if(json.ok){D=json.data;cur='total';document.querySelectorAll('.sl').forEach(b=>b.classList.remove('on'));document.querySelector('.sl[data-k="total"]').classList.add('on');document.getElementById('dash').style.display='';document.getElementById('hdr-meta').innerHTML=`원본: ${file.name}<br>갱신: ${new Date().toLocaleString('ko-KR')}`;upstat.textContent=`완료: ${file.name}`;upstat.className='up-status ok';render();}
    else{upstat.textContent='오류: '+json.error;upstat.className='up-status err';}
  }catch(e){upstat.textContent='서버 오류: '+e.message;upstat.className='up-status err';}
}
</script>
</body>
</html>"""


@app.route('/')
def index():
    return render_template_string(PAGE)


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': '파일이 없습니다'})
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'ok': False, 'error': '.xlsx 파일만 지원합니다'})
    try:
        data = extract(io.BytesIO(f.read()))
        return jsonify({'ok': True, 'data': data})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"서버 시작: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
