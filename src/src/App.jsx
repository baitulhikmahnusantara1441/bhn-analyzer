import { useState, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import Papa from 'papaparse'

const NAVY = '#1B4F72'
const NAVYDARK = '#0d2d42'
const NAVYLT = '#2E6F9E'
const GOLD = '#C9A227'
const COLORS = ['#1B4F72','#C9A227','#2E86C1','#E67E22','#1ABC9C','#8E44AD','#E74C3C','#2ECC71']
const S = { bg:'#f0f4f8', card:'#ffffff', sec:'#f5f7fa', txt:'#1a1a2e', muted:'#64748b', border:'#e2e8f0' }

/* ── Helpers ── */
const detectType = vals => {
  const f = vals.filter(v => v !== '' && v !== null && v !== undefined)
  if (!f.length) return 'empty'
  const nums = f.filter(v => !isNaN(Number(v)) && String(v).trim() !== '')
  if (nums.length === f.length && f.length > 3) return 'numeric'
  const u = [...new Set(f)]
  if (u.length <= 12 || u.length / f.length < 0.35) return 'categorical'
  return 'text'
}
const freqMap = arr => {
  const m = {}
  arr.forEach(v => { const k = String(v ?? '').trim(); if (k) m[k] = (m[k] || 0) + 1 })
  return Object.entries(m).sort((a, b) => b[1] - a[1])
}
const pct = (n, t) => t ? Math.round(n / t * 100) : 0
const avg = nums => nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0

/* ── PPTX Builder ── */
async function buildPptx(meta, cols) {
  if (!window.JSZip) {
    await new Promise((res, rej) => {
      const s = document.createElement('script')
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
      s.onload = res; s.onerror = rej; document.head.appendChild(s)
    })
  }
  const zip = new window.JSZip()
  const date = new Date().toLocaleDateString('id-ID', { year:'numeric', month:'long', day:'numeric' })
  const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')
  const px = n => Math.round(n * 914400 / 2.54) | 0
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

  const sp = (id, x, y, w, h, fill, texts = []) => {
    const fXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>'
    const tXml = texts.length
      ? `<p:txBody><a:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/>${
          texts.map(t => `<a:p><a:pPr algn="${t.align||'l'}"/><a:r><a:rPr lang="id-ID" sz="${(t.size||12)*100}" b="${t.bold?1:0}" dirty="0"><a:solidFill><a:srgbClr val="${t.color||'333333'}"/></a:solidFill></a:rPr><a:t>${esc(t.text)}</a:t></a:r></a:p>`).join('')
        }</p:txBody>`
      : '<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>'
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fXml}<a:ln><a:noFill/></a:ln></p:spPr>${tXml}</p:sp>`
  }

  const wSlide = (shapes, bg = 'FFFFFF') =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClr/></p:clrMapOvr></p:sld>`

  const hdr = title =>
    sp(2,0,0,px(25.4),px(1.2),'1B4F72') +
    sp(3,0,px(1.2),px(25.4),px(0.06),'C9A227') +
    sp(4,0,px(18.6),px(25.4),px(0.45),'C9A227') +
    sp(5,px(0.6),px(0.15),px(22),px(0.9),'',[{text:title,size:20,color:'FFFFFF',bold:true}]) +
    sp(6,px(0.6),px(18.7),px(20),px(0.3),'',[{text:'Baitul Hikmah Nusantara — '+date,size:8,color:'FFFFFF'}])

  const cover = wSlide(
    sp(2,0,0,px(25.4),px(19.05),'1B4F72') +
    sp(3,0,0,px(25.4),px(3.5),'0d2d42') +
    sp(4,0,px(18.2),px(25.4),px(0.85),'C9A227') +
    sp(10,px(1),px(0.6),px(23),px(1.5),'',[{text:'بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيْمِ',size:22,color:'C9A227',align:'ctr',bold:true}]) +
    sp(11,px(0.5),px(2.5),px(24),px(1),'',[{text:'BAITUL HIKMAH NUSANTARA',size:13,color:'FFFFFF',align:'ctr',bold:true}]) +
    sp(12,px(1),px(5.5),px(23),px(3.5),'',[{text:meta.title,size:30,color:'FFFFFF',align:'ctr',bold:true}]) +
    sp(13,px(1),px(9.5),px(23),px(1.2),'',[{text:meta.subtitle,size:15,color:'C9A227',align:'ctr'}]) +
    sp(14,px(1),px(11.5),px(23),px(0.8),'',[{text:`Total: ${meta.total} responden  |  ${meta.cols} pertanyaan`,size:12,color:'AAAAAA',align:'ctr'}]) +
    sp(15,px(1),px(16.5),px(23),px(0.8),'',[{text:'Bogor, '+date,size:11,color:'888888',align:'ctr'}])
  , '1B4F72')

  const stats = [
    {l:'Total Responden',v:String(meta.total),c:'1B4F72',x:0.4,y:2.0},
    {l:'Jumlah Pertanyaan',v:String(meta.cols),c:'C9A227',x:13.0,y:2.0},
    {l:'Rata-rata Rating',v:meta.avgRating?meta.avgRating.toFixed(2)+'/5':'N/A',c:'1ABC9C',x:0.4,y:9.5},
    {l:'Tingkat Pengisian',v:(meta.fillRate||0)+'%',c:'E67E22',x:13.0,y:9.5},
  ]
  const statSlide = wSlide(
    hdr('Statistik Kunci') +
    stats.map((st,i) =>
      sp(10+i*3,px(st.x),px(st.y),px(11.5),px(6.0),st.c+'22') +
      sp(11+i*3,px(st.x+.15),px(st.y+.4),px(11),px(1.5),'',[{text:st.l,size:13,color:'555555',align:'ctr'}]) +
      sp(12+i*3,px(st.x+.15),px(st.y+2.2),px(11),px(3),'',[{text:st.v,size:40,color:st.c,align:'ctr',bold:true}])
    ).join('')
  )

  const colSlides = cols.filter(c => c.type==='categorical'||c.type==='numeric').slice(0,8).map(col => {
    let body = ''
    if (col.type === 'categorical') {
      body = col.freq.slice(0,6).map(([k,v],i) => {
        const bw = Math.max(0.8,(v/(col.filled.length||1))*19)
        const clr = COLORS[i%COLORS.length].replace('#','')
        return sp(20+i*3,px(0.5),px(2.0+i*2.5),px(bw),px(1.9),clr) +
          sp(21+i*3,px(0.6),px(2.0+i*2.5),px(bw-.2),px(1.9),'',[{text:k.slice(0,40),size:11,color:'FFFFFF'}]) +
          sp(22+i*3,px(bw+.7),px(2.0+i*2.5),px(4),px(1.9),'',[{text:`${v} (${pct(v,col.filled.length)}%)`,size:12,color:'333333'}])
      }).join('')
    } else {
      body = sp(20,px(.5),px(2),px(12),px(3),'',[{text:`Rata-rata: ${col.mean.toFixed(2)}`,size:36,color:'1B4F72',bold:true}]) +
        sp(21,px(.5),px(5.5),px(22),px(1.5),'',[{text:`Min: ${col.min}  |  Max: ${col.max}  |  Median: ${col.median.toFixed(1)}`,size:13,color:'555555'}])
    }
    return wSlide(hdr(`Analisis: ${col.name.slice(0,45)}`) +
      sp(10,px(.5),px(1.3),px(22),px(.6),'',[{text:`Tipe: ${col.type}  |  Terisi: ${col.filled.length} dari ${col.total}`,size:10,color:'888888'}]) + body)
  })

  const closing = wSlide(
    sp(2,0,0,px(25.4),px(19.05),'0d2d42') +
    sp(3,0,px(17.8),px(25.4),px(1.25),'C9A227') +
    sp(10,px(1),px(4.5),px(23.4),px(2),'',[{text:'وَمَا تَوْفِيقِي إِلَّا بِاللهِ',size:26,color:'C9A227',align:'ctr',bold:true}]) +
    sp(11,px(1),px(6.8),px(23.4),px(1),'',[{text:'"Dan tidak ada taufik bagiku melainkan dengan pertolongan Allah"',size:12,color:'AAAAAA',align:'ctr'}]) +
    sp(12,px(1),px(9),px(23.4),px(2.5),'',[{text:'Terima Kasih',size:40,color:'FFFFFF',align:'ctr',bold:true}]) +
    sp(13,px(1),px(12),px(23.4),px(1),'',[{text:'Baitul Hikmah Nusantara — Bogor',size:15,color:'C9A227',align:'ctr'}])
  ,'0d2d42')

  const allSlides = [cover, statSlide, ...colSlides, closing]
  const slideRel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`

  zip.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>${allSlides.map((_,i)=>`<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`)
  zip.file('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`)
  zip.file('ppt/_rels/presentation.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>${allSlides.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join('')}</Relationships>`)
  zip.file('ppt/presentation.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId0"/></p:sldMasterIdLst><p:sldIdLst>${allSlides.map((_,i)=>`<p:sldId id="${256+i}" r:id="rId${i+1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`)
  zip.file('ppt/slideLayouts/slideLayout1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${NS} type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClr/></p:clrMapOvr></p:sldLayout>`)
  allSlides.forEach((xml,i) => { zip.file(`ppt/slides/slide${i+1}.xml`,xml); zip.file(`ppt/slides/_rels/slide${i+1}.xml.rels`,slideRel) })

  const blob = await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob)
  a.download = `BHN_${(meta.title||'Laporan').replace(/[^a-zA-Z0-9]/g,'_').slice(0,30)}.pptx`; a.click()
}

/* ── Components ── */
function StatCard({ icon, label, value, accent }) {
  return (
    <div style={{background:S.card,border:`1.5px solid ${accent}33`,borderRadius:14,padding:'18px 20px',flex:'1 1 120px',position:'relative',overflow:'hidden'}}>
      <div style={{position:'absolute',right:-10,top:-10,width:60,height:60,borderRadius:'50%',background:accent+'14'}}/>
      <div style={{fontSize:22,marginBottom:6}}>{icon}</div>
      <div style={{fontSize:10,color:S.muted,marginBottom:3,textTransform:'uppercase',letterSpacing:.8,fontWeight:600}}>{label}</div>
      <div style={{fontSize:24,fontWeight:700,color:accent}}>{value}</div>
    </div>
  )
}

function MiniBar({ data, total }) {
  const top = data.slice(0,7), max = top[0]?.[1]||1
  return (
    <div style={{marginTop:10}}>
      {top.map(([k,v],i) => (
        <div key={i} style={{marginBottom:8}}>
          <div style={{display:'flex',justifyContent:'space-between',fontSize:11.5,marginBottom:3}}>
            <span style={{maxWidth:190,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',color:S.txt}}>{k}</span>
            <span style={{fontWeight:700,color:COLORS[i%8],marginLeft:8}}>{v}<span style={{color:S.muted,fontWeight:400}}> ({pct(v,total)}%)</span></span>
          </div>
          <div style={{height:7,background:S.sec,borderRadius:6,overflow:'hidden'}}>
            <div style={{height:'100%',width:`${pct(v,max)}%`,background:COLORS[i%8],borderRadius:6}}/>
          </div>
        </div>
      ))}
    </div>
  )
}

/* ── App ── */
export default function App() {
  const [page, setPage] = useState('landing')
  const [files, setFiles] = useState([])
  const [activeFile, setActiveFile] = useState(0)
  const [drag, setDrag] = useState(false)
  const [error, setError] = useState('')
  const [sheetUrl, setSheetUrl] = useState('')
  const [loadingSheet, setLoadingSheet] = useState(false)
  const [tab, setTab] = useState('dashboard')
  const [pptxLoading, setPptxLoading] = useState(false)

  const processRows = (rows, name) => {
    if (!rows || rows.length < 2) { setError('Data terlalu sedikit.'); return }
    const h = rows[0].map(String)
    const body = rows.slice(1).filter(r => r.some(c => c !== '' && c !== null))
    setFiles(prev => [...prev, { name, headers: h, data: body }])
    setError(''); setPage('app'); setTab('dashboard')
  }

  const handleFile = useCallback(file => {
    const name = file.name
    if (name.toLowerCase().endsWith('.csv')) {
      Papa.parse(file, { complete: r => processRows(r.data, name), skipEmptyLines: true })
    } else if (name.toLowerCase().match(/\.xlsx?$/)) {
      const reader = new FileReader()
      reader.onload = e => {
        const wb = XLSX.read(e.target.result, { type: 'binary' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        processRows(XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }), name)
      }
      reader.readAsBinaryString(file)
    } else setError('Format didukung: .xlsx, .xls, .csv')
  }, [])

  const handleDrop = useCallback(e => {
    e.preventDefault(); setDrag(false)
    Array.from(e.dataTransfer.files).forEach(f => handleFile(f))
  }, [handleFile])

  const loadGSheet = async () => {
    setLoadingSheet(true); setError('')
    try {
      const id = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)?.[1]
      if (!id) { setError('URL tidak valid.'); setLoadingSheet(false); return }
      const res = await fetch(`https://docs.google.com/spreadsheets/d/${id}/export?format=csv`)
      if (!res.ok) throw new Error('Gagal. Pastikan Sheets sudah diset publik.')
      Papa.parse(await res.text(), { complete: r => processRows(r.data, 'Google Sheets'), skipEmptyLines: true })
    } catch (e) { setError(e.message) }
    setLoadingSheet(false)
  }

  const fileData = useMemo(() => {
    if (!files.length) return null
    const f = files[activeFile]; if (!f) return null
    const cols = f.headers.map((h, i) => {
      const vals = f.data.map(r => String(r[i] ?? '').trim())
      const filled = vals.filter(v => v !== '')
      const type = detectType(filled), freq = freqMap(filled)
      let mean = 0, min = 0, max = 0, median = 0
      if (type === 'numeric') {
        const nums = filled.map(Number).filter(n => !isNaN(n))
        mean = avg(nums); min = Math.min(...nums); max = Math.max(...nums)
        const s = [...nums].sort((a,b) => a-b); median = s[Math.floor(s.length/2)]
      }
      return { name: h, vals, filled, type, freq, mean, min, max, median, total: f.data.length }
    })
    const rCols = cols.filter(c => c.type === 'numeric' && c.mean >= 1 && c.mean <= 5)
    return { ...f, cols, avgRating: rCols.length ? avg(rCols.map(c => c.mean)) : null, fillRate: Math.round(cols.reduce((a,c) => a+c.filled.length/c.total,0)/cols.length*100) }
  }, [files, activeFile])

  const exportPptx = async () => {
    if (!fileData) return; setPptxLoading(true)
    try {
      await buildPptx({
        title: `Laporan — ${fileData.name}`,
        subtitle: new Date().toLocaleDateString('id-ID',{year:'numeric',month:'long'}),
        total: fileData.data?.length||0, cols: fileData.headers?.length||0,
        avgRating: fileData.avgRating, fillRate: fileData.fillRate||0
      }, fileData.cols||[])
    } catch(e) { setError('Gagal export: '+e.message) }
    setPptxLoading(false)
  }

  /* ── Landing ── */
  if (page === 'landing') return (
    <div style={{minHeight:'100vh',background:S.bg,display:'flex',alignItems:'center',justifyContent:'center',padding:'2rem 1rem'}}>
      <div style={{width:'100%',maxWidth:860}}>
        <div style={{background:`linear-gradient(135deg,${NAVYDARK},${NAVY} 55%,${NAVYLT})`,borderRadius:22,padding:'3rem 2.5rem 2.5rem',marginBottom:18,position:'relative',overflow:'hidden'}}>
          <div style={{position:'absolute',top:-60,right:-60,width:250,height:250,borderRadius:'50%',background:GOLD+'14'}}/>
          <div style={{position:'absolute',bottom:-40,left:-40,width:180,height:180,borderRadius:'50%',background:GOLD+'0C'}}/>
          <div style={{position:'relative'}}>
            <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:24}}>
              <div style={{background:GOLD,borderRadius:13,width:50,height:50,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:16,color:NAVY,flexShrink:0}}>BHN</div>
              <div>
                <div style={{fontSize:11,color:GOLD+'BB',letterSpacing:2.5,textTransform:'uppercase',fontWeight:600}}>Baitul Hikmah Nusantara</div>
                <div style={{fontSize:12,color:'#ffffff60'}}>Lembaga Pendidikan Islam — Bogor</div>
              </div>
            </div>
            <h1 style={{fontSize:32,fontWeight:700,color:'#fff',margin:'0 0 12px',lineHeight:1.25}}>Survey & Data<br/><span style={{color:GOLD}}>Analyzer Dashboard</span></h1>
            <p style={{fontSize:14,color:'#ffffffa0',margin:'0 0 28px',lineHeight:1.8,maxWidth:480}}>Upload data Google Form, analisis visual otomatis, dan download presentasi PowerPoint berbranding BHN — tanpa login, tanpa konfigurasi.</p>
            <div style={{display:'flex',gap:12,alignItems:'center',flexWrap:'wrap'}}>
              <button onClick={()=>setPage('app')} style={{background:GOLD,color:NAVY,border:'none',borderRadius:11,padding:'12px 28px',fontSize:14,fontWeight:700,cursor:'pointer'}}>Mulai Sekarang →</button>
              <span style={{fontSize:12,color:'#ffffff50',display:'flex',alignItems:'center',gap:5}}>🔒 Data tidak dikirim ke server</span>
            </div>
          </div>
        </div>

        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:12,marginBottom:18}}>
          {[
            {icon:'📁',title:'Upload Fleksibel',desc:'Excel, CSV, atau Google Sheets',color:NAVY},
            {icon:'📊',title:'Dashboard Visual',desc:'Grafik & statistik otomatis per kolom',color:NAVYLT},
            {icon:'🔍',title:'Data Mentah',desc:'Lihat dan telusuri semua responden',color:'#1ABC9C'},
            {icon:'📑',title:'Export PowerPoint',desc:'Slide BHN navy+gold langsung diedit',color:GOLD},
          ].map((f,i) => (
            <div key={i} style={{background:S.card,border:`0.5px solid ${S.border}`,borderRadius:14,padding:'18px 16px',boxShadow:'0 1px 4px #0000000A'}}>
              <div style={{fontSize:28,marginBottom:10}}>{f.icon}</div>
              <div style={{fontSize:13,fontWeight:700,marginBottom:4,color:S.txt}}>{f.title}</div>
              <div style={{fontSize:12,color:S.muted,lineHeight:1.55}}>{f.desc}</div>
            </div>
          ))}
        </div>

        <div style={{background:S.card,borderRadius:16,padding:'1.5rem',border:`0.5px solid ${S.border}`}}>
          <p style={{fontSize:13,color:S.muted,margin:'0 0 12px',textAlign:'center',fontWeight:500}}>Atau langsung drop file di sini</p>
          <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
            onDrop={e=>{handleDrop(e);setPage('app');}} onClick={()=>setPage('app')}
            style={{border:`2px dashed ${drag?GOLD:NAVY+'44'}`,borderRadius:12,padding:'1.5rem',textAlign:'center',background:drag?GOLD+'08':S.sec,cursor:'pointer',transition:'all .2s'}}>
            <div style={{fontSize:30,marginBottom:8}}>☁️</div>
            <p style={{margin:0,fontSize:13,color:S.muted}}>Drag & drop Excel/CSV, atau <span style={{color:NAVY,fontWeight:700}}>klik untuk pilih file</span></p>
          </div>
        </div>
      </div>
    </div>
  )

  /* ── App Page ── */
  return (
    <div style={{minHeight:'100vh',background:S.bg}}>
      <div style={{background:NAVY,padding:'0 1.5rem',display:'flex',alignItems:'center',height:52,gap:12,boxShadow:'0 2px 8px #0003'}}>
        <div style={{background:GOLD,borderRadius:8,width:36,height:36,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:13,color:NAVY,flexShrink:0}}>BHN</div>
        <span style={{color:'#fff',fontWeight:600,fontSize:14}}>Analyzer Dashboard</span>
        <div style={{flex:1}}/>
        {files.length>0 && (
          <button onClick={exportPptx} disabled={pptxLoading}
            style={{background:GOLD,color:NAVY,border:'none',borderRadius:9,padding:'7px 16px',fontSize:12,cursor:'pointer',fontWeight:700,display:'flex',alignItems:'center',gap:6}}>
            📑 {pptxLoading?'Membuat…':'Download PPTX'}
          </button>
        )}
        <button onClick={()=>{setPage('landing');setFiles([]);setSheetUrl('');}}
          style={{background:'transparent',border:'0.5px solid #ffffff44',color:'#fff',borderRadius:8,padding:'6px 12px',fontSize:12,cursor:'pointer'}}>
          🏠 Beranda
        </button>
      </div>

      <div style={{maxWidth:960,margin:'0 auto',padding:'1.25rem 1rem'}}>
        {error && (
          <div style={{background:'#fdecea',border:'0.5px solid #e74c3c55',borderRadius:10,padding:'10px 14px',fontSize:12.5,color:'#c0392b',marginBottom:12}}>
            {error}<button onClick={()=>setError('')} style={{float:'right',background:'none',border:'none',cursor:'pointer',color:'#c0392b',fontWeight:700}}>✕</button>
          </div>
        )}

        {!files.length && (
          <div style={{background:S.card,borderRadius:16,padding:'2rem',border:`0.5px solid ${S.border}`}}>
            <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop}
              style={{border:`2px dashed ${drag?GOLD:NAVY+'44'}`,borderRadius:12,padding:'2.5rem 1.5rem',textAlign:'center',background:drag?GOLD+'08':S.sec,transition:'all .2s',marginBottom:16}}>
              <div style={{fontSize:40,marginBottom:12}}>📂</div>
              <p style={{margin:'0 0 4px',fontSize:15,fontWeight:700,color:S.txt}}>Upload File Data</p>
              <p style={{margin:'0 0 16px',fontSize:12,color:S.muted}}>Excel (.xlsx/.xls) atau CSV</p>
              <label style={{background:NAVY,color:'#fff',borderRadius:10,padding:'10px 24px',fontSize:13,cursor:'pointer',display:'inline-block',fontWeight:600}}>
                📁 Pilih File
                <input type="file" accept=".xlsx,.xls,.csv" multiple style={{display:'none'}} onChange={e=>Array.from(e.target.files).forEach(f=>handleFile(f))}/>
              </label>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:14}}>
              <div style={{flex:1,height:1,background:S.border}}/><span style={{fontSize:12,color:S.muted,fontWeight:500}}>atau Google Sheets</span><div style={{flex:1,height:1,background:S.border}}/>
            </div>
            <div style={{display:'flex',gap:8}}>
              <input placeholder="Paste link Google Sheets publik…" value={sheetUrl} onChange={e=>setSheetUrl(e.target.value)}
                style={{flex:1,borderRadius:10,padding:'9px 14px',fontSize:13,border:`0.5px solid ${S.border}`,outline:'none',background:S.sec,color:S.txt}}/>
              <button onClick={loadGSheet} disabled={loadingSheet||!sheetUrl}
                style={{background:NAVY,color:'#fff',border:'none',borderRadius:10,padding:'9px 18px',fontSize:13,cursor:'pointer',fontWeight:600,whiteSpace:'nowrap'}}>
                {loadingSheet?'Memuat…':'Muat'}
              </button>
            </div>
          </div>
        )}

        {files.length>0 && (
          <div>
            <div style={{display:'flex',gap:6,flexWrap:'wrap',alignItems:'center',marginBottom:12}}>
              {files.map((f,i) => (
                <div key={i} onClick={()=>setActiveFile(i)}
                  style={{display:'flex',alignItems:'center',gap:5,padding:'6px 14px',borderRadius:20,cursor:'pointer',fontSize:12,fontWeight:activeFile===i?700:400,background:activeFile===i?NAVY:S.card,color:activeFile===i?'#fff':S.muted,border:`1px solid ${activeFile===i?NAVY:S.border}`,transition:'all .2s'}}>
                  📊 {f.name.length>22?f.name.slice(0,22)+'…':f.name}
                  <span onClick={e=>{e.stopPropagation();setFiles(p=>p.filter((_,j)=>j!==i));if(activeFile>=i)setActiveFile(0);}} style={{marginLeft:3,opacity:.5,fontSize:11,cursor:'pointer'}}>✕</span>
                </div>
              ))}
              <label style={{display:'flex',alignItems:'center',gap:4,padding:'6px 14px',borderRadius:20,cursor:'pointer',fontSize:12,border:`1px dashed ${GOLD}`,color:GOLD,background:GOLD+'0A',fontWeight:600}}>
                ➕ Tambah
                <input type="file" accept=".xlsx,.xls,.csv" multiple style={{display:'none'}} onChange={e=>Array.from(e.target.files).forEach(f=>handleFile(f))}/>
              </label>
            </div>

            <div style={{display:'flex',gap:2,marginBottom:16,background:S.card,borderRadius:11,padding:4,border:`0.5px solid ${S.border}`}}>
              {[['dashboard','📊','Dashboard'],['data','📋','Data Mentah']].map(([id,icon,label]) => (
                <button key={id} onClick={()=>setTab(id)}
                  style={{flex:1,background:tab===id?NAVY:'transparent',color:tab===id?'#fff':S.muted,border:'none',borderRadius:8,padding:'8px',fontSize:12.5,cursor:'pointer',fontWeight:tab===id?700:400,display:'flex',alignItems:'center',justifyContent:'center',gap:5,transition:'all .2s'}}>
                  {icon} {label}
                </button>
              ))}
            </div>

            {tab==='dashboard' && fileData && (
              <div>
                <div style={{display:'flex',gap:10,flexWrap:'wrap',marginBottom:16}}>
                  <StatCard icon="👥" label="Total Responden"   value={fileData.data?.length||0}        accent={NAVY}/>
                  <StatCard icon="📋" label="Pertanyaan"        value={fileData.headers?.length||0}     accent={GOLD}/>
                  <StatCard icon="✅" label="Tingkat Pengisian" value={(fileData.fillRate||0)+'%'}      accent="#27AE60"/>
                  {fileData.avgRating && <StatCard icon="⭐" label="Avg Rating" value={fileData.avgRating.toFixed(2)+'/5'} accent="#E67E22"/>}
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))',gap:12}}>
                  {fileData.cols?.filter(c=>c.type!=='empty').map((col,i) => (
                    <div key={i} style={{background:S.card,border:`0.5px solid ${S.border}`,borderRadius:14,padding:16,boxShadow:'0 1px 4px #0000000A'}}>
                      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8}}>
                        <span style={{fontSize:13,fontWeight:700,flex:1,lineHeight:1.35,color:S.txt}}>{col.name}</span>
                        <span style={{fontSize:10,padding:'2px 8px',borderRadius:20,background:col.type==='categorical'?NAVY+'18':col.type==='numeric'?GOLD+'22':'#f0f0f0',color:col.type==='categorical'?NAVY:col.type==='numeric'?'#8a6200':'#888',marginLeft:8,whiteSpace:'nowrap',fontWeight:600}}>{col.type}</span>
                      </div>
                      <div style={{fontSize:10.5,color:S.muted,marginBottom:10,display:'flex',alignItems:'center',gap:6}}>
                        {col.filled.length}/{col.total} terisi
                        <span style={{display:'inline-block',height:5,flex:1,borderRadius:4,background:S.sec,position:'relative',maxWidth:80}}>
                          <span style={{position:'absolute',left:0,top:0,height:5,width:`${pct(col.filled.length,col.total)}%`,borderRadius:4,background:NAVY}}/>
                        </span>
                        <span style={{fontWeight:700,color:NAVY}}>{pct(col.filled.length,col.total)}%</span>
                      </div>
                      {col.type==='categorical' && <MiniBar data={col.freq} total={col.filled.length}/>}
                      {col.type==='numeric' && (
                        <div>
                          <div style={{fontSize:28,fontWeight:700,color:NAVY,marginBottom:4}}>{col.mean.toFixed(2)}<span style={{fontSize:12,color:S.muted,marginLeft:6,fontWeight:400}}>rata-rata</span></div>
                          <div style={{display:'flex',gap:16,fontSize:11.5,color:S.muted}}>
                            <span>Min <strong style={{color:S.txt}}>{col.min}</strong></span>
                            <span>Max <strong style={{color:S.txt}}>{col.max}</strong></span>
                            <span>Median <strong style={{color:S.txt}}>{col.median.toFixed(1)}</strong></span>
                          </div>
                        </div>
                      )}
                      {col.type==='text' && <div style={{fontSize:11.5,color:S.muted,fontStyle:'italic',lineHeight:1.6}}>"{(col.vals.find(v=>v.length>10)||col.vals[0]||'').slice(0,80)}…"<div style={{marginTop:5,fontStyle:'normal',fontSize:11}}>{col.filled.length} respons teks</div></div>}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {tab==='data' && fileData && (
              <div style={{background:S.card,borderRadius:14,border:`0.5px solid ${S.border}`,overflow:'hidden'}}>
                <div style={{padding:'12px 16px',borderBottom:`0.5px solid ${S.border}`,fontSize:12,color:S.muted,display:'flex',justifyContent:'space-between'}}>
                  <span>50 baris pertama dari <strong style={{color:S.txt}}>{fileData.data?.length||0}</strong> entri</span>
                  <span>{fileData.headers?.length} kolom</span>
                </div>
                <div style={{overflowX:'auto'}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:11.5,tableLayout:'fixed'}}>
                    <thead>
                      <tr style={{background:NAVY}}>
                        <th style={{width:30,padding:'8px 6px',color:GOLD,textAlign:'center',fontWeight:700,fontSize:10}}>#</th>
                        {fileData.headers?.map((h,i) => <th key={i} style={{padding:'8px 10px',color:'#fff',textAlign:'left',fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:130,fontSize:11}}>{h}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {fileData.data?.slice(0,50).map((row,ri) => (
                        <tr key={ri} style={{background:ri%2===0?'#fff':S.sec}}>
                          <td style={{padding:'6px',textAlign:'center',color:S.muted,fontSize:10}}>{ri+1}</td>
                          {fileData.headers?.map((_,ci) => (
                            <td key={ci} title={String(row[ci]??'')}
                              style={{padding:'6px 10px',borderBottom:`0.5px solid ${S.border}`,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:130,color:S.txt}}>
                              {String(row[ci]??'')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
