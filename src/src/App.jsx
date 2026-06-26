import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

/* ─── CONSTANTS ─────────────────────────────────────────────────────────── */
const NAVY     = "#1B4F72";
const NAVYDARK = "#0d2d42";
const NAVYLT   = "#2E6F9E";
const GOLD     = "#C9A227";
const COLORS   = ["#1B4F72","#C9A227","#2E86C1","#E67E22","#1ABC9C","#8E44AD","#E74C3C","#2ECC71"];

/* ─── PPTX BUILDER (JSZip) ──────────────────────────────────────────────── */
async function buildPptx(meta, summary, colAnalyses, aiReport) {
  // Lazy-load JSZip from CDN
  if (!window.JSZip) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  const zip  = new window.JSZip();
  const date = new Date().toLocaleDateString("id-ID", { year:"numeric", month:"long", day:"numeric" });
  const esc  = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  const px   = n => Math.round(n * 914400 / 2.54 * 100) / 100 | 0;  // cm → EMU
  const pct  = (n, t) => t ? Math.round(n / t * 100) : 0;

  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
             'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  /* shape helper */
  const sp = (id, x, y, w, h, fill, texts = []) => {
    const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : `<a:noFill/>`;
    const txXml = texts.length
      ? `<p:txBody><a:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/>${
          texts.map(t => `<a:p><a:pPr algn="${t.align||"l"}"/><a:r>
<a:rPr lang="id-ID" sz="${(t.size||12)*100}" b="${t.bold?1:0}" i="${t.italic?1:0}" dirty="0">
<a:solidFill><a:srgbClr val="${t.color||"333333"}"/></a:solidFill></a:rPr>
<a:t>${esc(t.text)}</a:t></a:r></a:p>`).join("")
        }</p:txBody>`
      : `<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>`;
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/>
<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml}<a:ln><a:noFill/></a:ln></p:spPr>${txXml}</p:sp>`;
  };

  const wrapSlide = (shapes, bg = "FFFFFF") =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:bg><p:bgPr>
<a:solidFill><a:srgbClr val="${bg}"/></a:solidFill><a:effectLst/>
</p:bgPr></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClr/></p:clrMapOvr></p:sld>`;

  /* slide header */
  const hdr = title =>
    sp(2,  0, 0, px(25.4), px(1.2), "1B4F72") +
    sp(3,  0, px(1.2), px(25.4), px(0.06), "C9A227") +
    sp(4,  0, px(18.6), px(25.4), px(0.45), "C9A227") +
    sp(5,  px(0.6), px(0.15), px(22), px(0.9), "", [{ text: title, size: 20, color: "FFFFFF", bold: true }]) +
    sp(6,  px(0.6), px(18.7), px(20), px(0.3), "", [{ text: "Baitul Hikmah Nusantara  —  " + date, size: 8, color: "FFFFFF" }]);

  /* ── SLIDE 1: Cover ── */
  const s1 = wrapSlide(
    sp(2,  0, 0, px(25.4), px(19.05), "1B4F72") +
    sp(3,  0, 0, px(25.4), px(3.5),   "0d2d42") +
    sp(4,  0, px(18.2), px(25.4), px(0.85), "C9A227") +
    sp(10, px(1), px(0.6), px(23), px(1.5), "", [{ text:"بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيْمِ", size:22, color:"C9A227", align:"ctr", bold:true }]) +
    sp(11, px(0.5), px(2.5), px(24), px(1),  "", [{ text:"BAITUL HIKMAH NUSANTARA", size:13, color:"FFFFFF", align:"ctr", bold:true }]) +
    sp(12, px(1), px(5.5), px(23), px(3.5),  "", [{ text: meta.title||"Laporan Analisis", size:30, color:"FFFFFF", align:"ctr", bold:true }]) +
    sp(13, px(1), px(9.5), px(23), px(1.2),  "", [{ text: meta.subtitle||date, size:15, color:"C9A227", align:"ctr" }]) +
    sp(14, px(1), px(11.5),px(23), px(0.8),  "", [{ text:`Total Data: ${meta.total} entri   |   Pertanyaan: ${meta.cols}`, size:12, color:"AAAAAA", align:"ctr" }]) +
    sp(15, px(1), px(16.5),px(23), px(0.8),  "", [{ text:"Bogor, "+date, size:11, color:"888888", align:"ctr" }])
  , "1B4F72");

  /* ── SLIDE 2: Ringkasan Eksekutif ── */
  const s2 = wrapSlide(
    hdr("Ringkasan Eksekutif") +
    sp(10, px(0.5), px(1.5), px(24), px(16.5), "", [{ text: summary.slice(0,1400), size:13, color:"222222" }])
  );

  /* ── SLIDE 3: Statistik ── */
  const stats = [
    { l:"Total Responden",    v: String(meta.total),                           c:"1B4F72", x:0.4,  y:2.0 },
    { l:"Jumlah Pertanyaan",  v: String(meta.cols),                            c:"C9A227", x:13.0, y:2.0 },
    { l:"Rata-rata Rating",   v: meta.avgRating ? meta.avgRating.toFixed(2)+"/5":"N/A", c:"1ABC9C", x:0.4,  y:9.5 },
    { l:"Tingkat Pengisian",  v: (meta.fillRate||0)+"%",                       c:"E67E22", x:13.0, y:9.5 },
  ];
  const s3 = wrapSlide(
    hdr("Statistik Kunci") +
    stats.map((st, i) =>
      sp(10+i*3, px(st.x), px(st.y), px(11.5), px(6.0), st.c+"18") +
      sp(11+i*3, px(st.x+.15), px(st.y+.4),  px(11), px(1.5), "", [{ text:st.l, size:13, color:"555555", align:"ctr" }]) +
      sp(12+i*3, px(st.x+.15), px(st.y+2.2), px(11), px(3),   "", [{ text:st.v, size:40, color:st.c,     align:"ctr", bold:true }])
    ).join("")
  );

  /* ── Slides per kolom (max 6) ── */
  const catCols = colAnalyses.filter(c => c.type==="categorical"||c.type==="numeric").slice(0,6);
  const colSlides = catCols.map(col => {
    let body = "";
    if (col.type === "categorical") {
      body = col.freq.slice(0,6).map(([k,v], i) => {
        const bw = Math.max(0.8, (v/(col.filled.length||1)) * 19);
        const clr = COLORS[i%COLORS.length].replace("#","");
        return sp(20+i*3, px(0.5),    px(2.0+i*2.5), px(bw),    px(1.9), clr) +
               sp(21+i*3, px(0.6),    px(2.0+i*2.5), px(bw-.2), px(1.9), "", [{ text:k.slice(0,40), size:11, color:"FFFFFF" }]) +
               sp(22+i*3, px(bw+0.7), px(2.0+i*2.5), px(4),     px(1.9), "", [{ text:`${v} (${pct(v,col.filled.length)}%)`, size:12, color:"333333" }]);
      }).join("");
    } else {
      body =
        sp(20, px(.5), px(1.8), px(12), px(3),   "", [{ text:`Rata-rata: ${col.mean.toFixed(2)}`, size:36, color:"1B4F72", bold:true }]) +
        sp(21, px(.5), px(5.2), px(22), px(1.2), "", [{ text:`Min: ${col.min}   |   Max: ${col.max}   |   Median: ${col.median.toFixed(1)}`, size:13, color:"555555" }]);
    }
    const ins = col.insight
      ? sp(50, px(.5), px(15.8), px(24.4), px(2.5), "EAF2FB") +
        sp(51, px(.7), px(15.9), px(24),   px(.5),  "", [{ text:"Insight AI:", size:9, color:"1B4F72", bold:true }]) +
        sp(52, px(.7), px(16.5), px(23.5), px(1.7), "", [{ text:(col.insight||"").slice(0,280), size:10, color:"1a3a52" }])
      : "";
    return wrapSlide(
      hdr(`Analisis: ${col.name.slice(0,45)}`) +
      sp(10, px(.5), px(1.3), px(22), px(.6), "", [{ text:`Tipe: ${col.type}   |   Terisi: ${col.filled.length} dari ${col.total}`, size:10, color:"888888" }]) +
      body + ins
    );
  });

  /* ── Analisis Mendalam slides ── */
  const chunks = [];
  for (let i = 0; i < aiReport.length; i += 950) chunks.push(aiReport.slice(i, i+950));
  const deepSlides = chunks.slice(0,4).map((chunk, idx) => wrapSlide(
    hdr(idx===0 ? "Analisis Mendalam" : "Analisis Mendalam (Lanjutan)") +
    sp(10, px(.5), px(1.5), px(24), px(16.5), "", [{ text:chunk, size:12, color:"222222" }])
  ));

  /* ── Rekomendasi ── */
  const recSlide = wrapSlide(
    hdr("Rekomendasi & Rencana Tindak Lanjut") +
    sp(10, px(.4), px(1.4), px(24.6), px(16.8), "F7F9FB") +
    sp(11, px(.5), px(1.6), px(23.5), px(.9),   "", [{ text:"Berdasarkan analisis komprehensif data BHN:", size:12, color:"1B4F72", bold:true }]) +
    sp(12, px(.6), px(2.7), px(23),   px(15),   "", [{ text:aiReport.slice(-850), size:11, color:"333333" }])
  );

  /* ── Penutup ── */
  const closeSlide = wrapSlide(
    sp(2,  0, 0,         px(25.4), px(19.05), "0d2d42") +
    sp(3,  0, px(17.8),  px(25.4), px(1.25),  "C9A227") +
    sp(10, px(1), px(4.5), px(23.4), px(2),   "", [{ text:"وَمَا تَوْفِيقِي إِلَّا بِاللهِ", size:26, color:"C9A227", align:"ctr", bold:true }]) +
    sp(11, px(1), px(6.8), px(23.4), px(1),   "", [{ text:'"Dan tidak ada taufik bagiku melainkan dengan pertolongan Allah"', size:12, color:"AAAAAA", align:"ctr", italic:true }]) +
    sp(12, px(1), px(9),   px(23.4), px(2.5), "", [{ text:"Terima Kasih", size:40, color:"FFFFFF", align:"ctr", bold:true }]) +
    sp(13, px(1), px(12),  px(23.4), px(1),   "", [{ text:"Baitul Hikmah Nusantara — Bogor", size:15, color:"C9A227", align:"ctr" }])
  , "0d2d42");

  /* ── Assemble ZIP ── */
  const allSlides = [s1, s2, s3, ...colSlides, ...deepSlides, recSlide, closeSlide];

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml"  ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml"              ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
${allSlides.map((_,i)=>`<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("\n")}
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${allSlides.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join("\n")}
</Relationships>`);

  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS} saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId0"/></p:sldMasterIdLst>
<p:sldIdLst>${allSlides.map((_,i)=>`<p:sldId id="${256+i}" r:id="rId${i+1}"/>`).join("")}</p:sldIdLst>
<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`);

  const masterRel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", masterRel);
  zip.file("ppt/slideMasters/slideMaster1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`);

  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  zip.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS} type="blank"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClr/></p:clrMapOvr></p:sldLayout>`);

  allSlides.forEach((xml, i) => {
    zip.file(`ppt/slides/slide${i+1}.xml`, xml);
    zip.file(`ppt/slides/_rels/slide${i+1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
  });

  const blob = await zip.generateAsync({ type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `BHN_Analisis_${(meta.title||"Laporan").replace(/[^a-zA-Z0-9]/g,"_").slice(0,40)}.pptx`;
  a.click();
}

/* ─── AI CALL ───────────────────────────────────────────────────────────── */
async function callClaude(prompt, onChunk) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6", max_tokens: 1000, stream: true,
      messages: [{ role:"user", content: prompt }],
      system: `Kamu adalah analis data senior BHN (Baitul Hikmah Nusantara), lembaga pendidikan Islam di Bogor.
Tulis dalam Bahasa Indonesia yang tajam, akademis, dan bernuansa Islami.
Gunakan paragraf naratif mengalir — bukan poin-poin. Berikan insight mendalam dan rekomendasi konkret actionable.`
    })
  });
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = "";
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    const lines = buf.split("\n"); buf = lines.pop();
    for (const line of lines) {
      if (!line.startsWith("data:")) continue;
      const raw = line.slice(5).trim();
      if (raw === "[DONE]") continue;
      try { const j = JSON.parse(raw); if (j.type==="content_block_delta"&&j.delta?.text) onChunk(j.delta.text); } catch {}
    }
  }
}

/* ─── HELPERS ───────────────────────────────────────────────────────────── */
const detectType = vals => {
  const f = vals.filter(v => v!==""&&v!==null&&v!==undefined);
  if (!f.length) return "empty";
  const nums = f.filter(v => !isNaN(Number(v)) && String(v).trim() !== "");
  if (nums.length===f.length && f.length>3) return "numeric";
  const u = [...new Set(f)];
  if (u.length<=12 || u.length/f.length<0.35) return "categorical";
  return "text";
};
const freqMap = arr => {
  const m = {};
  arr.forEach(v => { const k=String(v??"").trim(); if(k) m[k]=(m[k]||0)+1; });
  return Object.entries(m).sort((a,b) => b[1]-a[1]);
};
const pct  = (n,t) => t ? Math.round(n/t*100) : 0;
const avg  = nums => nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0;

/* ─── UI COMPONENTS ─────────────────────────────────────────────────────── */
function MiniBar({ data, total }) {
  const top = data.slice(0,6), max = top[0]?.[1]||1;
  return (
    <div style={{ marginTop:8 }}>
      {top.map(([k,v],i) => (
        <div key={i} style={{ marginBottom:7 }}>
          <div style={{ display:"flex", justifyContent:"space-between", fontSize:11.5, marginBottom:3 }}>
            <span style={{ maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{k}</span>
            <span style={{ fontWeight:600, color:COLORS[i%8] }}>{v} <span style={{ color:"#888", fontWeight:400 }}>({pct(v,total)}%)</span></span>
          </div>
          <div style={{ height:7, background:"#eee", borderRadius:6, overflow:"hidden" }}>
            <div style={{ height:"100%", width:`${pct(v,max)}%`, background:COLORS[i%8], borderRadius:6, transition:"width .5s" }}/>
          </div>
        </div>
      ))}
    </div>
  );
}

function StatPill({ icon, label, value, accent }) {
  return (
    <div style={{ background:"#fff", border:`1px solid ${accent}33`, borderRadius:14, padding:"16px 20px", flex:"1 1 130px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", right:-8, top:-8, width:60, height:60, borderRadius:"50%", background:accent+"12" }}/>
      <div style={{ fontSize:22, marginBottom:6 }}>{icon}</div>
      <div style={{ fontSize:11, color:"#888", marginBottom:4, textTransform:"uppercase", letterSpacing:.6 }}>{label}</div>
      <div style={{ fontSize:24, fontWeight:700, color:accent }}>{value}</div>
    </div>
  );
}

/* ─── MAIN APP ──────────────────────────────────────────────────────────── */
export default function App() {
  const [page,        setPage]        = useState("landing");
  const [files,       setFiles]       = useState([]);
  const [activeFile,  setActiveFile]  = useState(0);
  const [drag,        setDrag]        = useState(false);
  const [error,       setError]       = useState("");
  const [sheetUrl,    setSheetUrl]    = useState("");
  const [loadingSheet,setLoadingSheet]= useState(false);
  const [tab,         setTab]         = useState("dashboard");
  const [aiSummary,   setAiSummary]   = useState("");
  const [aiReport,    setAiReport]    = useState("");
  const [colInsights, setColInsights] = useState({});
  const [analyzing,   setAnalyzing]   = useState(false);
  const [analyzeStep, setAnalyzeStep] = useState("");
  const [pptxLoading, setPptxLoading] = useState(false);

  const processRows = (rows, name) => {
    if (!rows||rows.length<2) { setError("Data terlalu sedikit."); return; }
    const h    = rows[0].map(String);
    const body = rows.slice(1).filter(r => r.some(c => c!==""&&c!==null));
    setFiles(prev => [...prev, { name, headers:h, data:body }]);
    setError(""); setPage("app"); setTab("dashboard");
    setAiSummary(""); setAiReport(""); setColInsights({});
  };

  const handleFile = useCallback(file => {
    const name = file.name;
    if (name.toLowerCase().endsWith(".csv")) {
      Papa.parse(file, { complete: r => processRows(r.data, name), skipEmptyLines:true });
    } else if (name.toLowerCase().match(/\.xlsx?$/)) {
      const reader = new FileReader();
      reader.onload = e => {
        const wb = XLSX.read(e.target.result, { type:"binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        processRows(XLSX.utils.sheet_to_json(ws, { header:1, defval:"" }), name);
      };
      reader.readAsBinaryString(file);
    } else setError("Format didukung: .xlsx, .xls, .csv");
  }, []);

  const handleDrop = useCallback(e => {
    e.preventDefault(); setDrag(false);
    Array.from(e.dataTransfer.files).forEach(f => handleFile(f));
  }, [handleFile]);

  const loadGSheet = async () => {
    setLoadingSheet(true); setError("");
    try {
      const id = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)?.[1];
      if (!id) { setError("URL tidak valid."); setLoadingSheet(false); return; }
      const res  = await fetch(`https://docs.google.com/spreadsheets/d/${id}/export?format=csv`);
      if (!res.ok) throw new Error("Gagal. Pastikan Sheets sudah diset publik.");
      const text = await res.text();
      Papa.parse(text, { complete: r => processRows(r.data, "Google Sheets"), skipEmptyLines:true });
    } catch(e) { setError(e.message); }
    setLoadingSheet(false);
  };

  const fileData = useMemo(() => {
    if (!files.length) return null;
    const f = files[activeFile]; if (!f) return null;
    const { headers, data } = f;
    const cols = headers.map((h,i) => {
      const vals   = data.map(r => String(r[i]??"").trim());
      const filled = vals.filter(v => v!=="");
      const type   = detectType(filled);
      const freq   = freqMap(filled);
      let mean=0, min=0, max=0, median=0;
      if (type==="numeric") {
        const nums = filled.map(Number).filter(n => !isNaN(n));
        mean=avg(nums); min=Math.min(...nums); max=Math.max(...nums);
        const s=[...nums].sort((a,b)=>a-b); median=s[Math.floor(s.length/2)];
      }
      return { name:h, vals, filled, type, freq, mean, min, max, median, total:data.length };
    });
    const rCols    = cols.filter(c => c.type==="numeric"&&c.mean>=1&&c.mean<=5);
    const avgRating= rCols.length ? avg(rCols.map(c=>c.mean)) : null;
    const fillRate = Math.round(cols.reduce((a,c)=>a+c.filled.length/c.total,0)/cols.length*100);
    return { ...f, cols, avgRating, fillRate };
  }, [files, activeFile]);

  const runAnalysis = async () => {
    if (!fileData) return;
    setAnalyzing(true); setTab("report");
    setAiSummary(""); setAiReport(""); setColInsights({});
    try {
      setAnalyzeStep("Menyusun ringkasan eksekutif…");
      const sp = `Data survei BHN "${fileData.name}": ${fileData.data.length} responden.
Kolom: ${fileData.headers.join(", ")}.
Statistik: ${fileData.cols?.map(c=>`${c.name}[${c.type}${c.type==="numeric"?",avg:"+c.mean.toFixed(2):""}${c.type==="categorical"?",top:"+c.freq[0]?.[0]:""}]`).join("; ")}.
Tulis ringkasan eksekutif 2 paragraf padat untuk laporan BHN.`;
      let sum = "";
      await callClaude(sp, chunk => { sum += chunk; setAiSummary(s => s+chunk); });

      const catCols = fileData.cols?.filter(c=>c.type!=="empty"&&c.type!=="text").slice(0,5)||[];
      for (const col of catCols) {
        setAnalyzeStep(`Menganalisis kolom: ${col.name}…`);
        const p = col.type==="categorical"
          ? `Kolom "${col.name}" survei BHN: distribusi — ${col.freq.slice(0,5).map(([k,v])=>`${k}:${v}`).join(", ")}. Total ${col.filled.length} responden. Tulis 1 paragraf insight.`
          : `Kolom numerik "${col.name}": rata-rata ${col.mean.toFixed(2)}, min ${col.min}, max ${col.max}. Tulis 1 paragraf insight singkat.`;
        let ins = "";
        await callClaude(p, chunk => { ins += chunk; setColInsights(prev => ({ ...prev, [col.name]:ins })); });
      }

      setAnalyzeStep("Menyusun analisis mendalam & rekomendasi…");
      const dp = `Lakukan analisis mendalam (5 paragraf) terhadap data survei BHN "${fileData.name}":
- Total ${fileData.data.length} responden | Kolom: ${fileData.headers.join(", ")}
- Statistik: ${fileData.cols?.slice(0,8).map(c=>`${c.name}[${c.type}${c.type==="numeric"?":avg "+c.mean.toFixed(1):""}${c.type==="categorical"?":top "+c.freq[0]?.[0]:""}]`).join(", ")}
- Ringkasan: ${sum.slice(0,300)}
Analisis mencakup: pola dominan, kekuatan & kelemahan, peluang perbaikan, dan 4-5 rekomendasi strategis konkret untuk BHN.`;
      await callClaude(dp, chunk => setAiReport(r => r+chunk));
    } catch(e) { setError("Gagal analisis: " + e.message); }
    setAnalyzeStep(""); setAnalyzing(false);
  };

  const exportPptx = async () => {
    if (!fileData) return; setPptxLoading(true);
    try {
      const meta = { title:`Laporan — ${fileData.name}`, subtitle:new Date().toLocaleDateString("id-ID",{year:"numeric",month:"long"}), total:fileData.data?.length||0, cols:fileData.headers?.length||0, avgRating:fileData.avgRating, fillRate:fileData.fillRate||0 };
      await buildPptx(meta, aiSummary||"(belum dianalisis)", (fileData.cols||[]).map(c=>({...c,insight:colInsights[c.name]||""})), aiReport||"(belum dianalisis)");
    } catch(e) { setError("Gagal export: "+e.message); }
    setPptxLoading(false);
  };

  /* ── CSS vars shim (tidak ada claude.ai host) ── */
  const css = {
    bg:    "#f4f6f9",
    bgPri: "#ffffff",
    bgSec: "#f0f2f5",
    txt:   "#1a1a2e",
    txtSec:"#666",
    bdr:   "#e0e4ea",
  };

  /* ── LANDING ── */
  if (page==="landing") return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,sans-serif", background:css.bg, minHeight:"100vh", padding:"2rem 1rem" }}>
      <div style={{ maxWidth:860, margin:"0 auto" }}>

        {/* Hero */}
        <div style={{ background:`linear-gradient(135deg,${NAVYDARK},${NAVY} 60%,${NAVYLT})`, borderRadius:20, padding:"3rem 2rem 2.5rem", marginBottom:20, position:"relative", overflow:"hidden" }}>
          <div style={{ position:"absolute", top:-40, right:-40, width:200, height:200, borderRadius:"50%", background:GOLD+"18" }}/>
          <div style={{ position:"absolute", bottom:-30, left:-30, width:150, height:150, borderRadius:"50%", background:GOLD+"10" }}/>
          <div style={{ position:"relative" }}>
            <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:"1.5rem" }}>
              <div style={{ background:GOLD, borderRadius:12, width:44, height:44, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:16, color:NAVY }}>BHN</div>
              <div>
                <div style={{ fontSize:11, color:GOLD+"CC", letterSpacing:2, textTransform:"uppercase", fontWeight:500 }}>Baitul Hikmah Nusantara</div>
                <div style={{ fontSize:12, color:"#ffffff88" }}>Bogor — Lembaga Pendidikan Islam</div>
              </div>
            </div>
            <h1 style={{ fontSize:30, fontWeight:700, color:"#fff", margin:"0 0 12px", lineHeight:1.25 }}>
              Deep Survey &<br/><span style={{ color:GOLD }}>Document Analyzer</span>
            </h1>
            <p style={{ fontSize:14, color:"#ffffffaa", margin:"0 0 2rem", lineHeight:1.7, maxWidth:480 }}>
              Analisis mendalam berbasis AI untuk data survei, dokumen evaluasi, dan laporan BHN —
              insight naratif, dashboard visual, dan PowerPoint siap pakai.
            </p>
            <button onClick={()=>setPage("app")} style={{ background:GOLD, color:NAVY, border:"none", borderRadius:10, padding:"12px 28px", fontSize:14, fontWeight:700, cursor:"pointer" }}>
              Mulai Analisis →
            </button>
          </div>
        </div>

        {/* Feature cards */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))", gap:12, marginBottom:20 }}>
          {[
            { icon:"ti-file-upload",    title:"Upload Fleksibel", desc:"Excel, CSV, atau Google Sheets", color:NAVY },
            { icon:"ti-brain",          title:"Analisis AI",      desc:"Insight naratif mendalam via Claude AI", color:"#8E44AD" },
            { icon:"ti-chart-bar",      title:"Dashboard Visual", desc:"Grafik otomatis per pertanyaan", color:NAVYLT },
            { icon:"ti-file-powerpoint",title:"Export PPTX",      desc:"Presentasi BHN navy+gold siap pakai", color:GOLD },
          ].map((f,i) => (
            <div key={i} style={{ background:css.bgPri, border:`0.5px solid ${css.bdr}`, borderRadius:14, padding:"18px 16px" }}>
              <div style={{ width:40, height:40, borderRadius:10, background:f.color+"18", display:"flex", alignItems:"center", justifyContent:"center", marginBottom:10 }}>
                <i className={`ti ${f.icon}`} aria-hidden style={{ fontSize:20, color:f.color }}/>
              </div>
              <div style={{ fontSize:13, fontWeight:600, marginBottom:4, color:css.txt }}>{f.title}</div>
              <div style={{ fontSize:12, color:css.txtSec, lineHeight:1.5 }}>{f.desc}</div>
            </div>
          ))}
        </div>

        {/* Quick upload */}
        <div style={{ background:css.bgPri, borderRadius:16, padding:"1.5rem", border:`0.5px solid ${css.bdr}` }}>
          <p style={{ fontSize:13, fontWeight:500, margin:"0 0 12px", textAlign:"center", color:css.txtSec }}>Langsung upload file di sini</p>
          <div
            onDragOver={e=>{e.preventDefault();setDrag(true);}}
            onDragLeave={()=>setDrag(false)}
            onDrop={e=>{handleDrop(e);setPage("app");}}
            style={{ border:`2px dashed ${drag?GOLD:NAVY+"44"}`, borderRadius:12, padding:"1.5rem", textAlign:"center", background:drag?GOLD+"08":"transparent", transition:"all .2s", cursor:"pointer" }}
            onClick={()=>setPage("app")}
          >
            <i className="ti ti-cloud-upload" aria-hidden style={{ fontSize:32, color:NAVY, display:"block", marginBottom:8 }}/>
            <p style={{ margin:0, fontSize:13, color:css.txtSec }}>
              Drag & drop Excel/CSV, atau <span style={{ color:NAVY, fontWeight:600 }}>klik untuk pilih file</span>
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  /* ── APP PAGE ── */
  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,sans-serif", background:css.bg, minHeight:"100vh", padding:"1rem" }}>
      <div style={{ maxWidth:900, margin:"0 auto" }}>

        {/* Topbar */}
        <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16, paddingBottom:14, borderBottom:`1px solid ${GOLD}44` }}>
          <button onClick={()=>setPage("landing")} style={{ background:"transparent", border:`0.5px solid ${css.bdr}`, cursor:"pointer", color:css.txtSec, padding:"6px 12px", borderRadius:8, fontSize:13, display:"flex", alignItems:"center", gap:4 }}>
            <i className="ti ti-arrow-left" aria-hidden style={{ fontSize:14 }}/> Beranda
          </button>
          <div style={{ background:NAVY, borderRadius:8, padding:"4px 10px", color:"#fff", fontSize:12, fontWeight:700 }}>BHN</div>
          <span style={{ fontSize:13, fontWeight:500, color:css.txt }}>Deep Analyzer</span>
          {aiReport && (
            <button onClick={exportPptx} disabled={pptxLoading} style={{ marginLeft:"auto", background:`linear-gradient(135deg,${NAVY},${NAVYLT})`, color:"#fff", border:"none", borderRadius:10, padding:"8px 18px", fontSize:13, cursor:"pointer", fontWeight:600, display:"flex", alignItems:"center", gap:6 }}>
              <i className="ti ti-file-powerpoint" aria-hidden/>{pptxLoading?"Membuat…":"Download PPTX"}
            </button>
          )}
        </div>

        {error && (
          <div style={{ background:"#fdecea", border:"0.5px solid #e74c3c44", borderRadius:10, padding:"10px 14px", fontSize:13, color:"#c0392b", marginBottom:12 }}>
            {error}<button onClick={()=>setError("")} style={{ float:"right", background:"none", border:"none", cursor:"pointer", color:"#c0392b", fontWeight:700 }}>✕</button>
          </div>
        )}

        {/* Upload area */}
        {!files.length && (
          <div style={{ background:css.bgPri, borderRadius:16, padding:"2rem", border:`0.5px solid ${css.bdr}` }}>
            <div
              onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop}
              style={{ border:`2px dashed ${drag?GOLD:NAVY+"44"}`, borderRadius:14, padding:"3rem 2rem", textAlign:"center", background:drag?GOLD+"08":"transparent", transition:"all .2s", marginBottom:16 }}
            >
              <i className="ti ti-table-import" aria-hidden style={{ fontSize:36, color:NAVY, display:"block", marginBottom:12 }}/>
              <p style={{ margin:"0 0 6px", fontSize:15, fontWeight:600, color:css.txt }}>Upload File Data</p>
              <p style={{ margin:"0 0 16px", fontSize:12, color:css.txtSec }}>Excel (.xlsx/.xls) atau CSV</p>
              <label style={{ background:NAVY, color:"#fff", borderRadius:10, padding:"10px 26px", fontSize:13, cursor:"pointer", display:"inline-block", fontWeight:600 }}>
                <i className="ti ti-upload" aria-hidden style={{ marginRight:6 }}/>Pilih File
                <input type="file" accept=".xlsx,.xls,.csv" multiple style={{ display:"none" }} onChange={e=>Array.from(e.target.files).forEach(f=>handleFile(f))}/>
              </label>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:8, margin:"8px 0 12px" }}>
              <div style={{ flex:1, height:1, background:css.bdr }}/><span style={{ fontSize:12, color:css.txtSec }}>atau Google Sheets</span><div style={{ flex:1, height:1, background:css.bdr }}/>
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <input placeholder="Paste link Google Sheets publik…" value={sheetUrl} onChange={e=>setSheetUrl(e.target.value)}
                style={{ flex:1, borderRadius:10, padding:"9px 14px", fontSize:13, border:`0.5px solid ${css.bdr}`, outline:"none" }}/>
              <button onClick={loadGSheet} disabled={loadingSheet||!sheetUrl}
                style={{ background:NAVY, color:"#fff", border:"none", borderRadius:10, padding:"9px 20px", fontSize:13, cursor:"pointer", fontWeight:600, whiteSpace:"nowrap" }}>
                {loadingSheet?"Memuat…":"Muat Data"}
              </button>
            </div>
          </div>
        )}

        {/* File tabs */}
        {!!files.length && (
          <div>
            <div style={{ display:"flex", gap:6, flexWrap:"wrap", alignItems:"center", marginBottom:12 }}>
              {files.map((f,i) => (
                <div key={i} onClick={()=>{setActiveFile(i);setAiSummary("");setAiReport("");setColInsights({});}}
                  style={{ display:"flex", alignItems:"center", gap:6, padding:"6px 14px", borderRadius:20, cursor:"pointer", fontSize:12, fontWeight:activeFile===i?600:400, background:activeFile===i?NAVY:css.bgSec, color:activeFile===i?"#fff":css.txtSec, border:`1px solid ${activeFile===i?NAVY:"transparent"}`, transition:"all .2s" }}>
                  <i className="ti ti-table" aria-hidden style={{ fontSize:12 }}/>
                  {f.name.length>22 ? f.name.slice(0,22)+"…" : f.name}
                  <span onClick={e=>{e.stopPropagation();setFiles(p=>p.filter((_,j)=>j!==i));if(activeFile>=i)setActiveFile(0);}} style={{ marginLeft:2, opacity:.6 }}>✕</span>
                </div>
              ))}
              <label style={{ display:"flex", alignItems:"center", gap:4, padding:"6px 14px", borderRadius:20, cursor:"pointer", fontSize:12, border:`1px dashed ${GOLD}`, color:GOLD, background:GOLD+"0A", fontWeight:500 }}>
                <i className="ti ti-plus" aria-hidden/>Tambah
                <input type="file" accept=".xlsx,.xls,.csv" multiple style={{ display:"none" }} onChange={e=>Array.from(e.target.files).forEach(f=>handleFile(f))}/>
              </label>
              {!aiReport ? (
                <button onClick={runAnalysis} disabled={analyzing}
                  style={{ marginLeft:"auto", background:analyzing?css.bgSec:`linear-gradient(135deg,${NAVY},${NAVYLT})`, color:analyzing?css.txtSec:"#fff", border:"none", borderRadius:20, padding:"8px 20px", fontSize:13, cursor:analyzing?"wait":"pointer", fontWeight:600, display:"flex", alignItems:"center", gap:6 }}>
                  {analyzing ? <><span style={{ animation:"spin 1s linear infinite", display:"inline-block" }}>⚙</span>{analyzeStep||"Menganalisis…"}</>
                             : <><i className="ti ti-brain" aria-hidden/>Analisis AI Mendalam</>}
                </button>
              ) : (
                <button onClick={runAnalysis} disabled={analyzing}
                  style={{ marginLeft:"auto", background:"transparent", color:NAVY, border:`1px solid ${NAVY}44`, borderRadius:20, padding:"7px 16px", fontSize:12, cursor:"pointer", display:"flex", alignItems:"center", gap:5 }}>
                  <i className="ti ti-refresh" aria-hidden/>Ulang Analisis
                </button>
              )}
            </div>

            {/* Nav tabs */}
            <div style={{ display:"flex", gap:2, marginBottom:16, background:css.bgSec, borderRadius:12, padding:4 }}>
              {[["dashboard","ti-layout-dashboard","Dashboard"],["report","ti-file-analytics","Laporan AI"],["data","ti-table","Data Mentah"]].map(([id,icon,label]) => (
                <button key={id} onClick={()=>setTab(id)} style={{ flex:1, background:tab===id?css.bgPri:"transparent", color:tab===id?NAVY:css.txtSec, border:tab===id?`1px solid ${NAVY}22`:"none", borderRadius:9, padding:"8px", fontSize:12, cursor:"pointer", fontWeight:tab===id?600:400, display:"flex", alignItems:"center", justifyContent:"center", gap:5, transition:"all .2s", boxShadow:tab===id?"0 1px 4px #0001":"none" }}>
                  <i className={`ti ${icon}`} aria-hidden style={{ fontSize:13 }}/>{label}
                </button>
              ))}
            </div>

            {/* DASHBOARD TAB */}
            {tab==="dashboard" && fileData && (
              <div>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:16 }}>
                  <StatPill icon="👥" label="Total Data"    value={fileData.data?.length||0}              accent={NAVY}/>
                  <StatPill icon="📋" label="Pertanyaan"    value={fileData.headers?.length||0}           accent={GOLD}/>
                  <StatPill icon="✅" label="Pengisian"     value={(fileData.fillRate||0)+"%"}            accent="#27AE60"/>
                  {fileData.avgRating && <StatPill icon="⭐" label="Avg Rating" value={fileData.avgRating.toFixed(2)+"/5"} accent="#E67E22"/>}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))", gap:12 }}>
                  {fileData.cols?.filter(c=>c.type!=="empty").map((col,i) => (
                    <div key={i} style={{ background:css.bgPri, border:`0.5px solid ${css.bdr}`, borderRadius:14, padding:16 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8 }}>
                        <span style={{ fontSize:13, fontWeight:600, flex:1, lineHeight:1.3, color:css.txt }}>{col.name}</span>
                        <span style={{ fontSize:10, padding:"2px 8px", borderRadius:20, background:col.type==="categorical"?NAVY+"18":col.type==="numeric"?GOLD+"22":"#f0f0f0", color:col.type==="categorical"?NAVY:col.type==="numeric"?GOLD:"#888", marginLeft:8, whiteSpace:"nowrap", fontWeight:500 }}>{col.type}</span>
                      </div>
                      <div style={{ fontSize:11, color:css.txtSec, marginBottom:10 }}>
                        {col.filled.length}/{col.total} terisi ({pct(col.filled.length,col.total)}%)
                      </div>
                      {col.type==="categorical" && <MiniBar data={col.freq} total={col.filled.length}/>}
                      {col.type==="numeric" && (
                        <div>
                          <div style={{ fontSize:26, fontWeight:700, color:NAVY, marginBottom:4 }}>{col.mean.toFixed(2)}<span style={{ fontSize:12, color:css.txtSec, marginLeft:4, fontWeight:400 }}>rata-rata</span></div>
                          <div style={{ display:"flex", gap:16, fontSize:11, color:css.txtSec }}>
                            <span>Min <strong style={{ color:css.txt }}>{col.min}</strong></span>
                            <span>Max <strong style={{ color:css.txt }}>{col.max}</strong></span>
                            <span>Med <strong style={{ color:css.txt }}>{col.median.toFixed(1)}</strong></span>
                          </div>
                        </div>
                      )}
                      {col.type==="text" && <div style={{ fontSize:11, color:css.txtSec, fontStyle:"italic", lineHeight:1.5 }}>"{(col.vals.find(v=>v.length>15)||col.vals[0]||"").slice(0,80)}…"<div style={{ marginTop:4, fontStyle:"normal" }}>{col.filled.length} respons teks</div></div>}
                      {colInsights[col.name] && (
                        <div style={{ marginTop:12, background:`linear-gradient(135deg,${NAVY}08,${GOLD}08)`, borderLeft:`3px solid ${GOLD}`, borderRadius:"0 8px 8px 0", padding:"10px 12px" }}>
                          <div style={{ fontSize:9, fontWeight:700, color:GOLD, marginBottom:4, letterSpacing:.5, textTransform:"uppercase" }}>Insight AI</div>
                          <p style={{ fontSize:11, color:css.txt, lineHeight:1.6, margin:0 }}>{colInsights[col.name].slice(0,180)}{colInsights[col.name].length>180?"…":""}</p>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* REPORT TAB */}
            {tab==="report" && (
              <div>
                {!aiSummary && !analyzing && (
                  <div style={{ textAlign:"center", padding:"4rem 2rem", color:css.txtSec }}>
                    <div style={{ width:72, height:72, borderRadius:20, background:NAVY+"12", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 16px" }}>
                      <i className="ti ti-brain" aria-hidden style={{ fontSize:36, color:NAVY+"66" }}/>
                    </div>
                    <p style={{ fontSize:14, margin:"0 0 16px" }}>Belum ada analisis AI.</p>
                    <button onClick={runAnalysis} style={{ background:`linear-gradient(135deg,${NAVY},${NAVYLT})`, color:"#fff", border:"none", borderRadius:10, padding:"10px 24px", fontSize:13, cursor:"pointer", fontWeight:600 }}>
                      Mulai Analisis AI
                    </button>
                  </div>
                )}
                {analyzing && analyzeStep && (
                  <div style={{ background:NAVY+"10", border:`1px solid ${NAVY}22`, borderRadius:12, padding:"12px 16px", fontSize:13, color:NAVY, marginBottom:14, display:"flex", alignItems:"center", gap:10 }}>
                    <span style={{ display:"inline-block", animation:"spin 1s linear infinite", fontSize:18 }}>⚙</span>
                    <div><strong>Sedang menganalisis</strong> — {analyzeStep}</div>
                  </div>
                )}
                {aiSummary && (
                  <div style={{ background:css.bgPri, border:`0.5px solid ${css.bdr}`, borderRadius:14, padding:20, marginBottom:12 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
                      <div style={{ width:32, height:32, borderRadius:8, background:NAVY, display:"flex", alignItems:"center", justifyContent:"center" }}><i className="ti ti-file-description" aria-hidden style={{ fontSize:16, color:GOLD }}/></div>
                      <h3 style={{ fontSize:14, fontWeight:700, color:NAVY, margin:0 }}>Ringkasan Eksekutif</h3>
                    </div>
                    <p style={{ fontSize:13, lineHeight:1.85, margin:0, color:css.txt }}>{aiSummary}</p>
                  </div>
                )}
                {Object.keys(colInsights).length>0 && (
                  <div style={{ background:css.bgPri, border:`0.5px solid ${css.bdr}`, borderRadius:14, padding:20, marginBottom:12 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
                      <div style={{ width:32, height:32, borderRadius:8, background:GOLD, display:"flex", alignItems:"center", justifyContent:"center" }}><i className="ti ti-chart-dots" aria-hidden style={{ fontSize:16, color:NAVY }}/></div>
                      <h3 style={{ fontSize:14, fontWeight:700, color:NAVY, margin:0 }}>Insight Per Kolom</h3>
                    </div>
                    {Object.entries(colInsights).map(([k,v]) => (
                      <div key={k} style={{ marginBottom:14, paddingBottom:14, borderBottom:`0.5px solid ${css.bdr}` }}>
                        <div style={{ fontSize:12, fontWeight:700, color:GOLD, marginBottom:5, display:"flex", alignItems:"center", gap:6 }}>
                          <span style={{ width:6, height:6, borderRadius:"50%", background:GOLD, display:"inline-block" }}/>{k}
                        </div>
                        <p style={{ fontSize:12, lineHeight:1.75, margin:0, color:css.txt }}>{v}</p>
                      </div>
                    ))}
                  </div>
                )}
                {aiReport && (
                  <div style={{ background:css.bgPri, border:`1px solid ${NAVY}22`, borderRadius:14, padding:20, marginBottom:14 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
                      <div style={{ width:32, height:32, borderRadius:8, background:`linear-gradient(135deg,${NAVY},${NAVYLT})`, display:"flex", alignItems:"center", justifyContent:"center" }}><i className="ti ti-report-analytics" aria-hidden style={{ fontSize:16, color:"#fff" }}/></div>
                      <h3 style={{ fontSize:14, fontWeight:700, color:NAVY, margin:0 }}>Analisis Mendalam & Rekomendasi</h3>
                    </div>
                    <div style={{ fontSize:13, lineHeight:1.85, color:css.txt, whiteSpace:"pre-wrap" }}>{aiReport}</div>
                  </div>
                )}
                {aiReport && (
                  <div style={{ display:"flex", justifyContent:"flex-end" }}>
                    <button onClick={exportPptx} disabled={pptxLoading}
                      style={{ background:`linear-gradient(135deg,${NAVY},${NAVYLT})`, color:"#fff", border:"none", borderRadius:10, padding:"10px 22px", fontSize:13, cursor:"pointer", fontWeight:700, display:"flex", alignItems:"center", gap:7, boxShadow:`0 4px 12px ${NAVY}44` }}>
                      <i className="ti ti-file-powerpoint" aria-hidden style={{ fontSize:16 }}/>
                      {pptxLoading ? "Menyusun presentasi…" : "Download PowerPoint"}
                    </button>
                  </div>
                )}
              </div>
            )}

            {/* DATA TAB */}
            {tab==="data" && fileData && (
              <div>
                <p style={{ fontSize:12, color:css.txtSec, marginBottom:10 }}>Menampilkan 50 baris pertama dari {fileData.data?.length||0} entri</p>
                <div style={{ overflowX:"auto", borderRadius:12, border:`0.5px solid ${css.bdr}` }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11.5, tableLayout:"fixed" }}>
                    <thead>
                      <tr style={{ background:NAVY }}>
                        <th style={{ width:32, padding:"8px 6px", color:GOLD, textAlign:"center", fontWeight:600, fontSize:10 }}>#</th>
                        {fileData.headers?.map((h,i) => (
                          <th key={i} style={{ padding:"8px 10px", color:"#fff", textAlign:"left", fontWeight:600, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:130, fontSize:11 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {fileData.data?.slice(0,50).map((row,ri) => (
                        <tr key={ri} style={{ background:ri%2===0?css.bgPri:css.bgSec }}>
                          <td style={{ padding:"6px", textAlign:"center", color:css.txtSec, fontSize:10 }}>{ri+1}</td>
                          {fileData.headers?.map((_,ci) => (
                            <td key={ci} title={String(row[ci]??"")}
                              style={{ padding:"6px 10px", borderBottom:`0.5px solid ${css.bdr}`, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:130, color:css.txt }}>
                              {String(row[ci]??"")}
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
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
