// =========================================
// Villa Contabilidade — Comparador SINTEGRA × XML
// (UTF-8)
// =========================================

// --------- COLAR A SUA IMAGEM AQUI (data URL completo!) ----------
const BRAND_DATA_URL =
  "data:image/png;base64,iVBORw0K..."; // ← troque pelo data URL COMPLETO que você enviou

// ---------- Helpers de UI ----------
const $ = (id) => document.getElementById(id);
const moneyBR = (v) => {
  if (v == null || v === "") return "";
  const n = Number(String(v).replace(/\./g, "").replace(",", "."));
  return Number.isFinite(n) ? n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }) : String(v);
};
const parseNumberBR = (s) => {
  if (!s) return null;
  const t = String(s).replace(/\./g, "").replace(",", ".");
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
};
const digits = (s) => (s || "").replace(/\D+/g, "");
const pad = (n, l = 2) => String(n).padStart(l, "0");
const toYMD = (d) => {
  if (!d) return "";
  const m = String(d).match(/^(\d{4})-?(\d{2})-?(\d{2})/);
  return m ? `${m[1]}${m[2]}${m[3]}` : "";
};
const ymdToBR = (ymd) => (ymd && ymd.length === 8 ? `${ymd.slice(6, 8)}/${ymd.slice(4, 6)}/${ymd.slice(0, 4)}` : ymd || "");
const normSerie = (s) => (digits(s) ? String(Number(digits(s))) : "");
const normNF = (n) => (digits(n) ? String(Number(digits(n))) : "");
const sameValue = (a, b, tol = 0.02) => {
  const na = parseNumberBR(a), nb = parseNumberBR(b);
  if (!Number.isFinite(na) || !Number.isFinite(nb)) return String(a) === String(b);
  return Math.abs(na - nb) <= tol;
};
const tableFromRows = (header, rows) => {
  const thead = `<thead><tr>${header.map(h => `<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${
    rows.length
      ? rows.map(r => `<tr>${r.map(c => `<td>${c ?? ""}</td>`).join("")}</tr>`).join("")
      : `<tr><td colspan="${header.length}">Sem registros.</td></tr>`
  }</tbody>`;
  return `<div class="table-wrap"><table>${thead}${tbody}</table></div>`;
};

// ---------- Branding dinâmico ----------
function setBranding() {
  const fav = $("dynamic-favicon");
  const logo = $("brandLogo");
  if (fav && BRAND_DATA_URL) fav.href = BRAND_DATA_URL;
  if (logo && BRAND_DATA_URL) logo.src = BRAND_DATA_URL;
}

// ---------- Upload UX (animação, nomes, reticências) ----------
function wireUploadUI(labelEl, inputEl, titulo) {
  function refresh() {
    const span = labelEl.querySelector("span");
    const files = Array.from(inputEl.files || []);
    if (!span) return;
    if (!files.length) {
      labelEl.classList.remove("filled");
      span.innerHTML = `<strong>${titulo}</strong><br><small>Arraste aqui ou clique</small>`;
      return;
    }
    const first = files[0]?.name || "arquivo";
    const tooltip = files.slice(0, 8).map(f => f.name).join("\n") + (files.length > 8 ? `\n+${files.length - 8} mais...` : "");
    span.innerHTML = `
      <strong>${titulo}</strong><br>
      <small>${files.length === 1 ? "1 arquivo" : files.length + " arquivos"}</small>
      <div class="file-meta">
        <span class="name" title="${tooltip.replace(/"/g, "&quot;")}" style="display:inline-block; max-width: 260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${first}</span>
        ${files.length > 1 ? `<span class="count">+${files.length - 1}</span>` : ``}
      </div>
    `;
    labelEl.classList.add("filled");
  }
  inputEl.addEventListener("change", refresh);
  ["dragenter", "dragover"].forEach(ev => {
    labelEl.addEventListener(ev, e => { e.preventDefault(); labelEl.classList.add("dragover"); });
  });
  ["dragleave", "drop"].forEach(ev => {
    labelEl.addEventListener(ev, e => { e.preventDefault(); labelEl.classList.remove("dragover"); });
  });
  labelEl.addEventListener("drop", e => {
    e.preventDefault();
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length) {
      inputEl.files = dt.files;
      refresh();
    }
  });
}

// =========================================
// Leitura de XML (NF-e) e ZIPs com XML
// =========================================
function parseXML(text) { return new DOMParser().parseFromString(text, "text/xml"); }
function byLocalNameAll(doc, name) {
  const all = doc.getElementsByTagName("*"); const out = [];
  for (let i = 0; i < all.length; i++) if ((all[i].localName || all[i].nodeName) === name) out.push(all[i]);
  return out;
}
function firstByLocalName(doc, name) { return byLocalNameAll(doc, name)[0] || null; }
function findPathLocal(root, path) {
  if (!root) return null;
  let cur = root;
  for (const seg of path) {
    let found = null;
    for (const ch of cur.children) if ((ch.localName || ch.nodeName) === seg) { found = ch; break; }
    if (!found) return null;
    cur = found;
  }
  return cur;
}
function pickInfNFe(doc) {
  const inf = firstByLocalName(doc, "infNFe");
  if (inf) return inf;
  const nfe = firstByLocalName(doc, "NFe");
  if (nfe) return firstByLocalName(nfe, "infNFe") || nfe;
  return doc.documentElement;
}
function getAccessKey(doc) {
  const infProt = firstByLocalName(doc, "infProt");
  if (infProt) {
    const ch = findPathLocal(infProt, ["chNFe"]);
    if (ch && ch.textContent) return digits(ch.textContent);
  }
  const inf = pickInfNFe(doc);
  if (inf && inf.getAttribute && inf.getAttribute("Id")) {
    const v = inf.getAttribute("Id");
    if (v && v.toUpperCase().startsWith("NFE")) return digits(v);
  }
  return "";
}
function extractParty(inf, which) {
  const base = which === "dest" ? "dest" : "emit";
  const name = findPathLocal(inf, [base, "xNome"])?.textContent || "—";
  const docn = findPathLocal(inf, [base, "CNPJ"]) || findPathLocal(inf, [base, "CPF"]);
  const doc = digits(docn?.textContent || "");
  const end = findPathLocal(inf, [base, "ender" + (base === "emit" ? "Emit" : "Dest")]);
  const uf = (end ? findPathLocal(end, ["UF"])?.textContent : "" || "").trim().toUpperCase();
  return { name, doc, uf };
}
function extractItems(inf) {
  const dets = byLocalNameAll(inf, "det");
  const rows = [];
  for (const d of dets) {
    const prod = findPathLocal(d, ["prod"]); if (!prod) continue;
    const xProd = findPathLocal(prod, ["xProd"])?.textContent || "";
    const CFOP  = findPathLocal(prod, ["CFOP"])?.textContent || "";
    const uCom  = findPathLocal(prod, ["uCom"])?.textContent || "";
    const qCom  = findPathLocal(prod, ["qCom"])?.textContent || "";
    const vProd = findPathLocal(prod, ["vProd"])?.textContent || "";
    rows.push({ xProd, CFOP, uCom, qCom, vProd });
  }
  return rows;
}
function extractBasicsXML(text) {
  const doc = parseXML(text);
  const inf = pickInfNFe(doc);
  const key  = getAccessKey(doc);
  const mod  = findPathLocal(inf, ["ide", "mod"])?.textContent || "";
  const nNF  = findPathLocal(inf, ["ide", "nNF"])?.textContent || "";
  const serie= findPathLocal(inf, ["ide", "serie"])?.textContent || "";
  const dhEmi= (findPathLocal(inf, ["ide", "dhEmi"])?.textContent || findPathLocal(inf, ["ide", "dEmi"])?.textContent || "");
  const vNF  = findPathLocal(inf, ["total", "ICMSTot", "vNF"])?.textContent || "";
  const emit = extractParty(inf, "emit");
  const dest = extractParty(inf, "dest");
  const dets = extractItems(inf);
  return { key, mod, nNF, serie, dhEmi, vNF, emit, dest, dets };
}

async function collectXMLs(fileList) {
  const items = [];
  for (const file of fileList) {
    const low = file.name.toLowerCase();
    if (low.endsWith(".xml")) {
      items.push({ name: file.name, text: await file.text() });
    } else if (low.endsWith(".zip")) {
      const zip = await JSZip.loadAsync(file);
      for (const entry of Object.values(zip.files)) {
        if (!entry.dir && entry.name.toLowerCase().endsWith(".xml")) {
          const txt = await entry.async("string");
          const base = entry.name.split("/").pop();
          items.push({ name: base, text: txt });
        }
      }
    }
  }
  const out = [];
  for (const it of items) {
    try {
      const b = extractBasicsXML(it.text);
      if ((b.mod || "") === "55") out.push({ ...b, filename: it.name, raw: it.text });
    } catch (e) {
      console.warn("XML ignorado (erro parse):", it.name, e);
    }
  }
  return out;
}

// Descobre CNPJ da empresa e marca tipo (ENTRADA/SAÍDA) para cada XML
function tagXmlDirection(xmlList) {
  if (!xmlList.length) return { companyDoc: "", entradas: [], saidas: [] };

  // Heurística: o CNPJ mais frequente entre (emit, dest) vira o CNPJ da empresa
  const freq = new Map();
  for (const x of xmlList) {
    if (x.emit?.doc) freq.set(x.emit.doc, (freq.get(x.emit.doc) || 0) + 1);
    if (x.dest?.doc) freq.set(x.dest.doc, (freq.get(x.dest.doc) || 0) + 1);
  }
  let companyDoc = ""; let max = -1;
  for (const [doc, c] of freq.entries()) if (c > max) { max = c; companyDoc = doc; }

  for (const x of xmlList) {
    x._tipo = (x.emit?.doc === companyDoc) ? "SAÍDA" : (x.dest?.doc === companyDoc) ? "ENTRADA" : "DESCONHECIDO";
  }
  const entradas = xmlList.filter(x => x._tipo === "ENTRADA");
  const saidas   = xmlList.filter(x => x._tipo === "SAÍDA");
  return { companyDoc, entradas, saidas };
}

// =========================================
// SINTEGRA (.txt) — suporta “zerado” sem quebrar
// - Lê Registro 10 (informante) para CNPJ da empresa
// - Lê Registro 50 (NF modelo 55) com heurísticas de campos
// - Tenta inferir ENTRADA/SAÍDA (se não der, marca DESCONHECIDO)
// =========================================
function parseSintegraDelimited(line) {
  const delim = [";", ",", "|"].find(d => line.includes(d));
  if (!delim) return null;
  const t = line.split(delim).map(s => s.trim());
  if (t[0] !== "50") return null;

  // Heurística comum de posições (varia por UF/layout)
  let idx = { data: 4, mod: 5, serie: 6, nf: 7, valor: 11, cnpjCampo: 1 };
  // Ajustes se necessário
  if (!/^\d{8}$/.test(t[idx.data])) {
    const id = t.findIndex(x => /^\d{8}$/.test(x));
    if (id >= 0) idx.data = id;
  }
  if (t[idx.mod] !== "55") {
    const im = t.findIndex(x => x === "55");
    if (im >= 0) idx.mod = im;
  }
  if (!t[idx.serie] || !t[idx.nf]) {
    idx.serie = Math.min(idx.mod + 1, t.length - 1);
    idx.nf    = Math.min(idx.mod + 2, t.length - 1);
  }
  if (!t[idx.valor] || !/[0-9]/.test(t[idx.valor])) {
    for (let i = t.length - 1; i >= 0; i--) {
      if (/[\d,\.]/.test(t[i])) { idx.valor = i; break; }
    }
  }
  return {
    mod: t[idx.mod],
    serie: t[idx.serie],
    nNF: t[idx.nf],
    data: t[idx.data],
    vNF: t[idx.valor],
    cnpjCampo: digits(t[idx.cnpjCampo] || "")
  };
}
function parseSintegraFixed(line) {
  if (!line.startsWith("50")) return null;

  // CNPJ provável perto do início (14 dígitos)
  const mCNPJ = line.match(/^\d{2}\s*([0-9]{14})/) || line.match(/([0-9]{14})/);
  const cnpjCampo = mCNPJ ? mCNPJ[1] : "";

  // data AAAAMMDD (primeira ocorrência)
  const mData = line.match(/(\d{8})/);
  const data = mData ? mData[1] : "";

  // modelo 55
  const m55 = line.match(/55/);
  if (!m55) return null;

  // série (1-3 dígitos) após 55
  const after55 = line.slice(m55.index + 2);
  const mSerie = after55.match(/^\s*([0-9]{1,3})/);
  const serie = mSerie ? mSerie[1] : "";

  // número após série
  const afterSerie = mSerie ? after55.slice(mSerie[0].length) : after55;
  const mNF = afterSerie.match(/^\s*([0-9]{1,9})/);
  const nNF = mNF ? mNF[1] : "";

  // valor: último número com 2 casas
  const mVals = [...line.matchAll(/([0-9]{1,15}[,\.][0-9]{2})/g)];
  const vNF = mVals.length ? mVals[mVals.length - 1][1] : "";

  return { mod: "55", serie, nNF, data, vNF, cnpjCampo };
}
function parseSintegraLine(line) {
  if (!line || !line.trim()) return null;
  if (line.startsWith("10")) { // Registro do informante — tratamos fora
    const cnpj = digits(line.slice(2, 16));
    return { _reg: 10, cnpjInformante: cnpj };
  }
  if (!line.startsWith("50")) return null;
  return parseSintegraDelimited(line) || parseSintegraFixed(line);
}

async function collectSintegra(file) {
  if (!file) return { header: null, rows: [] };
  const text = await file.text();
  const lines = text.split(/\r?\n/);
  let header = null;
  const rows50 = [];

  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    try {
      if (line.startsWith("10")) {
        const h = parseSintegraLine(line);
        if (h && h._reg === 10) header = { cnpj: h.cnpjInformante };
      } else if (line.startsWith("50")) {
        const rec = parseSintegraLine(line);
        if (rec && rec.mod === "55") rows50.push(rec);
      }
    } catch (e) {
      console.warn("Sintegra linha ignorada:", e);
    }
  }
  return { header, rows: rows50 };
}

// Marca ENTRADA/SAÍDA no SINTEGRA quando possível (heurística)
function tagSintegraDirection(rows, companyDoc) {
  for (const r of rows) {
    // se o campo CNPJ da linha for o mesmo CNPJ da empresa, assumimos SAÍDA
    // caso contrário, ENTRADA. Se faltou tudo, DESCONHECIDO.
    if (companyDoc && r.cnpjCampo) {
      r._tipo = (r.cnpjCampo === companyDoc) ? "SAÍDA" : "ENTRADA";
    } else {
      r._tipo = "DESCONHECIDO";
    }
  }
  const entradas = rows.filter(r => r._tipo === "ENTRADA");
  const saidas   = rows.filter(r => r._tipo === "SAÍDA");
  const desconhe = rows.filter(r => r._tipo === "DESCONHECIDO");
  return { entradas, saidas, desconhe };
}

// =========================================
// Relatório PDF (about:blank) — somente diferentes, sem chave
// =========================================
function openClientPDFPreview(meta, diffs) {
  const rows = diffs.map(r => ({
    origem: r.origem,
    nNF: r.nNF || "",
    serie: r.serie || "",
    emissao: ymdToBR(r.data || r.emissao || ""),
    valor: moneyBR(r.vNF || r.valor || ""),
    emitName: r.emit?.name || "—",
    emitDoc: r.emit?.doc || "",
    destName: r.dest?.name || "—",
    destDoc: r.dest?.doc || ""
  }));

  const css = `
  @page { margin: 16mm; }
  body{ font-family:Segoe UI,Roboto,Arial,sans-serif; color:#222; }
  header{ display:flex; align-items:center; gap:10px; margin-bottom:10px; }
  header img{ width:34px; height:34px; border-radius:8px; }
  header .title{ font-size:18px; font-weight:700; color:#8B0000; }
  .meta{ font-size:12px; color:#666; margin-bottom:8px; }
  table{ width:100%; border-collapse:collapse; font-size:12px; }
  thead{ display: table-header-group; }
  tr{ page-break-inside: avoid; }
  th,td{ border-bottom:1px solid #eee; padding:6px 6px; text-align:left; }
  th{ background:#8B0000; color:#fff; }
  td.num{ text-align:right; }
  .section-title{ margin:14px 0 6px; color:#8B0000; font-weight:700; }
  `;

  const headerCols = ["Origem","NF","Série","Emissão","Valor (R$)","Emitente","Documento","Destinatário","Documento"];
  const bodyRows = rows.map(r => `
    <tr>
      <td>${r.origem}</td>
      <td>${r.nNF}</td>
      <td>${r.serie}</td>
      <td>${r.emissao}</td>
      <td class="num">${r.valor}</td>
      <td>${r.emitName}</td>
      <td>${r.emitDoc}</td>
      <td>${r.destName}</td>
      <td>${r.destDoc}</td>
    </tr>`).join("");

  const html = `
<!doctype html><html><head><meta charset="utf-8">
  <title>Relatório — Notas Diferentes</title>
  <style>${css}</style>
</head><body>
  <header>
    <img src="${BRAND_DATA_URL}" alt="Logo">
    <div class="title">Relatório de Conferência — Notas Diferentes</div>
  </header>
  <div class="meta">
    Empresa: <strong>${meta.companyName || "—"}</strong> — ${meta.companyDoc || ""}<br>
    Gerado em: ${new Date().toLocaleString("pt-BR")} • Total: <strong>${rows.length}</strong>
  </div>

  <div class="section-title">Notas para conferência</div>
  <table>
    <thead><tr>${headerCols.map(h=>`<th>${h}</th>`).join("")}</tr></thead>
    <tbody>${bodyRows || `<tr><td colspan="9">Não há notas diferentes.</td></tr>`}</tbody>
  </table>
  <script>setTimeout(function(){try{window.focus();window.print();}catch(e){}},350);</script>
</body></html>`;
  const w = window.open("", "_blank");
  if (!w) { alert("Permita pop-ups para visualizar o PDF."); return; }
  w.document.open(); w.document.write(html); w.document.close();
}

// =========================================
// DANFE Pro (about:blank) com barcode
// =========================================
function danfeHTML(x) {
  const chave = digits(x.key || "");
  const chaveFmt = chave ? chave.replace(/(\d{4})(?=\d)/g, "$1 ").trim() : "—";
  const css = `
    @page { size:A4; margin:12mm; }
    body{ font-family:Segoe UI,Roboto,Arial,sans-serif; color:#111; }
    .box{ border:1px solid #000; border-radius:4px; padding:6px; margin-bottom:6px; }
    .grid2{ display:grid; grid-template-columns:1fr 1fr; gap:6px; }
    .grid3{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:6px; }
    h1{ font-size:18px; margin:0 0 4px; }
    .muted{ color:#555; font-size:12px; }
    table{ width:100%; border-collapse:collapse; font-size:12px; }
    thead{ display:table-header-group; }
    th,td{ border:1px solid #000; padding:4px 6px; }
    th{ background:#eee; }
    td.r{ text-align:right; }
    td.c{ text-align:center; }
    .barcode{ width:100%; height:52px; }
    .logo{ width:48px; height:48px; border-radius:8px; }
    .header{ display:flex; align-items:center; justify-content:space-between; margin-bottom:4px; }
    .header .left{ display:flex; align-items:center; gap:8px; }
  `;
  const items = (x.dets || []).map((d, i) => `
    <tr>
      <td class="c">${i + 1}</td>
      <td>${d.xProd || ""}</td>
      <td class="c">${d.CFOP || ""}</td>
      <td class="c">${d.uCom || ""}</td>
      <td class="r">${d.qCom || ""}</td>
      <td class="r">${moneyBR(d.vProd || "")}</td>
    </tr>`).join("");
  return `
<!doctype html>
<html><head><meta charset="utf-8"><title>DANFE — ${x.nNF || ""}/${x.serie || ""}</title>
<style>${css}</style></head>
<body>
  <div class="header">
    <div class="left">
      <img class="logo" src="${BRAND_DATA_URL}" alt="Logo">
      <div>
        <h1>DANFE — Documento Auxiliar da NF-e</h1>
        <div class="muted">NF-e nº <strong>${x.nNF || ""}</strong> — Série <strong>${x.serie || ""}</strong> — Emissão <strong>${ymdToBR(toYMD(x.dhEmi)) || ""}</strong></div>
      </div>
    </div>
    <div class="right"><svg id="barcode"></svg></div>
  </div>

  <div class="box"><strong>Chave de Acesso:</strong> ${chaveFmt}</div>

  <div class="grid2">
    <div class="box"><strong>Emitente</strong><br>${x.emit?.name || "—"}<br>Doc: ${x.emit?.doc || ""} • UF: ${x.emit?.uf || ""}</div>
    <div class="box"><strong>Destinatário</strong><br>${x.dest?.name || "—"}<br>Doc: ${x.dest?.doc || ""} • UF: ${x.dest?.uf || ""}</div>
  </div>

  <div class="grid3">
    <div class="box"><strong>Número</strong><br>${x.nNF || ""}</div>
    <div class="box"><strong>Série</strong><br>${x.serie || ""}</div>
    <div class="box"><strong>Valor Total</strong><br>${moneyBR(x.vNF || "")}</div>
  </div>

  <div class="box">
    <strong>Itens da NF-e</strong>
    <table>
      <thead><tr><th>#</th><th>Descrição</th><th>CFOP</th><th>Un</th><th>Qtd</th><th>Valor</th></tr></thead>
      <tbody>${items || '<tr><td colspan="6" class="c">Itens não informados no XML.</td></tr>'}</tbody>
    </table>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.6/dist/JsBarcode.all.min.js"></script>
  <script>
    (function(){
      try{
        var key = "${chave}";
        if (key && window.JsBarcode) JsBarcode("#barcode", key, {format:"CODE128", displayValue:false, height:48, margin:0});
      }catch(e){}
      setTimeout(function(){ try{ window.focus(); window.print(); }catch(_){} }, 400);
    })();
  </script>
</body></html>`;
}
function openDANFEWindow(x) {
  const w = window.open("", "_blank");
  if (!w) { alert("Permita pop-ups para visualizar o DANFE."); return; }
  w.document.open(); w.document.write(danfeHTML(x)); w.document.close();
}

// =========================================
// Excel
// =========================================
function buildExcelSheets(common, onlyS, onlyX) {
  const wb = XLSX.utils.book_new();
  const cols = ["origem","modelo","série","nNF","emissão","valor","emitente","emit_doc","destinatário","dest_doc","tipo"];
  const toAOA = (arr) => [cols, ...arr.map(r => ([
    r.origem || "",
    "55",
    r.serie || "",
    r.nNF || "",
    r.data || r.emissao || "",
    r.vNF || r.valor || "",
    r.emit?.name || "",
    r.emit?.doc || "",
    r.dest?.name || "",
    r.dest?.doc || "",
    r._tipo || ""
  ]))];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA(common)), "em_comum");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA(onlyS)), "so_sintegra");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA(onlyX)), "so_xmls");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// =========================================
async function compareAll(fileSintegra, filesXML) {
  const warn = $("warn"); warn.textContent = "";
  $("progressBar").classList.remove("hidden");

  // Carrega XML (mod 55)
  const X = await collectXMLs(filesXML);

  // Descobre CNPJ da empresa pelas NFe
  const { companyDoc, entradas: xmlEntradas, saidas: xmlSaidas } = tagXmlDirection(X);

  // Carrega SINTEGRA (suporta "zerado")
  let SHeader = null, SRows = [];
  if (fileSintegra) {
    const S = await collectSintegra(fileSintegra);
    SHeader = S.header; SRows = S.rows;
  }
  // Se SINTEGRA tiver CNPJ do informante, preferimos esse como companyDoc (mais confiável)
  const companyDocFinal = digits(SHeader?.cnpj || "") || companyDoc || "";

  // Classifica SINTEGRA (heurística com companyDocFinal)
  const taggedS = tagSintegraDirection(SRows, companyDocFinal);
  const S = SRows.map(r => ({
    origem: "SINTEGRA",
    serie: normSerie(r.serie),
    nNF: normNF(r.nNF),
    data: /^\d{8}$/.test(r.data) ? r.data : toYMD(r.data),
    vNF: r.vNF,
    emit: { name: "—", doc: "", uf: "" },
    dest: { name: "—", doc: "", uf: "" },
    _tipo: r._tipo,
    raw: r
  }));

  // Normaliza XML
  const XN = X.map(x => ({
    origem: "XML",
    serie: normSerie(x.serie),
    nNF: normNF(x.nNF),
    data: toYMD(x.dhEmi),
    vNF: x.vNF,
    emit: x.emit, dest: x.dest,
    _tipo: x._tipo,
    raw: x
  }));

  // Index por chave (serie|nNF)
  const key = (r) => `${r.serie}|${r.nNF}`;
  const mapS = new Map(), mapX = new Map();
  for (const r of S) if (r.serie && r.nNF) mapS.set(key(r), r);
  for (const r of XN) if (r.serie && r.nNF) mapX.set(key(r), r);

  const common = [];
  const onlyS  = [];
  const onlyX  = [];

  // Em comum (mesmo série+NF e valor/data coincidentes)
  for (const [k, s] of mapS) {
    const x = mapX.get(k);
    if (!x) continue;
    const sameDate = (s.data && x.data) ? (s.data === x.data) : true;
    const sameVal  = sameValue(s.vNF, x.vNF, 0.02);
    if (sameDate && sameVal) {
      common.push({
        origem: "COMUM",
        serie: s.serie, nNF: s.nNF,
        data: s.data || x.data, vNF: s.vNF || x.vNF,
        emit: x.emit, dest: x.dest,
        _tipo: x._tipo
      });
    }
  }
  for (const [k, s] of mapS) {
    if (!mapX.has(k) || !common.find(r => r.serie === s.serie && r.nNF === s.nNF)) {
      onlyS.push(s);
    }
  }
  for (const [k, x] of mapX) {
    if (!mapS.has(k) || !common.find(r => r.serie === x.serie && r.nNF === x.nNF)) {
      onlyX.push(x);
    }
  }

  // UI
  $("results").classList.remove("hidden");
  $("stats").innerHTML = `
    <div>NFe XML (mod 55): <br><strong>${XN.length}</strong><br><small>Entradas: ${xmlEntradas.length} • Saídas: ${xmlSaidas.length}</small></div>
    <div>Registros SINTEGRA 50: <br><strong>${S.length}</strong><br><small>Entradas: ${taggedS.entradas.length} • Saídas: ${taggedS.saidas.length} • Desconh.: ${taggedS.desconhe.length}</small></div>
    <div>Em comum: <br><strong>${common.length}</strong></div>
    <div>Diferenças: <br><strong>${onlyS.length + onlyX.length}</strong></div>
    <div>CNPJ Empresa: <br><strong>${companyDocFinal || "—"}</strong></div>
  `;

  const headersCommon = ["Tipo","Série","NF","Emissão","Valor","Emitente","Destinatário"];
  const rowsCommon = common.map(r => [
    r._tipo || "—", r.serie, r.nNF, ymdToBR(r.data || ""), moneyBR(r.vNF || ""),
    r.emit?.name || "—", r.dest?.name || "—"
  ]);
  $("tableCommon").innerHTML = tableFromRows(headersCommon, rowsCommon);

  const headersOnly = ["Origem","Tipo","Série","NF","Emissão","Valor"];
  $("tableOnlySintegra").innerHTML = tableFromRows(headersOnly, onlyS.map(r => [
    "SINTEGRA", r._tipo || "—", r.serie, r.nNF, ymdToBR(r.data || ""), moneyBR(r.vNF || "")
  ]));
  $("tableOnlyXML").innerHTML = tableFromRows(headersOnly, onlyX.map(r => [
    "XML", r._tipo || "—", r.serie, r.nNF, ymdToBR(r.data || ""), moneyBR(r.vNF || "")
  ]));

  // Ações
  const dlExcel = $("dlExcel");
  const dlPDF   = $("dlPDF");
  const viewDanfeBtn = $("viewDanfeBtn");

  dlExcel.disabled = dlPDF.disabled = viewDanfeBtn.disabled = false;

  dlExcel.onclick = () => {
    const blob = buildExcelSheets(common, onlyS, onlyX);
    saveAs(blob, "sintegra_xml_relatorio.xlsx");
  };

  dlPDF.onclick = () => {
    const diffs = [
      ...onlyS.map(r => ({ ...r, origem: "SÓ SINTEGRA" })),
      ...onlyX.map(r => ({ ...r, origem: "SÓ XML" }))
    ];
    // meta baseada nas NFe (se tivermos companyDocFinal, tenta achar nome)
    const any = XN.find(xx => (xx.emit?.doc === companyDocFinal) || (xx.dest?.doc === companyDocFinal)) || XN[0];
    const meta = {
      companyName: (any && (any.emit?.doc === companyDocFinal ? any.emit?.name : any.dest?.name)) || "",
      companyDoc: companyDocFinal
    };
    openClientPDFPreview(meta, diffs);
  };

  viewDanfeBtn.onclick = () => {
    showDanfeModal(onlyX, X); // DANFE só faz sentido para o lado XML
  };

  $("progressBar").classList.add("hidden");
  window.__cmp = { X, XN, S, common, onlyS, onlyX, companyDocFinal };
}

// Modal DANFE — lista cartões com botão "Abrir DANFE"
function danfeCardHTML(x) {
  return `
    <div class="danfe-card">
      <h4>NF ${x.nNF} — Série ${x.serie}</h4>
      <div style="font-size:13px;color:#555;margin-bottom:8px;">
        Emissão: ${ymdToBR(toYMD(x.dhEmi))} • Valor: ${moneyBR(x.vNF)}
      </div>
      <div class="row" style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn primary" data-key="${x.key}">Abrir DANFE</button>
        <a class="btn secondary" target="_blank" rel="noopener"
           href="https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g%3D">Consultar na SEFAZ</a>
      </div>
    </div>
  `;
}
function showDanfeModal(onlyXNorm, XRaw) {
  const modal = $("danfeModal");
  const listEl = $("danfeList");
  if (!onlyXNorm || !onlyXNorm.length) { alert("Não há notas do lado XML para visualizar."); return; }

  const byKey = new Map();
  for (const x of XRaw) if (x.key) byKey.set(x.key, x);

  const cards = onlyXNorm.map(nx => {
    const full = byKey.get(nx.raw.key) || nx.raw;
    nx.__full = full; // cache
    return danfeCardHTML(full);
  }).join("");

  listEl.innerHTML = `<div>${cards}</div>`;
  modal.classList.remove("hidden");

  listEl.querySelectorAll("button[data-key]").forEach(btn => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-key");
      const rec = onlyXNorm.find(o => (o.raw.key === key));
      const full = rec?.__full || (window.__cmp?.X || []).find(xx => xx.key === key);
      if (full) openDANFEWindow(full);
    });
  });

  $("closeDanfe").onclick = () => modal.classList.add("hidden");
  modal.addEventListener("click", (e) => { if (e.target.id === "danfeModal") modal.classList.add("hidden"); });
}

// =========================================
// Boot
// =========================================
document.addEventListener("DOMContentLoaded", () => {
  setBranding();

  const boxS = $("boxSintegra");
  const boxX = $("boxXML");
  const inpS = $("filesSintegra");
  const inpX = $("filesXML");

  wireUploadUI(boxS, inpS, "SINTEGRA (.txt)");
  wireUploadUI(boxX, inpX, "XMLs / ZIPs de XML");

  $("compareBtn").addEventListener("click", async () => {
    try {
      $("compareBtn").disabled = true; $("compareBtn").textContent = "Comparando...";
      await compareAll(inpS.files[0] || null, inpX.files);
    } catch (e) {
      console.error(e);
      $("warn").textContent = e?.message || "Erro ao comparar.";
    } finally {
      $("compareBtn").disabled = false; $("compareBtn").textContent = "Comparar";
    }
  });

  $("resetBtn").addEventListener("click", () => {
    inpS.value = ""; inpX.value = "";
    $("warn").textContent = "";
    $("results").classList.add("hidden");
    [boxS, boxX].forEach(labelEl => {
      const span = labelEl.querySelector("span");
      labelEl.classList.remove("filled", "dragover");
      span.innerHTML = (labelEl.id === "boxSintegra")
        ? "<strong>SINTEGRA (.txt)</strong><br><small>Arraste aqui ou clique</small>"
        : "<strong>XMLs / ZIPs de XML</strong><br><small>Entradas e saídas</small>";
    });
  });
});
