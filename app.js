/* =========================================================
   PADRONIZADOR DE EXTRATOS GT — app.js v4
   Correção crítica:
   - Alguns XLSX do Itaú vêm com a dimensão interna da aba errada,
     por exemplo !ref = A1:F13, embora existam linhas até 4927.
   - O SheetJS respeita esse !ref e por isso lia só 13 linhas.
   - Esta versão recalcula o range real da planilha antes de converter.
========================================================= */

let RAW_DATA = [];
let PIVOT_DATA = [];
let MONTHS = [];
let SELECTED_FILE = null;

document.addEventListener("DOMContentLoaded", initApp);

function $(id){
  return document.getElementById(id);
}

function initApp(){
  const fileInput = $("fileInput");
  const processBtn = $("processBtn");
  const downloadBtn = $("downloadBtn");
  const dropzone = $("dropzone");

  if(!fileInput || !processBtn || !downloadBtn){
    console.error("Elementos obrigatórios não encontrados no HTML.");
    return;
  }

  setStatus("Pronto. Selecione ou arraste um arquivo para iniciar.");

  fileInput.addEventListener("change", () => {
    SELECTED_FILE = fileInput.files[0] || null;
    setFileName(SELECTED_FILE);

    if(SELECTED_FILE){
      setStatus("Arquivo selecionado. Clique em Processar arquivo.");
    }
  });

  processBtn.addEventListener("click", async () => {
    await processFile();
  });

  downloadBtn.addEventListener("click", exportWorkbook);

  if(dropzone){
    ["dragenter","dragover"].forEach(evt => {
      dropzone.addEventListener(evt, e => {
        e.preventDefault();
        dropzone.classList.add("drag");
      });
    });

    ["dragleave","drop"].forEach(evt => {
      dropzone.addEventListener(evt, e => {
        e.preventDefault();
        dropzone.classList.remove("drag");
      });
    });

    dropzone.addEventListener("drop", e => {
      const file = e.dataTransfer.files[0];

      if(!file) return;

      SELECTED_FILE = file;
      fileInput.files = e.dataTransfer.files;
      setFileName(file);
      setStatus("Arquivo arrastado. Clique em Processar arquivo.");
    });
  }
}

function setFileName(file){
  const el = $("fileName");

  if(el){
    el.textContent = file ? file.name : "Nenhum arquivo selecionado";
  }
}

/* =========================================================
   PROCESSAMENTO PRINCIPAL
========================================================= */

async function processFile(){
  try{
    clearError();

    const fileInput = $("fileInput");
    const downloadBtn = $("downloadBtn");

    const file = SELECTED_FILE || (fileInput && fileInput.files ? fileInput.files[0] : null);

    if(!file){
      setError("Nenhum arquivo selecionado. Selecione o arquivo Excel/CSV antes de processar.");
      return;
    }

    if(downloadBtn) downloadBtn.disabled = true;

    setStatus("Processando arquivo... lendo planilha.");

    const rows = await readFileToRows(file);

    setStatus(`Arquivo lido com sucesso. ${rows.length.toLocaleString("pt-BR")} linhas encontradas. Padronizando...`);

    const parsed = parseItauRows(rows);

    RAW_DATA = parsed.raw;
    PIVOT_DATA = parsed.pivot;
    MONTHS = parsed.months;

    if(!RAW_DATA.length){
      throw new Error("Nenhuma linha válida foi encontrada. Verifique se a aba contém Data, Descrição/Lançamento e Valor.");
    }

    renderRaw();
    renderPivot();
    renderKpis();

    if(downloadBtn) downloadBtn.disabled = false;

    setStatus(
      `Processado com sucesso: ${RAW_DATA.length.toLocaleString("pt-BR")} lançamentos, ` +
      `${PIVOT_DATA.length.toLocaleString("pt-BR")} denominações e ` +
      `${MONTHS.length.toLocaleString("pt-BR")} meses.`
    );

  }catch(err){
    console.error(err);
    setError(err.message || String(err));
  }
}

/* =========================================================
   LEITURA DO ARQUIVO
========================================================= */

async function readFileToRows(file){
  const buffer = await file.arrayBuffer();
  const name = String(file.name || "").toLowerCase();

  if(name.endsWith(".csv")){
    const text = new TextDecoder("utf-8").decode(buffer);
    return parseCsv(text);
  }

  if(typeof XLSX === "undefined"){
    throw new Error("Biblioteca XLSX não carregou. Verifique se o index.html possui o script CDN do xlsx.");
  }

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: false
  });

  const sheetName = pickBestSheetName(workbook.SheetNames);

  if(!sheetName){
    throw new Error("Não encontrei nenhuma aba válida no arquivo.");
  }

  setStatus(`Lendo aba: ${sheetName}`);

  const worksheet = workbook.Sheets[sheetName];

  // CORREÇÃO CRÍTICA:
  // O XLSX do Itaú pode vir com worksheet["!ref"] errado, ex: A1:F13,
  // mesmo existindo células até A4927:F4927.
  // Esta função recalcula o range real baseado nas células existentes.
  fixWorksheetRefByExistingCells(worksheet);

  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });
}

function fixWorksheetRefByExistingCells(ws){
  const refs = Object.keys(ws).filter(k => k[0] !== "!");
  if(!refs.length) return;

  let minR = Infinity;
  let minC = Infinity;
  let maxR = 0;
  let maxC = 0;

  refs.forEach(ref => {
    try{
      const cell = XLSX.utils.decode_cell(ref);

      if(cell.r < minR) minR = cell.r;
      if(cell.c < minC) minC = cell.c;
      if(cell.r > maxR) maxR = cell.r;
      if(cell.c > maxC) maxC = cell.c;
    }catch(e){
      // ignora chaves que não sejam célula
    }
  });

  if(Number.isFinite(minR) && Number.isFinite(minC)){
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: minR, c: minC },
      e: { r: maxR, c: maxC }
    });
  }
}

function pickBestSheetName(sheetNames){
  if(!sheetNames || !sheetNames.length) return null;

  const lanc = sheetNames.find(s => {
    const n = normalizeHeader(s);
    return n.includes("lancamento") || n.includes("lancamentos");
  });

  return lanc || sheetNames[0];
}

/* =========================================================
   CSV
========================================================= */

function parseCsv(text){
  text = String(text || "").replace(/^\uFEFF/, "");

  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for(let i = 0; i < text.length; i++){
    const ch = text[i];

    if(ch === '"'){
      if(inQuotes && text[i + 1] === '"'){
        field += '"';
        i++;
      }else{
        inQuotes = !inQuotes;
      }
    }else if((ch === "," || ch === ";") && !inQuotes){
      row.push(field);
      field = "";
    }else if((ch === "\n" || ch === "\r") && !inQuotes){
      if(ch === "\r" && text[i + 1] === "\n") i++;

      row.push(field);

      if(row.some(v => norm(v) !== "")){
        rows.push(row);
      }

      row = [];
      field = "";
    }else{
      field += ch;
    }
  }

  if(field.length || row.length){
    row.push(field);

    if(row.some(v => norm(v) !== "")){
      rows.push(row);
    }
  }

  return rows;
}

/* =========================================================
   PARSER ITAÚ
   Layout fixo usado para seu extrato:
   A = Data
   B = Lançamento
   C = Razão Social
   D = CPF/CNPJ
   E = Valor
   F = Saldo
========================================================= */

function parseItauRows(rows){
  if(!rows || !rows.length){
    throw new Error("A planilha está vazia.");
  }

  const raw = [];

  for(let i = 0; i < rows.length; i++){
    const r = rows[i];

    // Estrutura fixa do Itaú:
    // A = Data
    // B = Lançamento
    // E = Valor da operação
    const dataRaw = r[0];
    const descRaw = r[1];
    const valorRaw = r[4];

    const descricao = normalizeDescription(descRaw);

    // Pula cabeçalho/metadados/linhas vazias
    if(!descricao) continue;

    // Pula linhas de saldo e resumo, mas CONTINUA a leitura
    if(shouldIgnoreDescription(descricao)){
      continue;
    }

    const data = parseDate(dataRaw);
    const valorOriginal = parseValue(valorRaw);

    if(!data) continue;
    if(valorOriginal === null || valorOriginal === undefined || Number.isNaN(valorOriginal)) continue;

    const valor = Math.abs(valorOriginal);

    if(!valor) continue;

    raw.push({
      Data: formatDateBR(data),
      "Denominação": descricao,
      Valor: round2(valor),
      Mes: toYM(data)
    });
  }

  const pivot = buildPivot(raw);
  const months = getMonths(raw);

  return { raw, pivot, months };
}

function shouldIgnoreDescription(descricao){
  const d = normalizeDescription(descricao);

  return (
    d.includes("SALDO") ||
    d.includes("TOTAL DISPON") ||
    d.includes("SALDO ANTERIOR") ||
    d.includes("SALDO EM CONTA") ||
    d.includes("LIMITE") ||
    d.includes("RESUMO") ||
    d.includes("LANCAMENTOS") ||
    d.includes("LANÇAMENTOS") ||
    d.includes("PERIODO") ||
    d.includes("PERÍODO") ||
    d.includes("ATUALIZACAO") ||
    d.includes("ATUALIZAÇÃO") ||
    d.includes("AGENCIA") ||
    d.includes("AGÊNCIA") ||
    d.includes("CONTA")
  );
}

/* =========================================================
   HELPERS
========================================================= */

function normalizeHeader(v){
  return String(v ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function norm(v){
  return String(v ?? "").replace(/\u00A0/g, " ").trim();
}

function parseDate(v){
  if(!v) return null;

  if(v instanceof Date && !isNaN(v)) return v;

  const s = norm(v);

  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(m){
    let y = Number(m[3]);
    if(y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }

  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m){
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const d = new Date(s);
  if(!isNaN(d)) return d;

  return null;
}

function parseValue(v){
  if(v === null || v === undefined || v === "") return null;

  if(typeof v === "number"){
    return Number.isFinite(v) ? v : null;
  }

  let s = norm(v);
  if(!s || s === "-" || s === "—") return null;

  let negative = false;

  if(/^\(.*\)$/.test(s)){
    negative = true;
    s = s.slice(1, -1);
  }

  if(s.includes("-")) negative = true;

  s = s
    .replace(/R\$/g, "")
    .replace(/\s/g, "")
    .replace(/[^\d,.-]/g, "");

  // Padrão BR: 13.529,60
  if(s.includes(",") && s.includes(".")){
    s = s.replace(/\./g, "").replace(",", ".");
  }
  // Padrão BR simples: 857,70
  else if(s.includes(",") && !s.includes(".")){
    s = s.replace(",", ".");
  }
  // Padrão XLSX/EN: 857.70 — mantém ponto decimal

  s = s.replace(/-/g, "");

  const n = Number(s);
  if(!Number.isFinite(n)) return null;

  return negative ? -n : n;
}

function normalizeDescription(v){
  return norm(v)
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function toYM(d){
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function formatDateBR(d){
  return d.toLocaleDateString("pt-BR");
}

function round2(n){
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

/* =========================================================
   PIVOT
========================================================= */

function getMonths(raw){
  return [...new Set(raw.map(r => r.Mes))].sort();
}

function buildPivot(raw){
  const months = getMonths(raw);
  const map = new Map();

  raw.forEach(r => {
    const denom = r["Denominação"];

    if(!map.has(denom)){
      map.set(denom, {});
    }

    const obj = map.get(denom);
    obj[r.Mes] = (obj[r.Mes] || 0) + r.Valor;
  });

  return [...map.entries()]
    .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"))
    .map(([denom, values]) => {
      const obj = { "Denominação": denom };
      months.forEach(m => obj[m] = round2(values[m] || 0));
      return obj;
    });
}

/* =========================================================
   RENDERIZAÇÃO
========================================================= */

function renderRaw(){
  const rows = RAW_DATA.slice(0, 100);
  const cols = ["Data", "Denominação", "Valor", "Mes"];
  const el = $("rawPreview");

  if(el){
    el.innerHTML = tableHtml(rows, cols);
  }
}

function renderPivot(){
  const rows = PIVOT_DATA.slice(0, 100);
  const el = $("pivotPreview");

  if(!el) return;

  if(!rows.length){
    el.innerHTML = "";
    return;
  }

  const cols = Object.keys(rows[0]);
  el.innerHTML = tableHtml(rows, cols);
}

function tableHtml(rows, cols){
  let html = "<table><thead><tr>";

  cols.forEach(c => {
    html += `<th>${escapeHtml(c)}</th>`;
  });

  html += "</tr></thead><tbody>";

  rows.forEach(r => {
    html += "<tr>";

    cols.forEach(c => {
      const isNum = typeof r[c] === "number";
      html += `<td class="${isNum ? "num" : ""}">${escapeHtml(formatCell(r[c]))}</td>`;
    });

    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function formatCell(v){
  if(typeof v === "number"){
    return v.toLocaleString("pt-BR", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  return String(v ?? "");
}

function escapeHtml(v){
  return String(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderKpis(){
  const total = RAW_DATA.reduce((acc, r) => acc + r.Valor, 0);

  if($("rawCount")) $("rawCount").textContent = RAW_DATA.length.toLocaleString("pt-BR");
  if($("denomCount")) $("denomCount").textContent = PIVOT_DATA.length.toLocaleString("pt-BR");
  if($("monthCount")) $("monthCount").textContent = MONTHS.length.toLocaleString("pt-BR");

  if($("totalValue")){
    $("totalValue").textContent = total.toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  }
}

/* =========================================================
   EXPORTAÇÃO
========================================================= */

function exportWorkbook(){
  if(!RAW_DATA.length || !PIVOT_DATA.length){
    alert("Processe um arquivo antes de exportar.");
    return;
  }

  if(typeof XLSX === "undefined"){
    alert("Biblioteca XLSX não carregou.");
    return;
  }

  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(RAW_DATA),
    "RAW_EXTRATO"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(PIVOT_DATA),
    "PIVOT_MENSAL"
  );

  XLSX.writeFile(wb, "extrato_padronizado_RAW_PIVOT.xlsx");
}

/* =========================================================
   STATUS / ERROS
========================================================= */

function setStatus(msg){
  const box = document.querySelector(".statusCard");
  if(box) box.classList.remove("error");

  const el = $("status");
  if(el) el.textContent = msg;
}

function setError(msg){
  const box = document.querySelector(".statusCard");
  if(box) box.classList.add("error");

  const el = $("status");
  if(el) el.textContent = msg;
}

function clearError(){
  const box = document.querySelector(".statusCard");
  if(box) box.classList.remove("error");
}
