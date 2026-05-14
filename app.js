/* CÓDIGO AJUSTADO 14052026  */

const $ = (id) => document.getElementById(id);

let RAW_DATA = [];
let PIVOT_DATA = [];
let MONTHS = [];
let SELECTED_FILE = null;

const fileInput = $("fileInput");
const dropzone = $("dropzone");

/* ESCOLHA DO ARQUIVO A SER PADRONIZADO */

fileInput.addEventListener("change", () => {
  SELECTED_FILE = fileInput.files[0] || null;
  $("fileName").textContent = SELECTED_FILE ? SELECTED_FILE.name : "Nenhum arquivo selecionado";
});

$("processBtn").addEventListener("click", processFile);
$("downloadBtn").addEventListener("click", exportWorkbook);

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
  if (!file) return;
  SELECTED_FILE = file;
  fileInput.files = e.dataTransfer.files;
  $("fileName").textContent = file.name;
});

/* --------- >  PROCESSAMENTO DO ARQUIVO  */

async function processFile(){
  try{
    clearError();
    const file = SELECTED_FILE || fileInput.files[0];
    if(!file){
      setError("Selecione um arquivo antes de processar.");
      return;
    }

    $("downloadBtn").disabled = true;
    setStatus("Lendo arquivo...");

    const rows = await readFileToRows(file);
    setStatus(`Arquivo lido. ${rows.length} linhas encontradas. Detectando cabeçalho...`);

    const parsed = parseItauRows(rows);
    RAW_DATA = parsed.raw;
    PIVOT_DATA = parsed.pivot;
    MONTHS = parsed.months;

    if(!RAW_DATA.length){
      throw new Error("Nenhuma linha válida encontrada. Verifique se o extrato possui Data, Lançamento/Descrição e Valor.");
    }

    renderRaw();
    renderPivot();
    renderKpis();

    $("downloadBtn").disabled = false;
    setStatus(`Processado com sucesso: ${RAW_DATA.length} lançamentos, ${PIVOT_DATA.length} denominações e ${MONTHS.length} meses.`);
  }catch(err){
    console.error(err);
    setError(err.message || String(err));
  }
}

async function readFileToRows(file){
  const buffer = await file.arrayBuffer();
  const name = file.name.toLowerCase();

  if(name.endsWith(".csv")){
    const text = new TextDecoder("utf-8").decode(buffer);
    return parseCsv(text);
  }

  const wb = XLSX.read(buffer, {type:"array",cellDates:true,raw:false});
  const ws = wb.Sheets[wb.SheetNames[0]];

  return XLSX.utils.sheet_to_json(ws, {
    header:1,
    defval:"",
    raw:false,
    blankrows:false
  });
}

function parseCsv(text){
  text = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [], field = "", inQuotes = false;

  for(let i=0;i<text.length;i++){
    const ch = text[i];
    if(ch === '"'){
      if(inQuotes && text[i+1] === '"'){ field += '"'; i++; }
      else inQuotes = !inQuotes;
    }else if((ch === "," || ch === ";") && !inQuotes){
      row.push(field); field = "";
    }else if((ch === "\n" || ch === "\r") && !inQuotes){
      if(ch === "\r" && text[i+1] === "\n") i++;
      row.push(field);
      if(row.some(v => norm(v) !== "")) rows.push(row);
      row = []; field = "";
    }else{
      field += ch;
    }
  }

  if(field.length || row.length){
    row.push(field);
    if(row.some(v => norm(v) !== "")) rows.push(row);
  }
  return rows;
}

/* -----> FUNÇÃO PARSE ATUALIZADA   */

function parseItauRows(rows){

  const headerInfo = findHeaderRow(rows);

  if(!headerInfo){
    throw new Error(
      "Não encontrei cabeçalho válido."
    );
  }

  const { headerIndex, map } = headerInfo;

  const raw = [];

  for(let i = headerIndex + 1; i < rows.length; i++){

    const r = rows[i];

    const dataRaw = getCell(r, map.data);
    const descRaw = getCell(r, map.descricao);
    const valorRaw = getCell(r, map.valor);

    const data = parseDate(dataRaw);

    let descricao = normalizeDescription(descRaw);

    const valor = parseValue(valorRaw);

    // ignora linhas sem data
    if(!data) continue;

    // ignora linhas sem descrição
    if(!descricao) continue;

    // ignora saldo
    if(
      descricao.includes("SALDO") ||
      descricao.includes("TOTAL DISPON")
    ){
      continue;
    }

    // ignora linhas sem valor numérico
    if(
      valor === null ||
      valor === undefined ||
      isNaN(valor)
    ){
      continue;
    }

    raw.push({
      Data: formatDateBR(data),
      "Denominação": descricao,
      Valor: round2(Math.abs(valor)),
      Mes: toYM(data)
    });
  }

  const pivot = buildPivot(raw);

  const months = getMonths(raw);

  return {
    raw,
    pivot,
    months
  };
}

/* -----------> FUNÇÃO PARSE VALUES     */

function parseValue(v){

  if(
    v === null ||
    v === undefined ||
    v === ""
  ){
    return null;
  }

  // número nativo excel
  if(typeof v === "number"){

    if(isNaN(v)) return null;

    return v;
  }

  let s = String(v)
    .trim()
    .replace(/R\\$/g,"")
    .replace(/\\s/g,"");

  if(!s) return null;

  let negative = false;

  if(s.includes("-")){
    negative = true;
  }

  s = s
    .replace(/\\./g,"")
    .replace(",",".")
    .replace(/-/g,"");

  const n = Number(s);

  if(isNaN(n)) return null;

  return negative ? -n : n;
}

/* --------> FUNÇÃO DE NORMATIUZAÇÃO DA DESCRIÇÃO     */

function normalizeDescription(v){
  return norm(v).replace(/\s+/g," ").trim().toUpperCase();
}

function toYM(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}

function formatDateBR(d){
  return d.toLocaleDateString("pt-BR");
}

function round2(n){
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function getMonths(raw){
  return [...new Set(raw.map(r => r.Mes))].sort();
}

function buildPivot(raw){
  const months = getMonths(raw);
  const map = new Map();

  raw.forEach(r => {
    const denom = r["Denominação"];
    if(!map.has(denom)) map.set(denom, {});
    const obj = map.get(denom);
    obj[r.Mes] = (obj[r.Mes] || 0) + r.Valor;
  });

  return [...map.entries()]
    .sort((a,b) => a[0].localeCompare(b[0], "pt-BR"))
    .map(([denom, values]) => {
      const obj = { "Denominação": denom };
      months.forEach(m => obj[m] = round2(values[m] || 0));
      return obj;
    });
}

function renderRaw(){
  const rows = RAW_DATA.slice(0, 80);
  const cols = ["Data", "Denominação", "Valor", "Mes"];
  $("rawPreview").innerHTML = tableHtml(rows, cols);
}

function renderPivot(){
  const rows = PIVOT_DATA.slice(0, 80);
  if(!rows.length){
    $("pivotPreview").innerHTML = "";
    return;
  }
  const cols = Object.keys(rows[0]);
  $("pivotPreview").innerHTML = tableHtml(rows, cols);
}

function tableHtml(rows, cols){
  let html = "<table><thead><tr>";
  cols.forEach(c => html += `<th>${escapeHtml(c)}</th>`);
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
  if(typeof v === "number") return v.toLocaleString("pt-BR", {minimumFractionDigits:2, maximumFractionDigits:2});
  return String(v ?? "");
}

function escapeHtml(v){
  return String(v)
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

function renderKpis(){
  const total = RAW_DATA.reduce((acc, r) => acc + r.Valor, 0);
  $("rawCount").textContent = RAW_DATA.length.toLocaleString("pt-BR");
  $("denomCount").textContent = PIVOT_DATA.length.toLocaleString("pt-BR");
  $("monthCount").textContent = MONTHS.length.toLocaleString("pt-BR");
  $("totalValue").textContent = total.toLocaleString("pt-BR", {style:"currency", currency:"BRL"});
}

function exportWorkbook(){
  if(!RAW_DATA.length || !PIVOT_DATA.length){
    alert("Processe um arquivo antes de exportar.");
    return;
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(RAW_DATA), "RAW_EXTRATO");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(PIVOT_DATA), "PIVOT_MENSAL");
  XLSX.writeFile(wb, "extrato_padronizado_RAW_PIVOT.xlsx");
}

function setStatus(msg){
  document.querySelector(".statusCard").classList.remove("error");
  $("status").textContent = msg;
}

function setError(msg){
  document.querySelector(".statusCard").classList.add("error");
  $("status").textContent = msg;
}

function clearError(){
  document.querySelector(".statusCard").classList.remove("error");
}
