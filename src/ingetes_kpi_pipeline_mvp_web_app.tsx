import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

// ====== Metas (Sheets) ======

type MetaRecord = {
  comercial: string;
  metaAnual: number;
  metaOfertas: number;
  metaVisitas: number;
};
type MetasResponse = { year: number; metas: MetaRecord[] };

// === Metas por comercial (Google Apps Script) ===
// GET público (el que te abre el JSON en el navegador)
const METAS_GET_URL = 'https://script.googleusercontent.com/macros/echo?user_content_key=AehSKLgYKi_vgRuH8EXtITmadj95wNjzUtI-IrgFtqVvQQm5Vkh6dNCYhtkyNDakA9TrdZo93Wp7umg0Mv_ifW6MEJxqmHHj36hscjlNFAYHBGCcky4NCy_DdYGL9Sk4h3MIg6dcL2R5t-LQTZnkxZ7vAj4Du1z6_DQz8km5U82Qj6Bj08n4l43iTUJr4omgCgPWI6M8idwJcx52QgULieG8HHqaamcztFr9cbH3PwOF9-BWMfJbDHr2EOESEpsEkZJqvWMdZ6LNIose2TikLaOVIowGZdIqZfMZ_e4hhk_v&lib=M_wwqKoJEkvpNDcscez3XfhqGZ0szE9sQ';

// POST (deployment /exec público) para guardar cambios
const METAS_POST_URL =   'https://script.google.com/macros/s/AKfycbxHfSCDgGArZFPLXzKmsw11rpayGLA33fX3kQGxWcgUcv_ymOW5cmgv3DupKBCxHMlzYA/exec'; // <-- pega tu URL /exec

// Debe coincidir con la clave del Apps Script (getApiKey → 'INGETES' por defecto)
const METAS_API_KEY = 'INGETES';

async function guardarMetas(year, metas){
  const res = await fetch(METAS_POST_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' }, // o 'text/plain' también sirve
    body: JSON.stringify({ apiKey: METAS_API_KEY, year, metas })
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'error');
  return true;
}

// Normaliza nombres para que coincidan con los de los archivos
function normName(s: string) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

async function fetchMetas(year: number): Promise<MetasResponse> {
  // añade ?year= o &year= según corresponda
  const url = `${METAS_GET_URL}${METAS_GET_URL.includes("?") ? "&" : "?"}year=${year}`;
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`GET metas ${res.status}`);
  const data = await res.json();

  const metas = Array.isArray(data?.metas) ? data.metas : [];

  // normaliza y fuerza números
  const parsed: MetaRecord[] = metas.map((m: any) => ({
    comercial: normName(m.comercial),
    metaAnual: Number(m.metaAnual ?? 0),
    metaOfertas: Number(m.metaOfertas ?? 0),
    metaVisitas: Number(m.metaVisitas ?? 0),
  }));

  return { year: Number(data?.year) || year, metas: parsed };
}

async function fetchMetasFromSheet(year: number) {
  const url = `${METAS_GET_URL}${METAS_GET_URL.includes('?') ? '&' : '?'}year=${year}`;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`GET metas ${res.status}`);
  const data = await res.json();

  // Siempre devuelve array
  const metas: Array<{
    year: number;
    comercial: string;
    metaAnual: number | null;
    metaOfertas: number | null;
    metaVisitas: number | null;
  }> = Array.isArray(data?.metas) ? data.metas : [];

  // Normaliza y fuerza números (0 si viene null o string)
  return metas.map(m => ({
    year: Number(m.year),
    comercial: normName(m.comercial),
    metaAnual: Number(m.metaAnual ?? 0),
    metaOfertas: Number(m.metaOfertas ?? 0),
    metaVisitas: Number(m.metaVisitas ?? 0),
  }));
}

async function saveMetas(year: number, metas: any[]) {
  try {
    const res = await fetch(METAS_POST_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: METAS_API_KEY, year, metas }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error);
    alert("Metas guardadas correctamente ✅");
  } catch (err) {
    alert("Error guardando metas: " + err);
  }
}



async function saveMetasToSheet(year: number, filas: Array<{ comercial: string; metaAnual: number; metaOfertas: number; metaVisitas: number }>) {
  const res = await fetch(METAS_POST_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      apiKey: METAS_API_KEY,
      year,
      metas: filas.map(f => ({
        comercial: f.comercial,
        metaAnual: Number(f.metaAnual || 0),
        metaOfertas: Number(f.metaOfertas || 0),
        metaVisitas: Number(f.metaVisitas || 0),
      })),
    }),
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`POST metas ${res.status}: ${txt}`);
  }
  const out = await res.json();
  if (!out?.ok) throw new Error(out?.error || 'unknown error');
}

// ========================= Utils =========================
const norm = (s:any) => String(s ?? "")
  .normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
  .toLowerCase().replace(/[()]/g, " ").replace(/\s+/g, " ").trim();

// Columna/clave que representa totales del pivot (no contarla)
const isTotalLike = (key: string) => {
  const z = norm(key);
  return (
    z === "total" ||
    z === "__total__" ||
    z === "totalgeneral" ||
    z.includes("total")
  );
};

function calcOfferCountFromDetail(detailModel:any){
  const by = new Map<string, number>();
  const all = detailModel?.allRows || [];
  for (const r of all) {
    // ya viene filtrado sin subtotales y con arrastre de comercial
    by.set(r.comercial, (by.get(r.comercial) || 0) + 1);
  }
  const porComercial = Array.from(by.entries())
    .map(([comercial, count]) => ({ comercial, count }))
    .sort((a,b)=> b.count - a.count);

  const total = porComercial.reduce((a,x)=> a + x.count, 0);
  const max   = porComercial.length ? porComercial[0].count : 0;
  return { total, max, porComercial };
}

function parseExcelDate(v:any): Date|null {
  if (v == null || v === "") return null;
  if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
  if (typeof v === "number") {
    const d = (XLSX as any).SSF.parse_date_code(v);
    if (!d) return null;
    return new Date(Date.UTC(d.y, d.m - 1, d.d));
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m){
    const dd = +m[1], mm = +m[2]-1, yy = +m[3]; const y = yy < 100 ? 2000 + yy : yy;
    return new Date(Date.UTC(y, mm, dd));
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : new Date(Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate()));
}

const toNumber = (v:any) => {
  if (v == null || v === "") return 0;
  let s = String(v).trim().replace(/[^\d,.-]/g, "");
  if (s.includes(",") && !s.includes(".")) s = s.replace(/,/g, ".");
  if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, "");
  const n = Number(s);
  return isFinite(n) ? n : 0;
};

// --- Clasificación de eventos SOLO por Asunto ---
function classifyEventFromSubject(asunto?: string): "llamadas" | "visitas" | "reuniones" | "otros" {
  const s = norm(asunto || "");

  // llamadas: call / llamada / teléfono
  if (/(?:\bcall\b|llamad|telefon)/i.test(s)) return "llamadas";

  // reuniones: reunion / reunión / meeting / "reunion con" / "reunion cliente"
  if (/(?:\breunion|\breunión|\bmeeting|reunion con|reunion cliente|reunion comercial|reunion virtual)/i.test(s))
    return "reuniones";

  // visitas: visita / visit / "visita a" / "visita cliente" / "visita comercial" / "visita planta"
  if (/(?:\bvisita|\bvisit|visita a|visita cliente|visita comercial|visita planta)/i.test(s))
    return "visitas";

  return "otros";
}

// Palabras clave para detectar columnas
const OWNER_KEYS  = ["propietario","owner","comercial","vendedor","ejecutivo"];
const STAGE_KEYS  = ["etapa","stage","estado"];
const CREATE_KEYS = ["fecha de creacion","fecha creación","created","created date","fecha creacion"];
const CLOSE_KEYS  = ["fecha de cierre","fecha cierre","close date","fecha cierre real","fecha cierre oportunidad"];

// Expresiones para cerradas / ganadas / perdidas
const CLOSED_WON_RX  = /(closed\s*won|ganad|cerrad[oa].*ganad)/i;
const CLOSED_LOST_RX = /(closed\s*lost|perdid|cerrad[oa].*perdid)/i;
const CLOSED_RX      = /(closed\s*won|closed\s*lost|ganad|perdid|cerrad[oa])/i;

// Semáforo por cumplimiento vs meta
function offerStatus(count: number, target: number) {
  // Si la meta es 0, considerar cumplido (verde) cuando haya ≥0 ofertas
  const ratio = target > 0 ? (count / target) : (count > 0 ? 1 : 1);
  if (ratio >= 1) return { ratio, bar: "bg-green-500", dot: "bg-green-500", text: "text-green-600" };
  if (ratio >= 0.8) return { ratio, bar: "bg-yellow-400", dot: "bg-yellow-400", text: "text-yellow-600" };
  return { ratio, bar: "bg-red-500", dot: "bg-red-500", text: "text-red-600" };
}

// Tarjeta compacta para los 3 recuadros
const StatCard = ({ label, children }: { label: string; children: React.ReactNode }) => (
  <div className="p-3 bg-gray-100 rounded">
    <div className="text-xs text-gray-500">{label}</div>
    <div className="text-2xl font-bold tabular-nums text-gray-900">{children}</div>
  </div>
);

// --- Helper: filtra filas por comercial seleccionado ---
function onlySelected<T extends { comercial: string }>(arr: T[], selected: string): T[] {
  return selected === "ALL" ? arr : arr.filter(r => r.comercial === selected);
}

// normaliza nombres: quita tildes, compacta espacios y pone mayúsculas
const nameKey = (s: string) =>
  String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")  // sin tildes
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

const parseDateCell = (val: any) => {
  if (val == null || val === "") return null;
  // Excel serial
  if (typeof val === "number") {
    const d: any = (XLSX as any).SSF.parse_date_code(val);
    if (!d) return null;
    // fecha UTC sin hora
    return new Date(Date.UTC(d.y, (d.m || 1) - 1, d.d || 1));
  }

  // Texto normalizado
  const s = String(val).trim();

  // dd/mm/yyyy o dd-mm-yyyy
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3].length === 2 ? ("20"+m[3]) : m[3]);
    const d = new Date(Date.UTC(yy, mm - 1, dd));
    return isNaN(d.getTime()) ? null : d;
  }

  // yyyy-mm-dd (ISO) y similares
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : new Date(Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate()));
};

function daysBetween(a: Date, b: Date) {
  const MS = 24 * 3600 * 1000;
  return Math.max(0, Math.round((b.getTime() - a.getTime()) / MS));
}

const fmtCOP = (n: number) => n.toLocaleString("es-CO");
const OPEN_STAGES = ["prospect", "qualification", "negotiation", "proposal", "open", "nuevo", "calificacion", "negociacion", "propuesta","QUALIFICATION","NEEDS ANALYSIS", "VALUE PROPOSITION","PERCEPTION ANALYSIS",
  "ID. DECISION MAKERS","PROSPECTING","PROPOSAL/PRICE QUOTE","NEGOTIATION"];
const WON_STAGES = ["closed won", "ganad"]; // detectar 'Closed Won' y variantes en español
const normStage = (s?: string) => (s || '').trim().toUpperCase();
const isOpenStage = (stage?: string) => {
  const st = String(stage || "").trim().toUpperCase();
  return st && st !== "CLOSED WON" && st !== "CLOSED LOST";
};

const coverageColor = (p: number) => (p >= 1 ? 'text-green-600' : p >= 0.8 ? 'text-yellow-600' : 'text-red-600');
const coverageDot   = (p: number) => (p >= 1 ? '●' : p >= 0.8 ? '●' : '●');

// ================= Comerciales: lista fija + mapa =================
const FIXED_COMERCIALES = [
  "ALL",
  "CLAUDIA RODRIGUEZ RODRIGUEZ",
  "HERNAN ROLDAN",
  "JHOAN ORTIZ",
  "JUAN GARZÓN LINARES",
  "KAREN CARRILLO",
  "LIZETH MARTINEZ",
  "PABLO RODRIGUEZ RODRIGUEZ",
];

const comercialMap: Record<string, string> = {
  // Claudia
  "claudia rodriguez": "CLAUDIA RODRIGUEZ RODRIGUEZ",
  "claudia rodriguez rodriguez": "CLAUDIA RODRIGUEZ RODRIGUEZ",
  "claudia patricia rodriguez": "CLAUDIA RODRIGUEZ RODRIGUEZ",

  // Hernán
  "hernan roldan": "HERNAN ROLDAN",
  "hernan benancio roldan": "HERNAN ROLDAN",
  "hernan b roldan": "HERNAN ROLDAN",

  // Jhoan
  "jhoan ortiz": "JHOAN ORTIZ",
  "jhoan sebastian ortiz": "JHOAN ORTIZ",

  // Juan
  "juan garzon": "JUAN GARZÓN LINARES",
  "juan garzon linares": "JUAN GARZÓN LINARES",
  "juan garzón linares": "JUAN GARZÓN LINARES",
  "juan sebastian garzon": "JUAN GARZÓN LINARES",
  "juan sebastian garzon linares": "JUAN GARZÓN LINARES",

  // Karen
  "karen carrillo": "KAREN CARRILLO",
  "karen ariana carrillo": "KAREN CARRILLO",

  // Lizeth
  "lizeth martinez": "LIZETH MARTINEZ",
  "lizeth natalia martinez": "LIZETH MARTINEZ",

  // Pablo
  "pablo rodriguez rodriguez": "PABLO RODRIGUEZ RODRIGUEZ",
  "pablo cesar rodriguez": "PABLO RODRIGUEZ RODRIGUEZ",
};

const mapComercial = (raw: any) => {
  let s = String(raw ?? "").trim();
  if (!s) return ""; // ← vacío para que no mate el arrastre

  // Email → nombre
  if (s.includes("@")) s = s.split("@")[0].replace(/[._]/g, " ");
  // "APELLIDO, NOMBRE" → reordenar
  if (s.includes(",")) { const [a, b] = s.split(",").map(t => t.trim()); if (a && b) s = `${b} ${a}`; }

  const key = norm(s);
  if (comercialMap[key]) return comercialMap[key];

  // Coincidencia exacta contra lista fija
  const fixedExact = FIXED_COMERCIALES.find(c => norm(c) === key);
  if (fixedExact) return fixedExact;

  // Fuzzy simple por tokens
  const candidateSet = FIXED_COMERCIALES.filter(c => c !== "ALL");
  const tokens = key.split(" ").filter(Boolean);
  let best = ""; let bestScore = 0;
  for (const cand of candidateSet) {
    const ck = norm(cand);
    const ctoks = ck.split(" ").filter(Boolean);
    const inter = new Set(tokens.filter(t => ctoks.includes(t))).size;
    const union = new Set([...tokens, ...ctoks]).size || 1;
    const score = inter / union;
    if (score > bestScore) { bestScore = score; best = cand; }
  }
  if (bestScore >= 0.5) return best;

  // Último intento: contiene el apellido principal
  for (const cand of candidateSet) {
    const ck = norm(cand);
    const ap = ck.split(" ").slice(-1)[0];
    if (ap && key.includes(ap)) return cand;
  }

  return "(Sin comercial)";
};

// ================= Workbook robust reader =================
const looksZip = (u8: Uint8Array) => u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 0x03 || u8[2] === 0x05 || u8[2] === 0x07);
const looksOLE = (u8: Uint8Array) => u8.length >= 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0;

async function readWorkbookRobust(file: File) {
  const buf = await file.arrayBuffer();
  const u8 = new Uint8Array(buf);
  const looksZip = (u8:Uint8Array) => u8.length>=4 && u8[0]===0x50 && u8[1]===0x4b;
  const looksOLE = (u8:Uint8Array) => u8.length>=8 && u8[0]===0xd0 && u8[1]===0xcf;

  if (looksZip(u8) || looksOLE(u8)) {
    try { return XLSX.read(u8, { type: "array", dense: true }); }
    catch { /* fallback a string/base64 si hace falta */ }
  }
  try { const text = await file.text(); return XLSX.read(text, { type:"string", dense:true }); } catch {}
  const bin = Array.from(u8).map(b=>String.fromCharCode(b)).join("");
  const b64 = typeof btoa!=="undefined" ? btoa(bin) : "";
  if (b64) return XLSX.read(b64, { type:"base64", dense:true });
  throw new Error("No se pudo leer el archivo.");
}

function parseVisitsFromSheet(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];
  if (!A.length) throw new Error("VISITAS: hoja vacía");

  // Detectar fila de encabezados
  const scoreHead = (row:any[]) => {
    const H = row.map(norm);
    let sc = 0;
    if (H.some(h => h.includes("comercial") || h.includes("propietario") || h.includes("owner") || h.includes("vendedor") || h.includes("ejecutivo"))) sc++;
    if (H.some(h => h.includes("fecha") || h.includes("date"))) sc++;
    return sc;
  };
  let headerRow = 0, best = -1;
  for (let r = 0; r < Math.min(40, A.length); r++) {
    const sc = scoreHead(A[r]||[]);
    if (sc > best) { best = sc; headerRow = r; }
  }

  const headers = A[headerRow] || [];
  const findIdx = (...cands:string[]) => {
    const NC = cands.map(norm);
    for (let c=0; c<headers.length; c++) {
      const h = norm(headers[c]);
      if (NC.some(k => h.includes(k))) return c;
    }
    return -1;
  };

// === FORZAR B (col 1) para Comercial y J (col 9) para Asunto ===
// seguimos detectando fecha/cliente como opcionales por si existen
let idxFec = findIdx("fecha de visita","fecha visita","fecha","date","created","evento"); // puede quedar en -1
let idxCli = findIdx("cliente","account","empresa","compania","company","account name");

// Comercial = COLUMNA B (índice 1) sí o sí
let idxCom = 1;
// Asunto     = COLUMNA J (índice 9) sí o sí
let idxAsu = 9;

// Si los encabezados no están donde esperamos, igual forzamos,
// pero hacemos una verificación suave para ayudarte en el futuro:
if (!headers[idxCom]) console.warn("VISITAS: no veo encabezado en B, igual usaré B como Comercial");
if (!headers[idxAsu]) console.warn("VISITAS: no veo encabezado en J, igual usaré J como Asunto");

  if (idxCom < 0) {
    throw new Error(`VISITAS: falta columna de Comercial/Propietario en hoja ${sheetName}`);
  }
  // La fecha ya NO es obligatoria (idxFec puede ser -1)

  const rows:any[] = [];
  let currentComercial = "";   // ← arrastre del comercial

  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || [];
    if (row.every((v:any)=>String(v).trim()==="")) continue;

    // Saltar filas de agrupación: subtotal/total/recuento/suma, o encabezados repetidos
    const line = norm((row.join(" ")) || "");
    if (
      line.startsWith("subtotal") || line.startsWith("total") ||
      line.includes("recuento")   || line.includes("suma de") ||
      scoreHead(row) >= 2 // parece cabecera repetida
    ) continue;

    // 1) ¿Fila "título" de bloque? (valor SOLO en la columna comercial y lo demás vacío)
    const hasOnlyComercial =
      String(row[idxCom] ?? "").trim() !== "" &&
      row.filter((v:any, c:number) => c !== idxCom && String(v ?? "").trim() !== "").length === 0;

    if (hasOnlyComercial) {
      // actualizar comercial vigente y seguir (no es una visita)
      const mapped = mapComercial(row[idxCom]);
      if (mapped) currentComercial = mapped;
      continue;
    }

    // 2) Fila de datos: si trae comercial explícito, actualiza; si viene vacío, arrastra
    const rawCom = row[idxCom];
    if (rawCom != null && String(rawCom).trim() !== "") {
      const mapped = mapComercial(rawCom);
      if (mapped) currentComercial = mapped;
    }
    const comercial = currentComercial;
    if (!comercial) continue; // aún no hay comercial vigente → no contar

    // Fecha (opcional). Si no hay o no parsea, igual contamos el evento.
    let ym = "";
    if (idxFec >= 0) {
      const fecha = parseDateCell(row[idxFec]);
      if (fecha instanceof Date && !isNaN(fecha.getTime())) {
        const d = new Date(fecha);
        ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}`;
      }
    }
    
    const cliente = idxCli >= 0 ? String(row[idxCli] ?? "").trim() : "";
    const asunto = idxAsu >= 0 ? String(row[idxAsu] ?? "").trim() : "";
    const kind = classifyEventFromSubject(asunto);

    rows.push({ comercial, fecha: undefined, ym, cliente, asunto, kind });
  }

  if (!rows.length) throw new Error(`VISITAS: sin filas válidas en ${sheetName}`);
  // Si no hubo fechas, tendremos ym="" → mostramos un único período vacío
  const periods = Array.from(new Set(rows.map(r => r.ym || ""))).sort((a,b)=>a.localeCompare(b));
  return { rows, sheetName, periods };
}

function buildVisitsModelFromWorkbook(wb: XLSX.WorkBook) {
  const errs:string[] = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      return parseVisitsFromSheet(ws, sn);
    } catch(e:any) { errs.push(`${sn}: ${e?.message || e}`); }
  }
  throw new Error("VISITAS: no pude interpretar ninguna hoja. "+errs.join(" | "));
}

// ================== ACTIVIDADES (archivo independiente) ==================

// (opcional) aún lo dejamos por si no hay fechas y trae "Estado"
function classifyActivityStatus(raw?: string): "completadas" | "vencidas" | "pendientes" | "otros" {
  const s = norm(raw || "");
  if (/(^|[^a-z])comp(?:let|leta|leto|leted|done)/i.test(s)) return "completadas";
  if (/(vencid|overdue|atrasad)/i.test(s)) return "vencidas";
  if (/(pendi|open|abiert)/i.test(s)) return "pendientes";
  return "otros";
}

// ================== ACTIVIDADES (archivo independiente) ==================
// Reemplaza TODA la función parseActivitiesFromSheet por esta

function parseActivitiesFromSheet(ws: XLSX.WorkSheet, sheetName: string) {
  // Leemos como matriz, sin asumir tipos
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];
  if (!A.length) throw new Error("ACTIVIDADES: hoja vacía");

  // Buscamos una fila probable de encabezados, pero IGUAL forzaremos B/D/M
  const normH = (s: any) => norm(String(s));
  let headerRow = 0, best = -1;
  for (let r = 0; r < Math.min(40, A.length); r++) {
    const H = (A[r] || []).map(normH);
    // heurística simple: que aparezca algo tipo “creado”, “estado”, “fecha”
    let sc = 0;
    if (H.some(h => h.includes("creado") || h.includes("owner") || h.includes("comercial"))) sc++;
    if (H.some(h => h.includes("estado") || h.includes("status"))) sc++;
    if (H.some(h => h.includes("fecha"))) sc++;
    if (sc > best) { best = sc; headerRow = r; }
  }

  // Forzamos posiciones físicas (A=0, B=1, C=2, D=3, …, M=12)
  const idxOwner  = 1;  // B: Creado por
  const idxFecha  = 3;  // D: Fecha
  const idxEstado = 12; // M: Estado

  const todayUTC = new Date(Date.UTC(
    new Date().getUTCFullYear(),
    new Date().getUTCMonth(),
    new Date().getUTCDate()
  ));

  const rows: Array<{ comercial: string; status: "completadas"|"vencidas"|"pendientes"|"otros"; raw: any[] }> = [];
  let currentComercial = "";

  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || [];
    // fila vacía → continuar
    if (row.every(v => String(v).trim() === "")) continue;

    // arrastre del comercial (columna B)
    const rawOwner = row[idxOwner];
    if (rawOwner != null && String(rawOwner).trim() !== "") {
      currentComercial = mapComercial(rawOwner);
      // si la fila es solo el título del comercial, saltamos
      const onlyOwner = row.filter((v:any, c:number)=> c!==idxOwner && String(v??"").trim()!=="").length===0;
      if (onlyOwner) continue;
    }
    const comercial = currentComercial;
    if (!comercial) continue;
    if (comercial === "(Sin comercial)") continue; // ← NO contamos filas sin comercial

    // estado (columna M)
    const estadoRaw = String(row[idxEstado] ?? "");
    const est = norm(estadoRaw);

    // fecha (columna D) → puede venir serial, texto o Date
    const fecha = parseDateCell(row[idxFecha]); // ya devuelve UTC sin hora o null

    let status: "completadas" | "vencidas" | "pendientes" | "otros" = "pendientes";

    // Regla pedida:
    // - estado "completed"  => completadas
    // - estado "open"       => mirar D: si hoy > D ⇒ vencidas; si hoy <= D ⇒ pendientes (0 días también pendientes)
    // - cualquier otro      => intentamos la misma lógica de fecha; si no hay fecha, queda "pendientes"
    if (/completed/i.test(estadoRaw)) {
      status = "completadas";
    } else if (/open/i.test(estadoRaw)) {
      if (fecha && todayUTC.getTime() > fecha.getTime()) status = "vencidas";
      else status = "pendientes";
    } else {
      // fallback si llega otro valor en M
      if (fecha && todayUTC.getTime() > fecha.getTime()) status = "vencidas";
      else status = "pendientes";
    }

    rows.push({ comercial, status, raw: row });
  }

  if (!rows.length) throw new Error(`ACTIVIDADES: sin filas válidas en ${sheetName}`);
  return { rows, sheetName };
}

function buildActivitiesModelFromWorkbook(wb: XLSX.WorkBook) {
  const errs: string[] = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      // usa tu parser que ya clasifica con fechas (fin/inicio) y, si falta, por Estado
      return parseActivitiesFromSheet(ws, sn);
    } catch (e: any) {
      errs.push(`${sn}: ${e?.message || e}`);
    }
  }
  throw new Error("ACTIVIDADES: no pude interpretar ninguna hoja. " + errs.join(" | "));
}

// ================== DETALLADO → OFERTAS ==================
type OfferRow = {
  comercial: string;
  fechaOferta: Date;
  etapa?: string;
  valor?: number;
  cuenta?: string;
  raw?: any;
};

/**
 * Busca el nombre estandarizado de una columna por varias variantes posibles.
 */
function pickCol(obj: any, variants: string[]): any {
  if (!obj) return undefined;
  const keys = Object.keys(obj);
  const found = keys.find((k) =>
    variants.some((v) => k.trim().toLowerCase() === v.trim().toLowerCase())
  );
  return found ? obj[found] : undefined;
}

/**
 * Convierte a número tolerante a separadores de miles/comas.
 */
function toNum(v: any): number | undefined {
  if (v == null || v === "") return undefined;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^\d,.-]/g, "");
  if (s.includes(",") && !s.includes(".")) {
    // formato 1.234.567,89 -> pasa coma a punto si no hay punto decimal
    return Number(s.replace(/,/g, ".").replace(/\.(?=.*\.)/g, ""));
  }
  // quita puntos de miles (dejar solo el decimal final)
  const parts = s.split(".");
  if (parts.length > 2) return Number(parts.join(""));
  const n = Number(s);
  return isFinite(n) ? n : undefined;
}

/**
 * Normaliza fecha desde Excel (número serial), string o Date.
 */
function toDate(v: any): Date | undefined {
  if (!v && v !== 0) return undefined;
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    // Excel serial date
    const d = XLSX.SSF ? XLSX.SSF.parse_date_code(v) : null;
    if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d));
    // fallback: días desde 1899-12-30
    const epoch = new Date(Date.UTC(1899, 11, 30)).getTime();
    return new Date(epoch + v * 24 * 60 * 60 * 1000);
  }
  // intenta parsear string DD/MM/YYYY o YYYY-MM-DD
  const s = String(v).trim();
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m1) {
    const d = Number(m1[1]);
    const mo = Number(m1[2]) - 1;
    const y = Number(m1[3].length === 2 ? "20" + m1[3] : m1[3]);
    return new Date(Date.UTC(y, mo, d));
  }
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? undefined : dt;
}

/**
 * Determina si una fila es noise (subtotales, recuentos, etc.)
 */
function isNoiseRow(obj: any): boolean {
  const txt = Object.values(obj)
    .filter((x) => typeof x === "string")
    .join(" ")
    .toLowerCase();
  return (
    txt.includes("subtotal") ||
    txt.includes("recuento") ||
    txt.includes("total") && !txt.includes("precio")
  );
}

// helper: YYYY-MM a partir de un Date
function ym(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

export async function parseOffersFromDetailSheet(
  file: ArrayBuffer,
  periodoYM: string
): Promise<{ rows: OfferRow[]; ranking: { comercial: string; count: number }[] }> {
  const wb = XLSX.read(file, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json: any[] = XLSX.utils.sheet_to_json(ws, { defval: "" });

  const rows: OfferRow[] = [];
  let lastComercial = "";

  for (const r of json) {
    if (isNoiseRow(r)) continue;

    const comercialRaw =
      pickCol(r, ["Comercial", "Vendedor", "Propietario", "Dueño"]) ?? "";
    const fechaRaw =
      pickCol(r, ["Fecha de oferta", "Fecha oferta", "Fecha", "Fecha creación", "Fecha de creación"]) ??
      pickCol(r, ["Fecha Oportunidad", "Oportunidad fecha"]) ?? "";
    const etapaRaw =
      pickCol(r, ["Etapa", "Stage", "Estado"]) ?? undefined;
    const valorRaw =
      pickCol(r, ["Valor de la oferta", "Importe", "Precio total", "Monto"]) ?? undefined;
    const cuentaRaw =
      pickCol(r, ["Cuenta", "Cliente"]) ?? undefined;

    const comercialTxt = String(comercialRaw || "").trim();
    if (comercialTxt) lastComercial = comercialTxt;

    const fecha = toDate(fechaRaw);
    const valor = toNum(valorRaw);
    const comercial = lastComercial.trim();

    if (!fecha || !comercial) continue;

    rows.push({
      comercial,
      fechaOferta: fecha,
      etapa: etapaRaw ? String(etapaRaw).trim() : undefined,
      valor,
      cuenta: cuentaRaw ? String(cuentaRaw).trim() : undefined,
      raw: r,
    });
  }

  const filtered = rows.filter((r) => ym(r.fechaOferta) === periodoYM);

  const counts = new Map<string, number>();
  for (const r of filtered) {
    counts.set(r.comercial, (counts.get(r.comercial) || 0) + 1);
  }
  const ranking = Array.from(counts.entries())
    .map(([comercial, count]) => ({ comercial, count }))
    .sort((a, b) => b.count - a.count);

  return { rows, ranking };
}

async function loadMetasFromSheet(year: number) {
  const url = `${METAS_GET_URL}?year=${encodeURIComponent(year)}`;
  const r = await fetch(url, { method: 'GET' });
  if (!r.ok) throw new Error(`GET metas: ${r.status}`);
  const data = await r.json() as {
    year: number;
    metas: { comercial: string; metaAnual: number; metaOfertas: number; metaVisitas: number }[];
  };
  return data.metas ?? [];
}


// ==================== Parser RESUMEN ====================
function parsePivotSheet(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any;
  if (!A.length) throw new Error("Resumen: hoja vacía");

  const { headerRow, propietarioCol } = findHeaderPosition(A);
  const stageRow = findStageHeaderRow(A, headerRow, propietarioCol + 1);

  const cols: any[] = [];
  let lastStage = "";
  const maxCols = Math.max(A[headerRow]?.length || 0, A[stageRow]?.length || 0);

  for (let c = propietarioCol + 1; c < maxCols; c++) {
    const stage = norm(A[stageRow]?.[c] ?? "") || lastStage; if (stage) lastStage = stage;
    const metric = norm(A[headerRow]?.[c] ?? "");
    if (!stage && !metric) continue;
    cols.push({ col: c, stage, metric });
  }

  const rows: any[] = [];
  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || [];
    const labelRaw = row[propietarioCol];
    const label = norm(labelRaw);
    const isEmpty = row.map((x: any) => norm(x)).join("") === "";
    if (isEmpty) continue;
    if (!label) continue;
    if (label.startsWith("total")) break;

    // arrastre + normalización del comercial
    const comercial = mapComercial(labelRaw);

    const values: Record<string, { sum: number; count: number }> = {};
    for (const cm of cols) {
      const st = cm.stage; if (!st) continue;
      const met = cm.metric; const cell = row[cm.col];
      if (!values[st]) values[st] = { sum: 0, count: 0 };
      if (met.includes("suma") || met.includes("total")) values[st].sum += toNumber(cell);
      else if (met.includes("recuento") || met.includes("count")) values[st].count += toNumber(cell);
    }

    const hasData = Object.values(values).some(v => v.sum || v.count);
    if (hasData) rows.push({ comercial, values });
  }
  return { rows, sheetName };
}

function tryParseAnyPivot(wb: XLSX.WorkBook) {
  const errs: string[] = [];
  for (const name of wb.SheetNames) {
    try { const ws = wb.Sheets[name]; if (!ws) continue; return parsePivotSheet(ws, name); }
    catch (e: any) { errs.push(`${name}: ${e?.message || e}`); }
  }
  throw new Error("Resumen: ninguna hoja válida. " + errs.join(" | "));
}

// Helpers de etapa
const isWonStage  = (k:string) => /closed won|ganad/.test(k);
const isLostStage = (k:string) => /closed lost|perdid/.test(k);

// === KPIs desde RESUMEN ===
function calcPipelineFromPivot(model: any) {
  const porComercial = model.rows.map((r: any) => {
    let sum = 0;
    for (const [stage, agg] of Object.entries(r.values)) {
      const st = norm(stage);
      if (["qualification","needs","needs analysis","proposal","negotiation"].some(s => st.includes(s))) {
        sum += (agg as any).sum || 0;
      }
    }
    return { comercial: r.comercial, pipeline: sum };
  });
  const total = porComercial.reduce((a: number, x: any) => a + x.pipeline, 0);
  return { total, porComercial };
}

// Mapa de ofertas abiertas por comercial: { monto, n }  (excluye Closed Won/Lost)
function openOffersByComercial(detailRows: Array<{ comercial: string; etapa?: string; importe?: number }>) {
  const m = new Map<string, { monto: number; n: number }>();
  for (const r of detailRows) {
    if (!isOpenStage(r.etapa)) continue;
    const imp = Number(r.importe || 0);
    if (!imp) continue;
    const cur = m.get(r.comercial) || { monto: 0, n: 0 };
    cur.monto += imp;
    cur.n += 1;
    m.set(r.comercial, cur);
  }
  return m;
}

function calcWinRateBudgetFromPivot(model: any) {
  const isClosedWon = (s: string) => {
    const z = norm(s);
    return z.includes("closed won") || z.includes("ganad");
  };

  // claves que el pivot suele poner y que NO debemos sumar en el denominador
  const isTotalLike = (key: string) => {
    const z = norm(key);
    return (
      z === "total" ||
      z === "__total__" ||
      z === "totalgeneral" ||
      z.includes("total") ||
      z.includes("registro") || // recuento de registros
      z.includes("count") ||    // count
      z.includes("recuento")    // recuento
    );
  };

  const porComercial = model.rows.map((r: any) => {
    let totalCOP = 0;
    let wonCOP = 0;

    // r.values = { Etapa -> { sum, count, ... }, Total -> { sum, ... } }
    for (const [stage, agg] of Object.entries(r.values)) {
      if (isTotalLike(stage)) continue; // <- EVITA DUPLICAR

      // Toma solo el agregado de SUMA en dinero
      const sum = Number(
        (agg as any)?.sum ??
        (agg as any)?.sumPrice ??
        (agg as any)?.precio ??
        0
      );

      totalCOP += sum;
      if (isClosedWon(stage)) wonCOP += sum;
    }

    const winRate = totalCOP > 0 ? (wonCOP / totalCOP) * 100 : 0;
    return { comercial: r.comercial, wonCOP, totalCOP, winRate };
  });

  const totalAgg = porComercial.reduce(
    (a, x) => ({ wonCOP: a.wonCOP + x.wonCOP, totalCOP: a.totalCOP + x.totalCOP }),
    { wonCOP: 0, totalCOP: 0 }
  );
  const totalWinRate = totalAgg.totalCOP > 0 ? (totalAgg.wonCOP / totalAgg.totalCOP) * 100 : 0;

  return {
    total: { winRate: totalWinRate, wonCOP: totalAgg.wonCOP, totalCOP: totalAgg.totalCOP },
    porComercial
  };
}

function calcWinRateFromPivot(model: any) {
  const porComercial = model.rows.map((r: any) => {
    let total = 0;
    let won   = 0;

    for (const [stage, agg] of Object.entries(r.values)) {
      // ⬅️ evitar duplicar contando la columna TOTAL del pivot
      if (isTotalLike(String(stage))) continue;

      const count = Number((agg as any)?.count || 0);
      total += count;

      const st = norm(String(stage));
      if (st.includes("closed won") || st.includes("ganad")) {
        won += count;
      }
    }

    const winRate = total > 0 ? (won / total) * 100 : 0;
    return { comercial: r.comercial, won, total, winRate };
  });

  const totalAgg = porComercial.reduce(
    (a, x) => ({ won: a.won + x.won, total: a.total + x.total }),
    { won: 0, total: 0 }
  );
  const totalWinRate = totalAgg.total > 0 ? (totalAgg.won / totalAgg.total) * 100 : 0;

  return { total: { winRate: totalWinRate, won: totalAgg.won, total: totalAgg.total }, porComercial };
}

function calcWinRateBudgetFromDetail(detail: any[]) {
  const isClosedWon = (s: string) => {
    const z = norm(s);
    return z.includes("closed won") || z.includes("ganad");
  };

  const byOwner: Record<string, { wonCOP: number; totalCOP: number }> = {};

  detail.forEach((row: any) => {
    const owner = String(row.ownerNorm || row.owner || row.comercial || "SIN_COMERCIAL");
    const amount = Number(row.totalPrice || row.importe || row.monto || row.precio || 0);
    const stage = String(row.stageNorm || row.stage || row.etapa || "");

    if (!byOwner[owner]) byOwner[owner] = { wonCOP: 0, totalCOP: 0 };
    byOwner[owner].totalCOP += amount;
    if (isClosedWon(stage)) byOwner[owner].wonCOP += amount;
  });

  const porComercial = Object.entries(byOwner).map(([comercial, v]) => ({
    comercial,
    wonCOP: v.wonCOP,
    totalCOP: v.totalCOP,
    winRate: v.totalCOP > 0 ? (v.wonCOP / v.totalCOP) * 100 : 0
  }));

  const totalAgg = porComercial.reduce(
    (a, x) => ({ wonCOP: a.wonCOP + x.wonCOP, totalCOP: a.totalCOP + x.totalCOP }),
    { wonCOP: 0, totalCOP: 0 }
  );

  const totalWinRate = totalAgg.totalCOP > 0 ? (totalAgg.wonCOP / totalAgg.totalCOP) * 100 : 0;

  return {
    total: { winRate: totalWinRate, wonCOP: totalAgg.wonCOP, totalCOP: totalAgg.totalCOP },
    porComercial
  };
}

function calcAttainmentFromPivot(model: any, goalFor: (comercial: string) => number) {
  const porComercial = model.rows.map((r: any) => {
    let wonCOP = 0;
    for (const [stage, agg] of Object.entries(r.values)) {
      const st = norm(stage);
      if (st.includes("closed won") || st.includes("ganad")) wonCOP += (agg as any).sum || 0;
    }
  const goal = goalFor(r.comercial);
    const pct = goal > 0 ? (wonCOP * 100) / goal : 0;
    return { comercial: r.comercial, wonCOP, goal, pct };
  });
  const agg = porComercial.reduce((acc: any, x: any) => ({ wonCOP: acc.wonCOP + x.wonCOP, goal: acc.goal + x.goal }), { wonCOP: 0, goal: 0 });
  const totalPct = agg.goal > 0 ? (agg.wonCOP * 100) / agg.goal : 0;
  return { total: { pct: totalPct, wonCOP: agg.wonCOP, goal: agg.goal }, porComercial };
}

// Necesidad de cotización para cumplir meta suponiendo win rate fijo (p.ej. 20%)
function calcForecastNeededFromPivot(
  model: any,
  goalFor: (comercial: string) => number,
  winRate: number // ej: 0.20
) {
  type RowF = { comercial: string; wonCOP: number; goal: number; remaining: number; needQuote: number };

  const porComercial: RowF[] = model.rows.map((r: any) => {
    // Sumar solo etapas ganadas (Closed Won / Ganada)
    let wonCOP = 0;
    for (const [stage, agg] of Object.entries(r.values)) {
      const s = String(stage).toLowerCase();
      if (s.includes("closed won") || s.includes("ganad")) {
        wonCOP += Number((agg as any)?.sum || 0);
      }
    }
    const goal = Math.max(0, goalFor(r.comercial) || 0);
    const remaining = Math.max(0, goal - wonCOP);
    const needQuote = winRate > 0 ? Math.ceil(remaining / winRate) : 0;
    return { comercial: r.comercial, wonCOP, goal, remaining, needQuote };
  });

  // Totales compañía
  const agg = porComercial.reduce(
    (a, x) => ({ wonCOP: a.wonCOP + x.wonCOP, goal: a.goal + x.goal, remaining: a.remaining + x.remaining, needQuote: a.needQuote + x.needQuote }),
    { wonCOP: 0, goal: 0, remaining: 0, needQuote: 0 }
  );

  // Ordena por necesidad de cotización desc
  porComercial.sort((a, b) => b.needQuote - a.needQuote);

  return { porComercial, total: agg };
}

function findHeaderPosition(A:any[][]) {
  let headerRow = 0, propietarioCol = 0, best = -1;
  for (let r=0; r<Math.min(40, A.length); r++) {
    const row = A[r] || [];
    const score = row.reduce((acc:number, v:any, c:number) => {
      const h = norm(v);
      if (h.includes("propietario") || h.includes("owner") || h.includes("comercial") || h.includes("vendedor")) {
        propietarioCol = c; acc += 2;
      }
      if (h.includes("etapa") || h.includes("stage")) acc += 1;
      return acc;
    }, 0);
    if (score > best) { best = score; headerRow = r; }
  }
  return { headerRow, propietarioCol };
}

function findStageHeaderRow(A: any[][], headerRow: number, startCol: number) {
  const looksStageRow = (ri: number) => {
    const row = A[ri] || []; let hits = 0;
    for (let c = startCol; c < row.length; c++) {
      const v = norm(row[c]); if (!v) continue;
      if (["qualification","needs","needs analysis","proposal","negotiation","closed","ganad","perdid"].some(k => v.includes(k))) hits++;
    }
    return hits >= 2;
  };
  for (let r = headerRow - 1; r >= Math.max(0, headerRow - 3); r--) if (looksStageRow(r)) return r;
  return headerRow;
}

function parseVisitsSheetRobust(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any;
  if (!A.length) throw new Error("Visitas: hoja vacía");

  const normH = (s: any) => norm(String(s).replace(/[()↑%]/g, " "));
  const scoreRow = (row: any[]) => {
    const set = row.map(normH);
    let sc = 0;
    if (set.some(h => h.includes("propietario") || h.includes("comercial") || h.includes("vendedor") || h.includes("owner"))) sc++;
    if (set.some(h => h.includes("fecha"))) sc++;
    if (set.some(h => h.includes("visita") || h.includes("cantidad") || h.includes("count"))) sc++;
    return sc;
  };

  let headerRow = 0, best = -1;
  for (let r = 0; r < Math.min(40, A.length); r++) {
    const sc = scoreRow(A[r] || []);
    if (sc > best) { best = sc; headerRow = r; }
  }
  const headers = A[headerRow] || [];
  const findIdx = (...cands: string[]) => {
    for (let c = 0; c < headers.length; c++) {
      const h = normH(headers[c]);
      if (cands.some(cd => h.includes(cd))) return c;
    }
    return -1;
  };

  const idxCom = findIdx("propietario", "comercial", "vendedor", "owner");
  const idxFecha = findIdx("fecha", "creacion", "created");
  const idxCnt = findIdx("visita", "cantidad", "count");

  if (idxCom < 0) throw new Error("Visitas: no encontré columna de Comercial/Propietario.");

  const rows: any[] = [];
  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || [];
    if (row.every((v: any) => String(v).trim() === "")) continue;

    const comercial = mapComercial(row[idxCom]);
    const fecha = idxFecha >= 0 ? parseDateCell(row[idxFecha]) : null;
    let n = 1;
    if (idxCnt >= 0) {
      const vv = toNumber(row[idxCnt]);
      n = vv > 0 ? vv : 0;
    }
    rows.push({ comercial, fecha, n });
  }
  return { rows, sheetName };
}

function tryParseAnyVisits(wb: XLSX.WorkBook) {
  const errs: string[] = [];
  for (const name of wb.SheetNames) {
    try {
      const ws = wb.Sheets[name];
      if (!ws) continue;
      return parseVisitsSheetRobust(ws, name);
    } catch (e: any) { errs.push(`${name}: ${e?.message || e}`); }
  }
  throw new Error("Visitas: ninguna hoja válida. Detalles: " + errs.join(" | "));
}

// ==================== Parser DETALLE ====================
function parseDetailSheetRobust(ws: XLSX.WorkSheet, sheetName: string) {
  const A:any[][] = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" }) as any[][];
  if (!A.length) throw new Error("Detalle: hoja vacía");

  // --- localizar fila de encabezado y columnas clave ---
  const pickCols = (mat:any[][]) => {
    let headerRow = 0, best = -1;
    let colOwner = -1, colStage = -1, colCreated = -1, colClosed = -1, colAmount = -1;

    for (let r = 0; r < Math.min(50, mat.length); r++){
      const row = mat[r] || [];
      let sc = 0, co=-1, cs=-1, cc=-1, cl=-1, ca=-1;
      row.forEach((v:any, c:number) => {
        const h = norm(v);
        if (OWNER_KEYS.some(k=> h.includes(k)))  { sc+=3; co=c; }
        if (STAGE_KEYS.some(k=> h.includes(k)))  { sc+=1; cs=c; }
        if (CREATE_KEYS.some(k=> h.includes(k))) { sc+=2; cc=c; }
        if (CLOSE_KEYS.some(k=> h.includes(k)))  { sc+=2; cl=c; }
        // monto / valor
        if (
          h.includes("importe") || h.includes("monto") || h.includes("valor") ||
          h.includes("precio total") || h.includes("total") && !h.includes("recuento") && !h.includes("count")
        ) { sc+=1; ca=c; }
      });
      if (sc > best && (co>=0 || cs>=0)) {
        best=sc; headerRow=r; colOwner=co; colStage=cs; colCreated=cc; colClosed=cl; colAmount=ca;
      }
    }
    return { headerRow, colOwner, colStage, colCreated, colClosed, colAmount };
  };

  const { headerRow, colOwner, colStage, colCreated, colClosed, colAmount } = pickCols(A);
  if (colOwner < 0) throw new Error("No se encontró columna de Propietario/Comercial");

  const outClosed:any[] = [];   // filas cerradas (para Sales Cycle "cerradas")
  const allRows:any[]   = [];   // todas las oportunidades (abiertas+cerradas)

  let carryCom = "";
  for (let r = headerRow+1; r < A.length; r++){
    const row = A[r] || [];
    const line = norm(row.join(" "));
    if (!line) continue;
    if (line.startsWith("subtotal") || line.startsWith("total") || line.includes("recuento")) continue;

    // arrastre de comercial
    const rawCom = row[colOwner];
    if (rawCom != null && String(rawCom).trim() !== ""){
      carryCom = String(rawCom).trim();
      // si es fila "título" (solo trae el comercial), sigue
      const soloTitulo = row.filter((v:any, c:number)=> c!==colOwner).every(v => String(v??"").trim()==="");
      if (soloTitulo) continue;
    }
    const comercial = carryCom.trim();
    if (!comercial) continue;

    const stage   = colStage>=0   ? String(row[colStage] ?? "") : "";
    const created = colCreated>=0 ? parseExcelDate(row[colCreated]) : null;
    const closed  = colClosed>=0  ? parseExcelDate(row[colClosed])  : null;
    const amount  = colAmount>=0  ? toNumber(row[colAmount]) : 0; // ← NUEVO

    allRows.push({ comercial, stage, created, closed, amount });  // ← guarda amount

    const isClosed = (!!closed) || CLOSED_RX.test(norm(stage));
    if (isClosed && created && closed) {
      outClosed.push({ comercial, created, closed, stage });
    }
  }

  return { rows: outClosed, allRows, sheetName } as const;
}

function tryParseAnyDetail(wb: XLSX.WorkBook){
  const errs:string[] = [];
  for (const sn of wb.SheetNames){
    try{
      const ws = wb.Sheets[sn]; if (!ws) continue;
      // Si parsea, salimos con esa hoja
      return parseDetailSheetRobust(ws, sn);
    }catch(e:any){ errs.push(`${sn}: ${e?.message || e}`); }
  }
  throw new Error("DETALLADO: ninguna hoja válida. " + errs.join(" | "));
}

// Construye el modelo de Ofertas (filas con comercial + "YYYY-MM") usando el DETALLE ya parseado
function buildOffersModelFromDetailModel(detailModel: any) {
  const rows = (detailModel?.allRows || [])
    .filter((r:any) => r?.created instanceof Date) // necesitamos fecha de creación
    .map((r:any) => {
      const d = r.created as Date;
      const ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}`;
      return { comercial: r.comercial, ym };
    });

  const periods = Array.from(new Set(rows.map((r:any)=>r.ym))).sort();
  return { rows, periods };
}

// ====================== KPI Calcs =======================

function calcSalesCycleFromDetail(detailModel: any, mode: "all" | "won") {
  const MS = 24 * 60 * 60 * 1000;

  const rows = (detailModel?.rows || []).filter((r: any) => {
    if (!r.created || !r.closed) return false;
    if (mode === "won") return CLOSED_WON_RX.test(String(r.stage || ""));
    return true; // todas cerradas (ganadas + perdidas)
  });

  // agrupamos por comercial acumulando milisegundos y cantidad
  const by = new Map<string, { sumMs: number; n: number }>();
  for (const r of rows) {
    const deltaMs = r.closed.getTime() - r.created.getTime();
    if (!isFinite(deltaMs) || deltaMs < 0 || deltaMs > 3650 * MS) continue; // saneamiento
    const acc = by.get(r.comercial) || { sumMs: 0, n: 0 };
    acc.sumMs += deltaMs; acc.n += 1;
    by.set(r.comercial, acc);
  }

  const porComercial = Array.from(by.entries())
    .map(([comercial, v]) => ({
      comercial,
      avgDays: v.n ? Math.round((v.sumMs / v.n) / MS) : 0, // ← redondeo sólo al final
      n: v.n
    }))
    .sort((a, b) => a.avgDays - b.avgDays);

  const totals = Array.from(by.values())
    .reduce((acc, v) => ({ sumMs: acc.sumMs + v.sumMs, n: acc.n + v.n }), { sumMs: 0, n: 0 });

  const totalAvgDays = totals.n ? Math.round((totals.sumMs / totals.n) / MS) : 0;

  return { totalAvgDays, totalCount: totals.n, porComercial };
}

// Promedio de días SOLO con oportunidades cerradas (won+lost) o solo won
function calcSalesCycleClosed(detailModel: any, onlyWon = false) {
  const MS = 24 * 60 * 60 * 1000;
  const rows = (detailModel?.allRows || []).filter((r: any) => {
    // Debe tener ambas fechas para cerradas
    if (!r?.created || !r?.closed) return false;

    const stage = String(r.stage || "");
    if (onlyWon) {
      return CLOSED_WON_RX.test(stage);
    }
    // won + lost
    return CLOSED_WON_RX.test(stage) || CLOSED_LOST_RX.test(stage);
  });

  const by = new Map<string, { sumMs: number; n: number }>();
  for (const r of rows) {
let deltaMs = r.closed.getTime() - r.created.getTime();
if (!isFinite(deltaMs) || deltaMs > 3650 * MS) continue;
if (deltaMs < 0) deltaMs = 0;

    const acc = by.get(r.comercial) || { sumMs: 0, n: 0 };
    acc.sumMs += deltaMs; acc.n += 1;
    by.set(r.comercial, acc);
  }

  const porComercial = Array.from(by.entries())
    .map(([comercial, v]) => ({
      comercial,
      avgDays: v.n ? Math.round((v.sumMs / v.n) / MS) : 0,
      n: v.n
    }))
    .sort((a, b) => a.avgDays - b.avgDays);

  const totals = Array.from(by.values())
    .reduce((acc, v) => ({ sumMs: acc.sumMs + v.sumMs, n: acc.n + v.n }), { sumMs: 0, n: 0 });
  const totalAvgDays = totals.n ? Math.round((totals.sumMs / totals.n) / MS) : 0;

  return { totalAvgDays, totalCount: totals.n, porComercial };
}

function calcSalesCycleAllOffers(detailModel: any) {
  // detailModel.allRows = [{ comercial, created: Date|null, closed: Date|null, stage: string }]
  const MS = 24 * 60 * 60 * 1000;
  const todayUTC = new Date(Date.UTC(
    new Date().getUTCFullYear(),
    new Date().getUTCMonth(),
    new Date().getUTCDate()
  ));

  const all = detailModel?.allRows || [];
  const by = new Map<string, { sumMs: number; n: number }>();

for (const r of all) {
  const created = r.created;
  if (!created) continue;

  const end = r.closed ?? todayUTC; // cerrada: cierre; abierta: hoy

  let deltaMs = end.getTime() - created.getTime();
  if (!isFinite(deltaMs) || deltaMs > 3650 * MS) continue;
  if (deltaMs < 0) deltaMs = 0;

  const acc = by.get(r.comercial) || { sumMs: 0, n: 0 };
  acc.sumMs += deltaMs; acc.n += 1;
  by.set(r.comercial, acc);
}

  const porComercial = Array.from(by.entries())
    .map(([comercial, v]) => ({
      comercial,
      avgDays: v.n ? Math.round((v.sumMs / v.n) / MS) : 0,
      n: v.n
    }))
    .sort((a, b) => a.avgDays - b.avgDays);

  const totals = Array.from(by.values())
    .reduce((acc, v) => ({ sumMs: acc.sumMs + v.sumMs, n: acc.n + v.n }), { sumMs: 0, n: 0 });
  const totalAvgDays = totals.n ? Math.round((totals.sumMs / totals.n) / MS) : 0;

  return { totalAvgDays, totalCount: totals.n, porComercial };
}

// ================== UI (Router + Screens) ==================
const RouteHome = ({ onEnter }: { onEnter: () => void }) => (
  <div className="min-h-screen bg-gray-50 flex items-center justify-center">
    <div className="bg-white border rounded-2xl p-8 max-w-md w-full text-center">
      <h1 className="text-2xl font-bold">INGETES • Portal KPI</h1>
      <p className="text-sm text-gray-600 mt-2">Bienvenido. Ingresa para gestionar y analizar los KPI comerciales.</p>
      <button className="mt-6 px-4 py-2 rounded bg-black text-white" onClick={onEnter}>Entrar al Portal KPI</button>
    </div>
  </div>
);

export default function IngetesKPIApp() {
 const [route, setRoute] = useState<
   "HOME" | "MENU" | "KPI_PIPELINE" | "KPI_WINRATE" | "KPI_CYCLE" | "KPI_ATTAIN" |
   "KPI_OFFERS" | "KPI_VISITS" | "KPI_ACTIVITIES" 
 >("MENU");

  const [fileDetailName, setFileDetailName] = useState("");
  const [fileVisitsName, setFileVisitsName] = useState("");
  const [filePivotName, setFilePivotName] = useState("");
  const [fileActivitiesName, setFileActivitiesName] = useState("");
  const [activitiesModel, setActivitiesModel] = useState<any>(null);
  const [detail, setDetail] = useState<any>(null);
  const [offersModel, setOffersModel] = useState<any>(null);
  const [visitsModel, setVisitsModel] = useState<any>(null);
  const [pivot, setPivot] = useState<any>(null);
  const [cycleMode, setCycleMode] = useState<"all" | "won" | "offers">("all");
  const [offersPeriod, setOffersPeriod] = useState<string>("");
  const [visitsPeriod, setVisitsPeriod] = useState<string>("");
  const [selectedComercial, setSelectedComercial] = useState("ALL");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [winRateTarget, setWinRateTarget] = useState(30);
  const [cycleTarget, setCycleTarget] = useState(45);
  const [assumedWinRate, setAssumedWinRate] = useState<number>(20); // 👈 aquí sí

// ===== Ajustes (metas por comercial) =====
const [showSettings, setShowSettings] = useState(false);
const [settingsYear, setSettingsYear] = useState<number>(new Date().getFullYear());
const [settingsRows, setSettingsRows] = useState<MetaRecord[]>([]);
const [loadingSettings, setLoadingSettings] = useState(false);
const [savingSettings, setSavingSettings] = useState(false);
const [metasYear, setMetasYear] = useState<number>(new Date().getFullYear());
const [metasModalRows, setMetasModalRows] = useState<
  { comercial: string; metaAnual: number; metaOfertas: number; metaVisitas: number; }[]
>([]);
const [savingMetas, setSavingMetas] = useState(false);

// ===== Metas para KPIs (por año desde Sheet) =====
type MetaSheetRow = {
  comercial: string;
  metaAnual: number;
  metaOfertas: number;
  metaVisitas: number;
};
const [metasByYear, setMetasByYear] = useState<Record<number, MetaSheetRow[]>>({});

const normalizeName = (s: string) =>
  String(s || "").normalize("NFKC").trim().replace(/\s+/g, " ").toUpperCase();

async function ensureMetasForYear(year: number) {
  if (!year || metasByYear[year]) return;
  const url = `${METAS_GET_URL}${METAS_GET_URL.includes("?") ? "&" : "?"}year=${year}`;
  const res = await fetch(url, { cache: "no-store" });
  const data = await res.json();
  const metas: MetaRecordForYear[] = (data?.metas || []).map((m: any) => ({
    comercial: normalizeName(m.comercial),
    metaAnual: Number(m.metaAnual || 0),
    metaOfertas: Number(m.metaOfertas || 0),
    metaVisitas: Number(m.metaVisitas || 0),
  }));
  setMetasByYear(prev => ({ ...prev, [year]: metas }));
}

function metaOfertasFor(comercial: string, year: number) {
  const arr = metasByYear[year] || [];
  const rec = arr.find(m => m.comercial === normalizeName(comercial));
  return rec?.metaOfertas ?? 0;
}
function metaVisitasFor(comercial: string, year: number) {
  const arr = metasByYear[year] || [];
  const rec = arr.find(m => m.comercial === normalizeName(comercial));
  return rec?.metaVisitas ?? 0;
}
function metaAnualFor(comercial: string, year: number) {
  const arr = metasByYear[year] || [];
  const rec = arr.find(m => m.comercial === normalizeName(comercial));
  return rec?.metaAnual ?? 0;
}

// Carga metas según los períodos seleccionados
React.useEffect(() => {
  const y = Number((offersPeriod || "").slice(0, 4)) || new Date().getFullYear();
  ensureMetasForYear(y);
}, [offersPeriod]);

React.useEffect(() => {
  const y = Number((visitsPeriod || "").slice(0, 4)) || new Date().getFullYear();
  ensureMetasForYear(y);
}, [visitsPeriod]);

React.useEffect(() => {
  ensureMetasForYear(settingsYear);
}, [settingsYear]);

// Utilidad: sacar lista de comerciales detectados (de los archivos cargados)
function getAllComerciales(): string[] {
  const set = new Set<string>();

  // RESUMEN (pivot)
  if (pivot?.rows) {
    pivot.rows.forEach((r: any) => {
      const c = String(r.comercial || r.Comercial || "").trim();
      if (c) set.add(c);
    });
  }

  // DETALLADO – ofertas / ciclo
  if (offersModel?.rows) {
    offersModel.rows.forEach((r: any) => {
      const c = String(r.comercial || r.Comercial || "").trim();
      if (c) set.add(c);
    });
  }
  if (detail?.allRows) {
    detail.allRows.forEach((r: any) => {
      const c = String(r.comercial || r.Comercial || "").trim();
      if (c) set.add(c);
    });
  }

  // VISITAS
  if (visitsModel?.rows) {
    visitsModel.rows.forEach((r: any) => {
      const c = String(r.comercial || r.Comercial || "").trim();
      if (c) set.add(c);
    });
  }

  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

// Abre el panel: lee metas del año y llena la tabla
async function openSettings() {
  try {
    setShowSettings(true);
    setLoadingSettings(true);

    // GET directo al Apps Script (echo), forzando el año
    const url = `${METAS_GET_URL}${METAS_GET_URL.includes("?") ? "&" : "?"}year=${settingsYear}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`GET metas ${res.status}`);
    const data = await res.json();

    const metas = Array.isArray(data?.metas) ? data.metas : [];

    // Normaliza y fuerza números (0 si viene null o string)
    const rows = metas.map((m: any) => ({
      comercial: String(m.comercial || "").trim().toUpperCase(),
      metaAnual: Number(m.metaAnual ?? 0),
      metaOfertas: Number(m.metaOfertas ?? 0),
      metaVisitas: Number(m.metaVisitas ?? 0),
    }))
    // orden alfabético
    .sort((a: any, b: any) => a.comercial.localeCompare(b.comercial));

    setSettingsRows(rows);
  } catch (e) {
    console.error(e);
    setSettingsRows([]); // deja el mensaje "No hay comerciales…" si falla
  } finally {
    setLoadingSettings(false);
  }
}

// Guarda en el Sheet vía /exec
async function saveSettings() {
  try {
    setSavingSettings(true);

    const payload = {
      apiKey: METAS_API_KEY,       // debe coincidir con getApiKey() del Apps Script
      year: settingsYear,
      metas: settingsRows.map(r => ({
        comercial: r.comercial,
        metaAnual: Number(r.metaAnual || 0),
        metaOfertas: Number(r.metaOfertas || 0),
        metaVisitas: Number(r.metaVisitas || 0),
      })),
    };

    const res = await fetch(METAS_POST_URL, {
      method: "POST",
      // Enviamos como texto para evitar preflight CORS
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });

    // Lee la respuesta (Apps Script devuelve JSON)
    const outText = await res.text();
    let out: any = {};
    try { out = JSON.parse(outText); } catch {}

    if (!res.ok || !out?.ok) {
      throw new Error(out?.error || `POST metas ${res.status}`);
    }

    alert("Metas guardadas ✅");
    setShowSettings(false);
  } catch (e: any) {
    alert("Error guardando metas: " + (e?.message || e));
  } finally {
    setSavingSettings(false);
  }
}

const resetAll = () => {
  setFilePivotName("");  setPivot(null);
  setFileDetailName(""); setDetail(null); setOffersModel(null); setOffersPeriod("");
  setFileVisitsName(""); setVisitsModel(null); setVisitsPeriod("");
  setSelectedComercial("ALL");
  setWinRateTarget(30); setCycleTarget(45);
  setError(""); setInfo("");
};

  const colorForWinRate = (valuePct: number) => valuePct >= winRateTarget ? "bg-green-500" : (valuePct >= winRateTarget * 0.8 ? "bg-yellow-400" : "bg-red-500");
  const colorForCycle = (days: number) => days <= cycleTarget ? "bg-green-500" : (days <= cycleTarget * 1.2 ? "bg-yellow-400" : "bg-red-500");

async function onDetailFile(f: File) {
  setError("");
  setFileDetailName(f.name);
  try {
    const wb = await readWorkbookRobust(f);      // <-- lector robusto
    const dm = tryParseAnyDetail(wb);            // <-- parser robusto
    setDetail(dm);

    // (opcional) construir modelo de ofertas por período si ya lo usas
    const rows = (dm.allRows || []).filter((r:any)=> r?.created instanceof Date)
      .map((r:any)=> {
        const d = r.created as Date;
        const ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}`;
        return { comercial: r.comercial, ym };
      });
    const periods = Array.from(new Set(rows.map((r:any)=>r.ym))).sort();
    setOffersModel({ rows, periods });
    if (periods.length) setOffersPeriod(periods[periods.length-1]);

  } catch (e:any) {
    setDetail(null);
    setOffersModel(null);
    setError(`Detalle: ${e?.message || e}`);
  }
}

async function onVisitsFile(f: File) {
  setError("");
  setInfo(prev => prev ? prev + "\n" : "");
  setFileVisitsName(f.name);

  try {
    const wb = await readWorkbookRobust(f);
    const vm = buildVisitsModelFromWorkbook(wb);
    setVisitsModel(vm);
    if (vm.periods && vm.periods.length) {
      setVisitsPeriod(vm.periods[vm.periods.length - 1]); // último mes
    }
    setInfo(prev => (prev + `Visitas OK • hoja: ${vm.sheetName}`).trim());
  } catch (e:any) {
    setVisitsModel(null);
    setError(prev => (prev ? prev + "\n" : "") + `Visitas: ${e?.message || e}`);
  }
}
async function onActivitiesFile(f: File) {
  setError(""); setFileActivitiesName(f.name);
  try {
    const wb = await readWorkbookRobust(f);
    const am = buildActivitiesModelFromWorkbook(wb); // ← ahora existe
    setActivitiesModel(am);
  } catch (e:any) {
    setActivitiesModel(null);
    setError(prev => (prev ? prev + "\n" : "") + `Actividades: ${e?.message || e}`);
  }
}
async function onPivotFile(f: File) {
  setError(""); setInfo(prev => prev ? prev + "\n" : "");
  setFilePivotName(f.name);
  try {
    const wb = await readWorkbookRobust(f);
    const pv = tryParseAnyPivot(wb);   // definido abajo
    setPivot(pv);
    setInfo(prev => (prev + `Resumen OK • hoja: ${pv.sheetName}`).trim());
  } catch (e:any) {
    setPivot(null);
    setError(prev => (prev ? prev + "\n" : "") + `Resumen: ${e?.message || e}`);
  }
}

const cycleData = useMemo(() => {
  if (!detail) return { kind: cycleMode, data: null };
  try {
    if (cycleMode === "offers") {
      // TODAS las ofertas (abiertas + cerradas): cierre–creación para cerradas, hoy–creación para abiertas
      return { kind: "offers", data: calcSalesCycleAllOffers(detail) };
    }
    if (cycleMode === "won") {
      // Solo ganadas
      return { kind: "won", data: calcSalesCycleClosed(detail, true) };
    }
    // "Todas cerradas" = won + lost (con fecha de cierre)
    return { kind: "all", data: calcSalesCycleClosed(detail, false) };
  } catch {
    return { kind: cycleMode, data: null };
  }
}, [detail, cycleMode]);

  const comercialesMenu = useMemo(() => FIXED_COMERCIALES, []);
  const pipeline = useMemo(() => pivot ? calcPipelineFromPivot(pivot) : { total: 0, porComercial: [] }, [pivot]);
  const winRate  = useMemo(() => pivot ? calcWinRateFromPivot(pivot)   : { total: { winRate: 0, won: 0, total: 0 }, porComercial: [] }, [pivot]);
  const visitsKPI = useMemo(() => {
  if (!visitsModel) return { total: 0, porComercial: [] as any[], periods: [] as string[], period: "" };
  const periods = visitsModel.periods || [];
  const sel = periods.includes(visitsPeriod) ? visitsPeriod : (periods[periods.length-1] || "");
  const rows = visitsModel.rows.filter((r:any)=> r.ym === sel);

  const by = new Map<string, number>();
  for (const r of rows) {
    by.set(r.comercial, (by.get(r.comercial) || 0) + 1); // una fila = una visita
  }

  const porComercial = Array.from(by.entries())
    .map(([comercial, count]) => ({ comercial, count }))
    .sort((a,b)=> b.count - a.count);

  const total = porComercial.reduce((a,x)=>a+x.count, 0);
  return { total, porComercial, periods, period: sel };
}, [visitsModel, visitsPeriod]);

  const BackBar = ({ title }: { title: string }) => (
    <header className="px-4 py-3 bg-white border-b sticky top-0 z-10">
      <div className="max-w-6xl mx-auto flex items-center justify-between gap-3">
        <h2 className="text-xl md:text-2xl font-bold">{title}</h2>
        <div className="flex gap-2">
          <button className="px-4 py-2 text-sm font-medium rounded-lg border bg-white hover:bg-gray-50" onClick={() => setRoute("MENU")}>Volver al menú</button>
        </div>
      </div>
    </header>
  );

// === MONTO DE OFERTAS ABIERTAS por comercial (excluye Won y Lost) ===
const openAmountsByComercial = React.useMemo(() => {
  const m = new Map<string, number>();
  const rows = detail?.allRows || [];

  for (const r of rows) {
    const st = String(r.stage || "").toUpperCase();
    const isWon  = /CLOSED\s*WON|GANAD/.test(st);
    const isLost = /CLOSED\s*LOST|PERDID/.test(st);
    if (isWon || isLost) continue; // <- SOLO abiertas

    const imp = Number((r as any).amount || 0);
    if (!imp) continue;

    const com = mapComercial(r.comercial);
    if (!com) continue;

    const key = nameKey(com);
    m.set(key, (m.get(key) || 0) + imp);
  }
  return m;
}, [detail]);

// === MONTO DE TODAS LAS OFERTAS (abiertas + perdidas + ganadas) por comercial ===
const openPlusLostAmountsByComercial = React.useMemo(() => {
  const m = new Map<string, number>();
  const rows = detail?.allRows || [];

  for (const r of rows) {
    // tomamos todas las etapas, sin excluir ninguna
    const imp = Number((r as any).amount || 0);
    if (!imp) continue;

    const com = mapComercial(r.comercial);
    if (!com) continue;

    const key = nameKey(com);
    m.set(key, (m.get(key) || 0) + imp);
  }
  return m;
}, [detail]);


const ScreenPipeline = () => {
  if (!pivot) {
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Forecast de meta (cotización necesaria)" />
        <main className="max-w-6xl mx-auto p-4">
          <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg text-yellow-800">
            Carga primero el archivo <b>RESUMEN</b> para ver este KPI.
          </div>
        </main>
      </div>
    );
  }

  // Tomamos metas del año que usas en Ajustes (settingsYear)
const WIN_RATE = Math.max(0, Math.min(100, assumedWinRate)) / 100; // 0..1

  const goalFor = (com: string) => metaAnualFor(com, settingsYear);

  // Asegura que tenemos las metas de ese año
  React.useEffect(() => { ensureMetasForYear(settingsYear); }, [settingsYear]);

  // Calcula: won (cerrado), remaining (faltante) y needQuote (faltante/0.2)
const kpi = React.useMemo(
  () => calcForecastNeededFromPivot(pivot, goalFor, WIN_RATE),
  [pivot, goalFor, WIN_RATE]
);

  // Comodines seleccionados
  const selected = React.useMemo(() => {
    if (selectedComercial === "ALL") return null;
    return kpi.porComercial.find(r => r.comercial === selectedComercial) ||
           { comercial: selectedComercial, wonCOP: 0, goal: goalFor(selectedComercial), remaining: 0, needQuote: 0 };
  }, [kpi, selectedComercial]);

  const fmtCOP = (n: number) => (Number(n)||0).toLocaleString("es-CO", { style:"currency", currency:"COP", maximumFractionDigits:0 });

  // Para barras
  const maxBar = React.useMemo(() => {
    const arr = onlySelected(kpi.porComercial, selectedComercial);
    return Math.max(kpi.total.needQuote || 0, ...(arr.map(r => r.needQuote || 0)), 1);
  }, [kpi, selectedComercial]);

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Forecast de meta (cotización necesaria)" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        {/* Header */}
<section className="p-4 bg-white rounded-xl border">
  <div className="flex flex-col md:flex-row md:items-center md:gap-4">
    <div className="text-base text-gray-700 font-semibold">
      Comercial: <b>{selectedComercial}</b>
    </div>
    <div className="text-base text-gray-700 font-semibold">
      Año (Sheet): <b>{settingsYear}</b>
    </div>
<div className="text-base text-gray-700 font-semibold md:ml-auto flex items-center gap-2">
  Win Rate asumido:
  <input
    type="number"
    min={1}
    max={100}
    step={1}
    className="w-20 border rounded-lg px-2 py-1 text-sm text-right"
    value={assumedWinRate}
    onChange={(e) => setAssumedWinRate(Number(e.target.value || 0))}
    title="Porcentaje de conversión esperado de cotizaciones a ganadas"
  />
  <span>%</span>
</div>
  </div>

  {(() => {
    const isAll = selectedComercial === "ALL";
    const sel = isAll
      ? null
      : (kpi.porComercial.find(r => r.comercial === selectedComercial) ||
         { comercial: selectedComercial, wonCOP: 0, goal: metaAnualFor(selectedComercial, settingsYear), remaining: 0, needQuote: 0 });

    // Montos abiertos
    const openAmtSel = isAll
      ? Array.from(openAmountsByComercial.values()).reduce((a, b) => a + b, 0)
      : (openAmountsByComercial.get(nameKey(selectedComercial)) || 0);

    const needSel = isAll ? kpi.total.needQuote : sel?.needQuote || 0;

    const pctOpenSel = needSel > 0 ? Math.min(100, Math.round((openAmtSel / needSel) * 100)) : 0;
    const colorClassSel = coverageColor(needSel > 0 ? openAmtSel / needSel : 0);

    const remainingVal = isAll ? kpi.total.remaining : sel?.remaining || 0;

    return (
      <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
        {/* Faltante */}
        <div className="bg-white p-4 rounded-lg border text-center">
          <div className="text-sm text-gray-500 mb-1">
            {isAll ? "Faltante (compañía)" : "Faltante (comercial)"}
          </div>
          <div className="text-xl font-bold text-gray-800">
            {fmtCOP(remainingVal)}
          </div>
        </div>

        {/* Cotización necesaria */}
        <div className="bg-white p-4 rounded-lg border text-center">
          <div className="text-sm text-gray-500 mb-1">
            {isAll ? "Cotización necesaria (compañía)" : "Cotización necesaria (comercial)"}
          </div>
          <div className="text-xl font-bold text-gray-800">
            {fmtCOP(needSel)}
          </div>
        </div>

        {/* Avance con semáforo */}
        <div className="bg-white p-4 rounded-lg border text-center">
          <div className="text-sm text-gray-500 mb-1">
            {isAll ? "Avance total (abiertas vs necesita)" : "Avance (abiertas vs necesita)"}
          </div>
          <div className="text-xl font-bold flex items-center justify-center gap-2 text-gray-800">
            <span>{pctOpenSel}%</span>
            <span className={`inline-block w-5 h-5 rounded-full ${colorClassSel}`} />
          </div>
        </div>
      </div>
    );
  })()}

  <div className="text-xs text-gray-500 mt-3">
    Fórmula: <em>necesidad de cotización</em> = <em>(meta anual − cerrado)</em> ÷ <em>0,20</em>.
  </div>
</section>

{/* Ranking visual */}
<section className="p-4 bg-white rounded-xl border">
  <div className="mb-3 font-semibold text-gray-800">
    Ranking por comercial (Forecast necesario para cumplir meta)
  </div>

  <div className="grid grid-cols-1 gap-3">
    {onlySelected(kpi.porComercial, selectedComercial).map((row, i) => {
      const pctBar = Math.min(
        Math.round(((row.needQuote || 0) / (maxBar || 1)) * 100),
        100
      );

      const fmtCOP = (n: number) =>
        (Number(n) || 0).toLocaleString("es-CO", {
          style: "currency",
          currency: "COP",
          maximumFractionDigits: 0,
        });

      const faltantePct = row.goal
        ? Math.min(100, Math.round((row.remaining / row.goal) * 100))
        : 0;

      return (
        <div
          key={row.comercial}
          className="rounded-xl border border-gray-200 shadow-sm bg-gray-50 hover:bg-gray-100 transition-all"
        >
          <div className="p-4 flex flex-col gap-2">
            {/* Encabezado */}
            <div className="flex justify-between items-center">
              <div className="font-semibold text-gray-900 text-base">
                {i + 1}. {row.comercial}
              </div>
            </div>
<div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 mt-2">
  {/* Meta anual */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Meta anual</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(row.goal)}
    </div>
  </div>

  {/* Ofertas ganadas */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Ofertas ganadas</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(row.wonCOP)}
    </div>
  </div>

  {/* Faltante para cumplimiento */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Faltante para cumplimiento</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(row.remaining)}
    </div>
  </div>

  {/* Necesita cotizar */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Necesita cotizar</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(row.needQuote)}
    </div>
  </div>

  {/* Ofertas abiertas */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Ofertas abiertas</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(openAmountsByComercial.get(nameKey(row.comercial)) || 0)}
    </div>
  </div>

  {/* Abiertas + perdidas */}
  <div className="bg-white p-4 rounded-xl border text-center shadow-sm">
    <div className="text-base md:text-lg font-semibold text-gray-600">Total en ofertas</div>
<div className="mt-1 text-3xl md:text-4xl font-normal tabular-nums text-gray-700">
      {fmtCOP(openPlusLostAmountsByComercial.get(nameKey(row.comercial)) || 0)}
    </div>
  </div>
</div>

{/* Barra: abiertas vs lo que necesita cotizar */}
{(() => {
  const openAmt = openAmountsByComercial.get(nameKey(row.comercial)) || 0;
  const pctOpen = row.needQuote > 0
    ? Math.min(100, Math.round((openAmt / row.needQuote) * 100))
    : 0;
  return (
    <div className="mt-3">
<div className="flex justify-between text-xs text-gray-500 mb-1">
  <span className="flex items-center gap-1">
    {pctOpen}% 
    <span className={`${coverageColor(row.needQuote > 0 ? (openAmt / row.needQuote) : 0)} text-lg leading-none`}>●</span>
  </span>
</div>
      <div className="h-2 bg-gray-200 rounded-full overflow-hidden">
        <div
          className="h-2 bg-blue-600 rounded-full transition-all"
          style={{ width: `${pctOpen}%` }}
        />
      </div>
      <div className="mt-1 text-xs text-gray-500">
        Abiertas: <b>{fmtCOP(openAmt)}</b> / Necesita: <b>{fmtCOP(row.needQuote)}</b>
      </div>
    </div>
  );
})()}
          </div>
        </div>
      );
    })}
  </div>
</section>
      </main>
    </div>
  );
};

  const ScreenWinRate = () => {
const [mode, setMode] = React.useState<"cantidad" | "presupuesto">("cantidad");
const data = useMemo(() => {
  if (!pivot) {
    return { total: { winRate: 0, won: 0, total: 0 }, porComercial: [] as any[] };
  }
  return mode === "cantidad"
    ? calcWinRateFromPivot(pivot)
    : calcWinRateBudgetFromPivot(pivot);
}, [pivot, mode]);

// formateo de COP reutilizable en toda la pantalla
const money = (n: number) =>
  (Number(n) || 0).toLocaleString("es-CO", {
    style: "currency",
    currency: "COP",
    maximumFractionDigits: 0,
  });

// Helpers para leer won / total según el modo activo
const getWon = (row: any) =>
  mode === "presupuesto" ? Number(row?.wonCOP || 0) : Number(row?.won || 0);

const getTotal = (row: any) =>
  mode === "presupuesto" ? Number(row?.totalCOP || 0) : Number(row?.total || 0);

const selected = useMemo(() => {
  if (!pivot) return 0; 
  if (selectedComercial === "ALL") return data.total.winRate; 
  return data.porComercial.find(r => r.comercial === selectedComercial)?.winRate || 0;
}, [pivot, data, selectedComercial]);
    
  const max = useMemo(() => {
  const arr = onlySelected(data.porComercial, selectedComercial);
  return Math.max(data.total.winRate || 0, ...(arr.map((row: any) => row.winRate || 0)));
}, [data, selectedComercial]);
    
    const color = (v: number)=> v>=winRateTarget?"bg-green-500":(v>=winRateTarget*0.8?"bg-yellow-400":"bg-red-500");
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Tasa de Cierre (Win Rate)" />
        <main className="max-w-6xl mx-auto p-4 space-y-6">
<section className="p-4 bg-white rounded-xl border">
  <div className="flex flex-col md:flex-row md:items-center md:gap-4">
    <div className="text-sm text-gray-500">Comercial: <b>{selectedComercial}</b></div>

    <div className="text-sm text-gray-500 md:ml-auto">
      Meta Win Rate (%):
      <input
        type="number"
        className="ml-2 border rounded-lg px-2 py-1 text-sm w-20"
        value={winRateTarget}
        onChange={(e) => setWinRateTarget(Number(e.target.value))}
      />
    </div>
  </div>

{(() => {
const totalWon   = getWon(data.total);
const totalCount = getTotal(data.total);
const totalRate  = Math.round(data.total.winRate || 0);

const rowSel = selectedComercial === "ALL"
  ? null
  : data.porComercial.find((r: any) => r.comercial === selectedComercial);

const selWon   = rowSel ? getWon(rowSel)   : 0;
const selCount = rowSel ? getTotal(rowSel) : 0;
const selRate  = Math.round(
  selectedComercial === "ALL"
    ? (data.total.winRate || 0)
    : (rowSel?.winRate || 0)
);

const money = (n: number) =>
  (Number(n) || 0).toLocaleString("es-CO", {
    style: "currency",
    currency: "COP",
    maximumFractionDigits: 0,
  });

// Etiqueta won/total según el modo
const pair = (won: number, total: number) =>
  mode === "presupuesto" ? `${money(won)} / ${money(total)}` : `${won}/${total}`;

  const cumplPct = winRateTarget > 0
    ? Math.round((selRate / winRateTarget) * 100)
    : (selRate > 0 ? 100 : 0);

  const st = color(selRate);

  return (
    <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
<StatCard label="Win Rate (compañía)">
  {totalRate}% ({pair(totalWon, totalCount)})
</StatCard>

<StatCard label="Del comercial seleccionado">
  {selRate}% ({pair(selWon, selCount)})
</StatCard>
      <StatCard label="Cumplimiento vs meta (%)">
        <span className="flex items-center gap-2">
          <span>{cumplPct}%</span>
          <span className={`inline-block w-3 h-3 rounded-full ${st}`} />
        </span>
      </StatCard>
    </div>
  );
})()}

</section>
          {pivot && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="mb-3 font-semibold">Win Rate por comercial</div>
{/* Selector de modo (mismo look que Sales Cycle) */}
<div className="flex items-center justify-center mb-4">
  <div className="inline-flex rounded-lg border overflow-hidden">
    <button
      className={`px-3 py-1 text-sm ${mode === "cantidad" ? "bg-gray-900 text-white" : "bg-white"}`}
      onClick={() => setMode("cantidad")}
      title="Win rate por cantidad (# ganadas / # totales)"
    >
      Por cantidad
    </button>
    <button
      className={`px-3 py-1 text-sm border-l ${mode === "presupuesto" ? "bg-gray-900 text-white" : "bg-white"}`}
      onClick={() => setMode("presupuesto")}
      title="Win rate por presupuesto ($ ganadas / $ totales)"
    >
      Por presupuesto
    </button>
  </div>
</div>
              <div className="space-y-2">
                {onlySelected(data.porComercial, selectedComercial).map((row: any) => {
                  const pct = Math.round((row.winRate / (max || 1)) * 100);
                  return (
                    <div key={row.comercial} className="text-sm">
                      <div className="flex justify-between items-center">
                        <span className="font-medium">{row.comercial}</span>
                        <span className="flex items-center gap-2"><span className={`inline-block w-2 h-2 rounded-full ${color(row.winRate)}`}></span>
{(() => {
  const won = getWon(row);
  const tot = getTotal(row);
  const txt = mode === "presupuesto"
    ? `${money(won)} / ${money(tot)}`
    : `${won}/${tot}`;
  return <span>{Math.round(row.winRate)}% ({txt})</span>;
})()}
                        </span>
                      </div>
                      <div className="h-2 bg-gray-200 rounded"><div className="h-2 rounded bg-gray-700" style={{ width: pct + "%" }} /></div>
                    </div>
                  );
                })}
              </div>
            </section>
          )}
        </main>
      </div>
    );
  };

// === ScreenOffers ===
const ScreenOffers = () => {
  const data = offersKPI;

  // Ofertas del comercial seleccionado en el período activo
const selected = useMemo(() => {
  if (!offersModel) return 0;
  if (selectedComercial === "ALL") return data.total;
  const row = data.porComercial.find((r: any) => r.comercial === selectedComercial);
  return row ? row.count : 0;
}, [offersModel, data, selectedComercial]);

  // Año del período (formato YYYY-MM)
  const yearForOffers = useMemo(
    () => Number((data?.period || "").slice(0, 4)) || new Date().getFullYear(),
    [data?.period]
  );

  // Metas del Sheet en un mapa por comercial (normalizado)
  const [metasMap, setMetasMap] = React.useState<Map<string, MetaRecord>>(new Map());
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { metas } = await fetchMetas(yearForOffers); // ya existe en tu archivo
        const map = new Map<string, MetaRecord>();
        metas.forEach(m => map.set(normName(m.comercial), m)); // normName ya existe
        if (!cancelled) setMetasMap(map);
      } catch {
        if (!cancelled) setMetasMap(new Map());
      }
    })();
    return () => { cancelled = true; };
  }, [yearForOffers]);

  // Meta Ofertas del comercial seleccionado (desde Sheet)
  const targetSelected = metasMap.get(normName(selectedComercial))?.metaOfertas ?? 0;

  // Para el ancho de barra: tope visual (no depende de meta global ya)
  const max = useMemo(
    () => data.porComercial.reduce((m: number, x: any) => Math.max(m, x.count), 0) || 1,
    [data]
  );

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Ofertas (desde DETALLADO)" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
<section className="p-4 bg-white rounded-xl border">
  <div className="flex flex-col md:flex-row md:items-center md:gap-4">
    <div className="text-sm text-gray-500">Comercial: <b>{selectedComercial}</b></div>

    <div className="text-sm text-gray-500">Periodo:
      <select
        className="ml-2 border rounded px-2 py-1 text-sm"
        value={offersPeriod}
        onChange={(e) => setOffersPeriod(e.target.value)}
      >
        {(data.periods || []).map((p: string) => <option key={p} value={p}>{p}</option>)}
      </select>
    </div>

    <div className="text-sm text-gray-500">
      Meta mensual (Sheet): <b className="tabular-nums">{targetSelected}</b>
    </div>
  </div>

  <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
    <StatCard label="Ofertas del período (compañía)">{data.total}</StatCard>
    <StatCard label="Del comercial seleccionado">{selected}</StatCard>
    <StatCard label="Cumplimiento (vs meta del Sheet)">
      {(() => {
        const pct = targetSelected > 0 ? Math.round((selected / targetSelected) * 100) : (selected > 0 ? 100 : 100);
        const st  = offerStatus(selected, targetSelected);
        return (
          <span className="flex items-center gap-2">
            <span>{pct}% ({selected}/{targetSelected})</span>
            <span className={`inline-block w-3 h-3 rounded-full ${st.dot}`} />
          </span>
        );
      })()}
    </StatCard>
  </div>

  <div className="text-xs text-gray-500 mt-2">
    Fuente: Archivo DETALLADO (una fila = una oferta). Requiere columnas: <em>Comercial</em> y <em>Fecha de oferta</em>.
  </div>
</section>

        {offersModel && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">Ranking de ofertas por comercial ({data.period})</div>
            <div className="space-y-2">
              {onlySelected(data.porComercial, selectedComercial).map((row: any, i: number) => {
                const target = metasMap.get(normName(row.comercial))?.metaOfertas ?? 0;
                const pct = target > 0
                  ? Math.round((row.count / target) * 100)
                  : (row.count > 0 ? 100 : 100);
                const pctBar = Math.min(100, pct);
                const st = offerStatus(row.count, target);

                return (
                  <div key={row.comercial} className="text-sm">
                    <div className="flex items-center justify-between gap-2">
                      <div className="font-medium">{i + 1}. {row.comercial}</div>
                      <div className="flex items-center gap-2">
                        <span className="tabular-nums text-gray-900">
                          {pct}% ({row.count}/{target})
                        </span>
                        <span className={`inline-block w-2 h-2 rounded-full ${st.dot}`} />
                      </div>
                    </div>
                    <div className="h-2 bg-gray-200 rounded mt-1">
                      <div className="h-2 rounded bg-gray-700" style={{ width: pctBar + "%" }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        )}
      </main>
    </div>
  );
};

// === ScreenVisits (solo visitas + meeting) ===
const ScreenActivities = () => {
  const [mode, setMode] = React.useState<"completadas" | "vencidas" | "pendientes">("completadas");

  const data = React.useMemo(() => {
    if (!activitiesModel) return { total: 0, porComercial: [] as any[] };
    const rows = activitiesModel.rows || [];

  // Totales por comercial (todas las actividades) para el denominador del %
const totalsByCom = React.useMemo(() => {
  const m = new Map<string, number>();
  const rows = activitiesModel?.rows || [];
  for (const r of rows) {
    if (!r?.comercial || r.comercial === "(Sin comercial)") continue; // coherente con el parser
    m.set(r.comercial, (m.get(r.comercial) || 0) + 1);
  }
  return m;
}, [activitiesModel]);

    const by = new Map<string, number>();
    for (const r of rows) {
      if (r.status !== mode) continue;                 // filtra por estado elegido
      by.set(r.comercial, (by.get(r.comercial) || 0) + 1);
    }

    const porComercial = Array.from(by.entries())
      .map(([comercial, count]) => ({ comercial, count }))
      .sort((a, b) => b.count - a.count);

    const total = porComercial.reduce((a, x) => a + x.count, 0);
    return { total, porComercial };
  }, [activitiesModel, mode]);

  const selectedCount = React.useMemo(() => {
    if (selectedComercial === "ALL") return data.total;
    return data.porComercial.find((r: any) => r.comercial === selectedComercial)?.count ?? 0;
  }, [data, selectedComercial]);

  const label =
    mode === "completadas" ? "Completadas"
    : mode === "vencidas" ? "Vencidas"
    : "Pendientes";

  // Totales del archivo para el denominador (completadas/total, etc.)
  const totalFileCount = React.useMemo(() => Number(activitiesModel?.rows?.length || 0), [activitiesModel]);

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Actividades" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        <section className="p-4 bg-white rounded-xl border">
          <div className="flex flex-col md:flex-row md:items-center md:gap-4">
            <div className="text-sm text-gray-500">
              Comercial: <b>{selectedComercial}</b>
            </div>

            {/* Botonera tipo 'Ciclo de venta' */}
            <div className="md:ml-auto">
              <div className="inline-flex rounded-lg border overflow-hidden">
                <button
                  className={`px-3 py-1 text-sm ${mode === "completadas" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("completadas")}
                >
                  Completadas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${mode === "vencidas" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("vencidas")}
                >
                  Vencidas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${mode === "pendientes" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("pendientes")}
                >
                  Pendientes
                </button>
              </div>
            </div>
          </div>

          {/* Tarjetas superiores */}
          <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
            <StatCard label={`${label} (compañía)`}>
              {data.total} / {totalFileCount}
            </StatCard>
            <StatCard label="Del comercial seleccionado">
              {selectedCount}
            </StatCard>
            <StatCard label="Meta">{/* sin meta por ahora */}—</StatCard>
          </div>

<div className="text-xs text-gray-500 mt-2">
  Fuente: archivo <b>ACTIVIDADES</b>. Se clasifica por <em>fechas</em> (fin/vencimiento). Si no hay fechas, se usa <em>Estado/Status</em>.
</div>
        </section>

        {/* Ranking */}
        {activitiesModel && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">
              Ranking por comercial ({label.toLowerCase()})
            </div>
            <div className="space-y-2">
{onlySelected(data.porComercial, selectedComercial).map((row: any, i: number) => {
  const denom = totalsByCom.get(row.comercial) || 0; // total del comercial
  const pct = denom > 0 ? Math.round((row.count / denom) * 100) : 0; // % del comercial
  return (
    <div key={row.comercial} className="text-sm">
      <div className="flex items-center justify-between gap-2">
        <div className="font-medium">{i + 1}. {row.comercial}</div>
        <div className="tabular-nums text-gray-900">
          {row.count} ({pct}%)
        </div>
      </div>
      <div className="h-2 bg-gray-200 rounded mt-1">
        <div className="h-2 rounded bg-gray-700" style={{ width: Math.min(100, pct) + "%" }} />
      </div>
    </div>
  );
})}
            </div>
          </section>
        )}
      </main>
    </div>
  );
};

// === ScreenVisits (llamadas/visitas/reuniones desde ASUNTO) ===
const ScreenVisits = () => {
  const [mode, setMode] = React.useState<"llamadas" | "visitas" | "reuniones">("visitas");

  // Filas del período seleccionado
  const periodRows = React.useMemo(() => {
    if (!visitsModel) return [] as any[];
    const periods: string[] = visitsModel.periods || [];
    const sel = periods.includes(visitsPeriod) ? visitsPeriod : (periods[periods.length - 1] || "");
    return (visitsModel.rows || []).filter((r: any) => r.ym === sel);
  }, [visitsModel, visitsPeriod]);

  // Agregación por modo (usa r.kind que se armó desde "Asunto")
  const data = React.useMemo(() => {
    const by = new Map<string, number>();
    for (const r of periodRows) {
      if ((r.kind || "otros") !== mode) continue;
      by.set(r.comercial, (by.get(r.comercial) || 0) + 1);
    }
    const porComercial = Array.from(by.entries())
      .map(([comercial, count]) => ({ comercial, count }))
      .sort((a, b) => b.count - a.count);
    const total = porComercial.reduce((a, x) => a + x.count, 0);

    const sel = (visitsModel?.periods || []).includes(visitsPeriod)
      ? visitsPeriod
      : ((visitsModel?.periods || []).slice(-1)[0] || "");

    return { total, porComercial, period: sel, periods: visitsModel?.periods || [] };
  }, [periodRows, mode, visitsModel, visitsPeriod]);

  const selectedCount = React.useMemo(() => {
    if (selectedComercial === "ALL") return data.total;
    return data.porComercial.find((r: any) => r.comercial === selectedComercial)?.count ?? 0;
  }, [data, selectedComercial]);

  const labelForMode =
    mode === "llamadas" ? "Llamadas" : mode === "reuniones" ? "Reuniones" : "Visitas";

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • visitas" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        <section className="p-4 bg-white rounded-xl border">
          <div className="flex flex-col md:flex-row md:items-center md:gap-4">
            <div className="text-sm text-gray-500">
              Comercial: <b>{selectedComercial}</b>
            </div>

            <div className="text-sm text-gray-500">
              Periodo:
              <select
                className="ml-2 border rounded px-2 py-1 text-sm"
                value={visitsPeriod}
                onChange={(e) => setVisitsPeriod(e.target.value)}
              >
                {(data.periods || []).map((p: string) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </div>

            {/* Botones como en ciclo de ventas */}
            <div className="md:ml-auto">
              <div className="inline-flex rounded-lg border overflow-hidden">
                <button
                  className={`px-3 py-1 text-sm ${mode === "llamadas" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("llamadas")}
                >
                  Llamadas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${mode === "visitas" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("visitas")}
                >
                  Visitas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${mode === "reuniones" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setMode("reuniones")}
                >
                  Reuniones
                </button>
              </div>
            </div>
          </div>

          <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
            <StatCard label={`${labelForMode} del período (compañía)`}>{data.total}</StatCard>
            <StatCard label="Del comercial seleccionado">{selectedCount}</StatCard>
            <StatCard label="Meta">—</StatCard>
          </div>

          <div className="text-xs text-gray-500 mt-2">
            Fuente: archivo de eventos. Se clasifica por la columna <em>Asunto</em>.
          </div>
        </section>

        {/* Ranking */}
        {visitsModel && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">
              Ranking por comercial ({labelForMode.toLowerCase()}) · {data.period}
            </div>
            <div className="space-y-2">
              {onlySelected(data.porComercial, selectedComercial).map((row: any, i: number) => {
                return (
                  <div key={row.comercial} className="text-sm">
                    <div className="flex items-center justify-between gap-2">
                      <div className="font-medium">{i + 1}. {row.comercial}</div>
                      <div className="tabular-nums text-gray-900">{row.count}</div>
                    </div>
                    <div className="h-2 bg-gray-200 rounded mt-1">
                      <div className="h-2 rounded bg-gray-700" style={{ width: (row.count / Math.max(1, data.total)) * 100 + "%" }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        )}
      </main>
    </div>
  );
};

const ScreenCycle = () => {
  // Usa el memo que arma los datos según cycleMode
  const cd = useMemo(() => cycleData, [cycleData]);

  // Color del semáforo según meta de días
  const colorDays = (d: number) =>
    d <= cycleTarget ? "bg-green-500" :
    d <= cycleTarget * 1.2 ? "bg-yellow-400" : "bg-red-500";

  // Valor grande de la tarjeta superior (siempre promedio de días)
const headerValue = useMemo(() => {
  if (!detail || !cd?.data) return 0;
  if (selectedComercial === "ALL") return cd.data.totalAvgDays || 0;
  const row = (cd.data.porComercial || []).find((r: any) => r.comercial === selectedComercial);
  return row ? (row.avgDays || 0) : 0;
}, [detail, cd, selectedComercial]);

  // Máximo para escalar barras
  const maxBar = useMemo(() => {
    if (!cd?.data) return 1;
    const arr = cd.data.porComercial || [];
    return Math.max(cd.data.totalAvgDays || 0, ...arr.map((r: any) => r.avgDays || 0)) || 1;
  }, [cd]);

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Ciclo de venta (días)" />

      <main className="max-w-6xl mx-auto p-4 space-y-6">
        {/* Tarjeta superior */}
<section className="p-4 bg-white rounded-xl border">
  <div className="flex flex-col md:flex-row md:items-center md:gap-4">
    <div className="text-sm text-gray-500">Comercial: <b>{selectedComercial}</b></div>
    <div className="text-sm text-gray-500">Meta (días): <b>{cycleTarget}</b></div>
  </div>

  {(() => {
    const totalAvg = Math.round(cd?.data?.totalAvgDays || 0);
    const rowSel = selectedComercial === "ALL"
      ? null
      : (cd?.data?.porComercial || []).find((r: any) => r.comercial === selectedComercial);
    const selAvg = Math.round(rowSel?.avgDays || 0);

    // Cumplimiento: a menos días es mejor → target / promedio (cap 200%)
    const cumpl = selAvg > 0 ? Math.min(200, Math.round((cycleTarget / selAvg) * 100)) : 0;

    return (
      <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
        <StatCard label="Promedio (compañía)">{totalAvg} días</StatCard>
        <StatCard label="Promedio (seleccionado)">
          <span className="flex items-center gap-2">
            <span>{selAvg} días</span>
            <span className={`inline-block w-3 h-3 rounded-full ${colorDays(selAvg)}`} />
          </span>
        </StatCard>
        <StatCard label="Cumplimiento vs meta (días)">
          {cumpl}%
        </StatCard>
      </div>
    );
  })()}
</section>
        
        {/* Selector de modo */}
        {detail && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">Sales Cycle por comercial</div>

            <div className="flex items-center gap-2 mb-3">
              <span className="text-sm text-gray-600">Modo:</span>
              <div className="inline-flex rounded-lg border overflow-hidden">
                <button
                  className={`px-3 py-1 text-sm ${cycleMode === "all" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setCycleMode("all")}
                  title="Promedio de días (Closed Won + Closed Lost)"
                >
                  Todas cerradas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${cycleMode === "won" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setCycleMode("won")}
                  title="Promedio de días (solo Closed Won)"
                >
                  Solo ganadas
                </button>
                <button
                  className={`px-3 py-1 text-sm border-l ${cycleMode === "offers" ? "bg-gray-900 text-white" : "bg-white"}`}
                  onClick={() => setCycleMode("offers")}
                  title="Promedio de días de todas las oportunidades (abiertas + cerradas)"
                >
                  Todas las ofertas
                </button>
              </div>
            </div>

            {/* Lista por comercial */}
            <div className="space-y-2">
              {onlySelected((cd?.data?.porComercial || []), selectedComercial).map((row: any) => {
                const pct = Math.round(((row.avgDays || 0) / (maxBar || 1)) * 100);
                return (
                  <div key={row.comercial} className="text-sm">
                    <div className="flex justify-between items-center">
                      <span className="font-medium">{row.comercial}</span>
                      <span className="flex items-center gap-2">
                        <span className={`inline-block w-2 h-2 rounded-full ${colorDays(row.avgDays || 0)}`}></span>
                        <span className="tabular-nums text-gray-900">
                          {Math.round(row.avgDays || 0)} días (n={row.n})
                        </span>
                      </span>
                    </div>
                    <div className="h-2 bg-gray-200 rounded">
                      <div
                        className="h-2 rounded bg-gray-700"
                        style={{ width: pct + "%" }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        )}
      </main>
    </div>
  );
};

  // ===== KPI: Cumplimiento de Meta (Anual) =====
const offersKPI = useMemo(() => {
  if (!offersModel) return { total: 0, porComercial: [], periods: [], period: "" };
  const periods = offersModel.periods || [];
  const sel = periods.includes(offersPeriod) ? offersPeriod : (periods[periods.length - 1] || "");
  const rows = offersModel.rows.filter((r:any) => r.ym === sel);

  const by = new Map<string, number>();
  for (const r of rows) {
    const key = r.comercial; // ya viene normalizado
    by.set(key, (by.get(key) || 0) + 1);
  }
  const porComercial = Array.from(by.entries())
    .map(([comercial, count]) => ({ comercial, count }))
    .sort((a,b)=> b.count - a.count);

  const total = porComercial.reduce((a,x)=>a+x.count, 0);
  return { total, porComercial, periods, period: sel };
}, [offersModel, offersPeriod]);


// === ScreenAttainment ===
const ScreenAttainment = () => {
  // `pivot` debe existir (viene del RESUMEN que ya cargas en tu app)
  if (!pivot) {
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Cumplimiento de Meta (Anual)" />
        <main className="max-w-6xl mx-auto p-4">
          <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg text-yellow-800">
            Carga primero el archivo <b>RESUMEN</b> para ver este KPI.
          </div>
        </main>
      </div>
    );
  }

  // ================== Metas desde Sheet (año = settingsYear) ==================
  const [metasMap, setMetasMap] = React.useState<Map<string, any>>(new Map());
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        // fetchMetas(year) ya existe en tu archivo y devuelve { year, metas }
        const { metas } = await fetchMetas(settingsYear);
        const map = new Map<string, any>();
        metas.forEach((m: any) => map.set(normName(m.comercial), m)); // normName ya existe
        if (!cancelled) setMetasMap(map);
      } catch {
        if (!cancelled) setMetasMap(new Map());
      }
    })();
    return () => { cancelled = true; };
  }, [settingsYear]);

  // Helper: meta anual del comercial (si no hay en hoja, 0)
  const goalFor = React.useCallback((com: string) => {
    return metasMap.get(normName(com))?.metaAnual ?? 0;
  }, [metasMap]);

  // ================== Construir KPI desde el pivot ==================
  type RowAtt = { comercial: string; wonCOP: number; goal: number; pct: number };

const kpi = React.useMemo(() => {
  return calcAttainmentFromPivot(pivot, goalFor);
}, [pivot, goalFor]);

  // Comercial seleccionado
  const selectedAtt = React.useMemo(() => {
    if (selectedComercial === "ALL") return null;
    const row = kpi.porComercial.find(r => r.comercial === selectedComercial);
    return row || { comercial: selectedComercial, wonCOP: 0, goal: goalFor(selectedComercial), pct: 0 };
  }, [kpi, selectedComercial, goalFor]);

  // Formateo COP local (evita depender de util externo)
  const fmtCOP = (n: number) =>
    (Number(n) || 0).toLocaleString("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Cumplimiento de Meta (Anual)" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        {/* Header / filtros */}
<section className="p-4 bg-white rounded-xl border">
  <div className="flex flex-col md:flex-row md:items-center md:gap-4">
    <div className="text-sm text-gray-500">Comercial: <b>{selectedComercial}</b></div>
    <div className="text-sm text-gray-500">Año (Sheet): <b>{settingsYear}</b></div>
  </div>

  <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
    <StatCard label="Cumplimiento (compañía)">
      {Math.round(kpi.total.pct)}% ({fmtCOP(kpi.total.wonCOP)} / {fmtCOP(kpi.total.goal)})
    </StatCard>

    <StatCard label="Cerrado del seleccionado">
      {selectedComercial === "ALL"
        ? "—"
        : `${fmtCOP(selectedAtt?.wonCOP ?? 0)} / ${fmtCOP(selectedAtt?.goal ?? 0)}`}
    </StatCard>

    <StatCard label="Cumplimiento (seleccionado)">
      {selectedComercial === "ALL" ? "—" : `${Math.round(selectedAtt?.pct ?? 0)}%`}
    </StatCard>
  </div>
</section>
        
        {/* Ranking por comercial */}
        <section className="p-4 bg-white rounded-xl border">
          <div className="mb-3 font-semibold">Ranking por comercial (cumplimiento vs meta anual del Sheet)</div>
          <div className="space-y-2">
            {onlySelected(kpi.porComercial, selectedComercial).map((row, i) => {
              const pct = Math.round(row.pct);
              const pctBar = Math.min(100, pct);
              const st = offerStatus(row.wonCOP, row.goal);
              return (
                <div key={row.comercial} className="text-sm">
                  <div className="flex items-center justify-between gap-2">
                    <div className="font-medium">{i + 1}. {row.comercial}</div>
                    <div className="flex items-center gap-2">
                      <span className="tabular-nums text-gray-900">
                        {pct}% ({fmtCOP(row.wonCOP)} / {fmtCOP(row.goal)})
                      </span>
                      <span className={`inline-block w-2 h-2 rounded-full ${st.dot}`} />
                    </div>
                  </div>
                  <div className="h-2 bg-gray-200 rounded mt-1">
                    <div className="h-2 rounded bg-gray-700" style={{ width: pctBar + "%" }} />
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      </main>
    </div>
  );
};

  // ========= Router principal: una pantalla a la vez =========
  if (route === "HOME") return <RouteHome onEnter={() => setRoute("MENU")} />;
  if (route === "KPI_PIPELINE") return <ScreenPipeline />;
  if (route === "KPI_WINRATE") return <ScreenWinRate />;
  if (route === "KPI_CYCLE") return <ScreenCycle />;
  if (route === "KPI_ATTAIN") return <ScreenAttainment />;
  if (route === "KPI_OFFERS") return <ScreenOffers />;
  if (route === "KPI_VISITS") return <ScreenVisits />;
  if (route === "KPI_ACTIVITIES") return <ScreenActivities />;

  // ===== Menú ===== (route === "MENU")
  return (
    <div className="min-h-screen bg-gray-50">
      <header className="px-4 py-3 bg-white border-b sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-3">
          <h2 className="text-xl md:text-2xl font-bold">Menú principal</h2>
<div className="flex items-center gap-2">
  <button
    className="px-4 py-2 text-sm font-medium rounded-lg border bg-white hover:bg-gray-50"
    onClick={resetAll}
  >
    Reiniciar
  </button>
  <button
    className="px-4 py-2 text-sm font-medium rounded-lg border bg-white hover:bg-gray-50"
    onClick={openSettings}
    title="Editar metas por comercial"
  >
    Ajustes
  </button>
</div>
        </div>
      </header>
{showSettings && (
  <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4">
    <div className="w-full max-w-3xl bg-white rounded-xl shadow-lg p-4">
      <div className="flex items-center justify-between mb-2">
        <h3 className="text-lg font-semibold">Ajustes de metas por comercial</h3>
        <button
          className="text-sm px-2 py-1 rounded hover:bg-gray-100"
          onClick={() => setShowSettings(false)}
        >
          Cerrar
        </button>
      </div>

      <div className="flex items-center gap-2 mb-3">
        <span className="text-sm text-gray-600">Año:</span>
        <input
          type="number"
          className="w-28 border rounded px-2 py-1 text-sm"
          value={settingsYear}
          onChange={(e) => setSettingsYear(Number(e.target.value))}
          onBlur={openSettings} // recarga metas al cambiar de año
        />
        {loadingSettings && <span className="text-xs text-gray-500">Cargando…</span>}
      </div>

      <div className="overflow-auto max-h-[60vh]">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-gray-500">
              <th className="py-2 pr-2">Comercial</th>
              <th className="py-2 pr-2">Meta anual</th>
              <th className="py-2 pr-2">Meta ofertas</th>
              <th className="py-2 pr-2">Meta visitas</th>
            </tr>
          </thead>
          <tbody>
            {settingsRows.map((r, idx) => (
              <tr key={r.comercial} className="border-t">
                <td className="py-2 pr-2">{r.comercial}</td>
                <td className="py-2 pr-2">
                  <input
                    type="number"
                    className="w-40 border rounded px-2 py-1 text-sm text-right"
                    value={r.metaAnual}
                    onChange={(e) => {
                      const v = Number(e.target.value);
                      setSettingsRows(prev => {
                        const copy = [...prev];
                        copy[idx] = { ...copy[idx], metaAnual: isNaN(v) ? 0 : v };
                        return copy;
                      });
                    }}
                  />
                </td>
                <td className="py-2 pr-2">
                  <input
                    type="number"
                    className="w-24 border rounded px-2 py-1 text-sm text-right"
                    value={r.metaOfertas}
                    onChange={(e) => {
                      const v = Number(e.target.value);
                      setSettingsRows(prev => {
                        const copy = [...prev];
                        copy[idx] = { ...copy[idx], metaOfertas: isNaN(v) ? 0 : v };
                        return copy;
                      });
                    }}
                  />
                </td>
                <td className="py-2 pr-2">
                  <input
                    type="number"
                    className="w-24 border rounded px-2 py-1 text-sm text-right"
                    value={r.metaVisitas}
                    onChange={(e) => {
                      const v = Number(e.target.value);
                      setSettingsRows(prev => {
                        const copy = [...prev];
                        copy[idx] = { ...copy[idx], metaVisitas: isNaN(v) ? 0 : v };
                        return copy;
                      });
                    }}
                  />
                </td>
              </tr>
            ))}
            {!loadingSettings && settingsRows.length === 0 && (
              <tr>
                <td colSpan={4} className="py-6 text-center text-gray-500">
                  No hay comerciales detectados todavía. Carga un archivo o ingresa metas en tu Sheet.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="mt-3 flex justify-end gap-2">
        <button
          className="px-3 py-1.5 text-sm rounded-lg border hover:bg-gray-50"
          onClick={() => setShowSettings(false)}
        >
          Cancelar
        </button>
        <button
          className="px-3 py-1.5 text-sm rounded-lg text-white bg-gray-900 disabled:opacity-50"
          disabled={savingSettings}
          onClick={saveSettings}
        >
          {savingSettings ? "Guardando…" : "Guardar cambios"}
        </button>
      </div>
    </div>
  </div>
)}
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        {(error || info) && (
          <div className="space-y-2">
            {error && <div className="p-3 rounded border border-red-300 bg-red-50 text-sm text-red-700 whitespace-pre-wrap">{error}</div>}
          </div>
        )}

        {/* Comercial + Metas */}
        <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Comercial a evaluar</div>
            <select className="mt-2 w-full border rounded px-2 py-1" value={selectedComercial} onChange={(e) => setSelectedComercial(e.target.value)}>
              {comercialesMenu.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </section>

        {/* Cargar informes */}
        <section className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="p-4 bg-white rounded-xl border">
            <div className="font-semibold">Archivo RESUMEN (tabla dinámica)</div>
            <input
              type="file"
              accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
              onChange={(e) => e.target.files && onPivotFile(e.target.files[0])}
              className="block text-sm"
            />
            <div className="text-xs text-gray-500 mt-1">{filePivotName || "Sin archivo"}</div>
          </div>
          <div className="p-4 bg-white rounded-xl border">
            <div className="font-semibold">Archivo DETALLADO</div>
            <input type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv" onChange={(e) => e.target.files && onDetailFile(e.target.files[0])} className="block text-sm" />
            <div className="text-xs text-gray-500 mt-1">{fileDetailName || "Sin archivo"}</div>
          </div>
         <div className="p-4 bg-white rounded-xl border">
          <div className="font-semibold">Archivo VISITAS</div>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
            onChange={(e) => e.target.files && onVisitsFile(e.target.files[0])}
            className="block text-sm"
          />
          <div className="text-xs text-gray-500 mt-1">{fileVisitsName || "Sin archivo"}</div>
        </div>
        <div className="p-4 bg-white rounded-xl border">
  <div className="font-semibold">Archivo ACTIVIDADES</div>
  <input
    type="file"
    accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
    onChange={(e) => e.target.files && onActivitiesFile(e.target.files[0])}
    className="block text-sm"
  />
  <div className="text-xs text-gray-500 mt-1">{fileActivitiesName || "Sin archivo"}</div>
</div>

      </section>

        {/* Tarjetas de acceso a KPIs */}
        <section className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="p-4 bg-white rounded-xl border flex flex-col">
          <div className="font-semibold">📊 Forecast para cumplir meta</div>
          <p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN + Meta anual + Win Rate 20%</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_PIPELINE")} disabled={!pivot}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">🎯 Tasa de Cierre (Win Rate)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_WINRATE")} disabled={!pivot}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">⏱️ Ciclo de venta (días)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: DETALLADO</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_CYCLE")} disabled={!detail}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">🏁 Cumplimiento de Meta (Anual)</div>
<p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN + Meta anual (Sheet)</p>
<button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_ATTAIN")} disabled={!pivot}>Ver KPI</button>
          </div>
        </section>
        <section className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">🧾 Ofertas</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: DETALLADO (fecha de oferta)</p>
            <button
              className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40"
              onClick={() => setRoute("KPI_OFFERS")}
              disabled={!offersModel}
            >
              Ver KPI
            </button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">📅 Visitas</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: archivo VISITAS</p>
            <button
              className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40"
              onClick={() => setRoute("KPI_VISITS")}
              disabled={!visitsModel}
            >
              Ver KPI
            </button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
  <div className="font-semibold">📌 Actividades</div>
  <p className="text-xs text-gray-500 mt-1">Fuente: archivo VISITAS (Asunto)</p>
<button
  className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40"
  onClick={() => setRoute("KPI_ACTIVITIES")}
  disabled={!activitiesModel}
>
  Ver KPI
</button>
</div>
        </section>
      </main>
    </div>
  );
}
