import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

// ========================= Utils =========================
const norm = (s: any) => String(s ?? "")
  .replace(/\u00A0/g, " ")
  .toLowerCase()
  .normalize("NFD")
  .replace(/\p{Diacritic}/gu, "")
  .replace(/\s+/g, " ")
  .trim();

const toNumber = (v: any) => {
  if (v == null || v === "") return 0;
  let s = String(v).trim();
  s = s.replace(/[^\d,.-]/g, "");
  if (s.includes(",") && !s.includes(".")) s = s.replace(/,/g, ".");
  if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, "");
  const n = Number(s);
  return isFinite(n) ? n : 0;
};

// Semáforo por cumplimiento vs meta (target)
function offerStatus(count: number, target: number) {
  const ratio = target > 0 ? count / target : 0; // 1.0 = 100%
  if (ratio >= 1) return { ratio, bar: "bg-green-500", dot: "bg-green-500", text: "text-green-600" };     // verde
  if (ratio >= 0.8) return { ratio, bar: "bg-yellow-400", dot: "bg-yellow-400", text: "text-yellow-600" }; // amarillo (70%-99%)
  return { ratio, bar: "bg-red-500", dot: "bg-red-500", text: "text-red-600" };                             // rojo (<70%)
}

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

const daysBetween = (d1: Date | null, d2: Date | null) => {
  if (!d1 || !d2) return null as any;
  const ms = d2.getTime() - d1.getTime();
  return Math.max(0, Math.round(ms / (1000 * 60 * 60 * 24)));
};

const fmtCOP = (n: number) => n.toLocaleString("es-CO");

const OPEN_STAGES = ["qualification", "needs analysis", "needs", "proposal", "negotiation"];
const WON_STAGES = ["closed won", "ganad"]; // detectar 'Closed Won' y variantes en español

// Metas mensuales (COP) -> se convierten a meta anual multiplicando por 12
const METAS_MENSUALES: Record<string, number> = {
  "CLAUDIA RODRIGUEZ RODRIGUEZ": 66666000,
  "HERNAN ROLDAN": 37500000,
  "JHOAN ORTIZ": 198333000,
  "Juan Garzón Linares": 216666000,
  "KAREN CARRILLO": 91666000,
  "LIZETH MARTINEZ": 83333000,
  "PABLO RODRIGUEZ RODRIGUEZ": 200000000,
};
const metaAnual = (com: string) => (METAS_MENSUALES[com] || 0) * 12;

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
  "juan garzon": "Juan Garzón Linares",
  "juan garzon linares": "Juan Garzón Linares",
  "juan garzón linares": "Juan Garzón Linares",
  "juan sebastian garzon": "Juan Garzón Linares",
  "juan sebastian garzon linares": "Juan Garzón Linares",

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
  if (looksZip(u8) || looksOLE(u8)) {
    try { return XLSX.read(u8, { type: "array", dense: true }); }
    catch (e: any) {
      const msg = String(e?.message || "");
      if (/bad uncompressed size/i.test(msg) || /End of data reached/i.test(msg)) {
        try { const text = await file.text(); return XLSX.read(text, { type: "string", dense: true }); } catch {}
        try {
          const bin = Array.from(u8).map(b => String.fromCharCode(b)).join("");
          const b64 = typeof btoa !== 'undefined' ? btoa(bin) : (typeof Buffer !== 'undefined' ? Buffer.from(bin, 'binary').toString('base64') : "");
          if (b64) return XLSX.read(b64, { type: "base64", dense: true });
        } catch {}
        throw new Error("El archivo parece .xlsx pero no se pudo descomprimir. Reexporta o sube CSV.");
      }
      throw e;
    }
  }
  try { const text = await file.text(); return XLSX.read(text, { type: "string", dense: true }); } catch {}
  try {
    const bin = Array.from(u8).map(b => String.fromCharCode(b)).join("");
    const b64 = typeof btoa !== 'undefined' ? btoa(bin) : (typeof Buffer !== 'undefined' ? Buffer.from(bin, 'binary').toString('base64') : "");
    if (b64) return XLSX.read(b64, { type: "base64", dense: true });
  } catch {}
  throw new Error("No se pudo leer el archivo. Sube .xlsx/.xlsm/.xlsb/.xls o .csv válido.");
}

function parseOffersFromDetailSheet(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];
  if (!A.length) throw new Error("DETALLADO (Ofertas): hoja vacía");

  // --- helpers ---
  const n = (s: any) =>
    String(s ?? "")
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "") // tildes
      .replace(/[()↑%]/g, " ")         // símbolos del reporte
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

  const scoreHead = (row: any[]) => {
    const H = row.map(n);
    let sc = 0;
    if (H.some(h => h.includes("propietario") || h.includes("comercial") || h.includes("owner") || h.includes("vendedor"))) sc++;
    if (H.some(h => h.includes("fecha"))) sc++;
    if (H.some(h => h.includes("oportunidad") || h.includes("nombre"))) sc++;
    if (H.some(h => h.includes("valor") || h.includes("monto") || h.includes("importe") || h.includes("precio") || h.includes("amount"))) sc++;
    return sc;
  };

  // localizar fila de encabezados
  let headerRow = 0, best = -1;
  for (let r = 0; r < Math.min(40, A.length); r++) {
    const sc = scoreHead(A[r] || []);
    if (sc > best) { best = sc; headerRow = r; }
  }

  const headers = A[headerRow] || [];
  const findIdx = (...cands: string[]) => {
    const nc = cands.map(n);
    for (let c = 0; c < headers.length; c++) {
      const h = n(headers[c]);
      if (nc.some(k => h.includes(k))) return c;
    }
    return -1;
  };

  const idxCom = findIdx("propietario de oportunidad", "comercial", "propietario", "owner", "vendedor");
  const idxFec = findIdx(
    "fecha de oferta", "fecha oferta", "fecha de envio", "fecha envio", "fecha propuesta",
    "fecha de creacion", "fecha creacion", "fecha de creación", "fecha creación",
    "created", "close date", "fecha"
  );
  const idxNom = findIdx("nombre de la oportunidad", "oportunidad", "nombre", "asunto", "subject");
  const idxVal = findIdx("valor", "monto", "importe", "amount", "precio total", "total");

  if (idxCom < 0 || idxFec < 0) {
    throw new Error(`DETALLADO (Ofertas): faltan columnas mínimas (Comercial y Fecha) en hoja ${sheetName}`);
  }

  const rows: any[] = [];
  let currentComercial = ""; // ← aquí se arrastra el comercial del bloque

  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || [];
    if (row.every((v: any) => String(v).trim() === "")) continue;

    // saltar filas de subtotal/total/recuento/suma propias del reporte
    const line = n((row.join(" ")) || "");
    if (
      line.startsWith("subtotal") || line.startsWith("total") ||
      line.includes("recuento") || line.includes("suma de")
    ) continue;

    // si trae comercial explícito en esta fila, actualizar el "current"
// si la celda trae comercial explícito (texto NO vacío), actualiza el "current"
const rawCom = row[idxCom];
if (rawCom != null && String(rawCom).trim() !== "") {
  const mapped = mapComercial(rawCom);
  if (mapped && mapped !== "(Sin comercial)") {
    currentComercial = mapped;
  }
}

// usar el arrastrado; si aún no hay, no cuentes la fila (sigue siendo cabecera/total)
const comercial = currentComercial;
if (!comercial || comercial === "(Sin comercial)") continue;


    // fecha robusta (dd/mm/yyyy, serial de Excel, ISO, etc.)
    const fecha = parseDateCell(row[idxFec]);
    if (!fecha) continue;

    // periodo UTC (evita deslizar de mes por zona horaria)
    const d = new Date(fecha);
    const ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;

    const nombre = idxNom >= 0 ? String(row[idxNom] ?? "").trim() : "";
    const valor = idxVal >= 0 ? Number(row[idxVal] ?? 0) : 0;

    rows.push({ comercial, fecha: d, ym, nombre, valor });
  }

  if (!rows.length) throw new Error(`DETALLADO (Ofertas): sin filas válidas en ${sheetName}`);
  const periods = Array.from(new Set(rows.map(r => r.ym))).sort(); // solo meses (nada de "ALL")
  return { rows, sheetName, periods };
}

function buildOffersModelFromDetail(wb: XLSX.WorkBook) {
  const errs: string[] = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      const parsed = parseOffersFromDetailSheet(ws, sn);
      // catálogo de periodos
      const periods = Array.from(new Set(parsed.rows.map(r => r.ym))).sort();
      return { ...parsed, periods };
    } catch (e:any) { errs.push(`${sn}: ${e?.message || e}`); }
  }
  throw new Error("DETALLADO (Ofertas): no pude interpretar ninguna hoja. " + errs.join(" | "));
}
// ==================== Parser RESUMEN ====================
function findHeaderPosition(A: any[][]) {
  for (let r = 0; r < A.length; r++) {
    for (let c = 0; c < (A[r]?.length || 0); c++) {
      if (norm(A[r][c]).includes("propietario de oportunidad")) return { headerRow: r, propietarioCol: c };
    }
  }
  throw new Error("Resumen: No se encontró 'Propietario de oportunidad' en el encabezado.");
}
function findStageHeaderRow(A: any[][], headerRow: number, startCol: number) {
  const looksStageRow = (ri: number) => {
    const row = A[ri] || []; let hits = 0;
    for (let c = startCol; c < row.length; c++) {
      const v = norm(row[c]); if (!v) continue;
      if (["qualification", "needs", "needs analysis", "proposal", "negotiation", "closed", "ganad", "perdid"].some(k => v.includes(k))) hits++;
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
function parsePivotSheet(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any;
  if (!A.length) throw new Error("Resumen: hoja vacía");
  const { headerRow, propietarioCol } = findHeaderPosition(A);
  const stageRow = findStageHeaderRow(A, headerRow, propietarioCol + 1);
  const cols: any[] = []; let lastStage = "";
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
  const names = (wb.Sheets as any)["Informe medicion KPI"] ? ["Informe medicion KPI", ...wb.SheetNames.filter(n => n !== "Informe medicion KPI")] : wb.SheetNames;
  const errs: string[] = [];
  for (const name of names) {
    try { const ws = wb.Sheets[name]; if (!ws) continue; return parsePivotSheet(ws, name); } catch (e: any) { errs.push(`${name}: ${e?.message || e}`); }
  }
  throw new Error("Resumen: ninguna hoja válida. Detalles: " + errs.join(" | "));
}

// ==================== Parser DETALLE ====================
function parseDetailSheetRobust(ws: XLSX.WorkSheet, sheetName: string) {
  const A: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any;
  if (!A.length) throw new Error("Detalle: hoja vacía");
  const hnorm = (s: any) => norm(String(s).replace(/[()↑%]/g, " "));
  const EXPECT = [
    "propietario de oportunidad","nombre de la oportunidad","importe","probabilidad","etapa",
    "fecha de creacion","fecha de cierre","antiguedad","nombre del producto","precio total",
    "descripcion del producto","nombre de la cuenta"
  ];
  const scoreRow = (row: any[]) => { const set = row.map(hnorm); let score = 0; for (const exp of EXPECT) if (set.some(h => h.includes(exp))) score++; return score; };
  let headerRow = 0, bestScore = -1; for (let r = 0; r < Math.min(A.length, 40); r++) { const sc = scoreRow(A[r] || []); if (sc > bestScore) { bestScore = sc; headerRow = r; } }
  if (bestScore < 3) headerRow = 0;
  const headersRaw = (A[headerRow] || []).slice();
  const dbg = [ `Fila header detectada: ${headerRow + 1}`, "Encabezados detectados: " + headersRaw.join(" | ") ];
  const findIdx = (cands: string[]) => { for (let c = 0; c < headersRaw.length; c++) { const h = hnorm(headersRaw[c]); for (const cand of cands) if (h === cand || h.includes(cand)) return c; } return -1; };
  const idxOwner   = findIdx(["propietario de oportunidad","propietario","owner"]);
  const idxStage   = findIdx(["etapa","stage","estado"]);
  const idxAging   = findIdx(["antiguedad","antigüedad","aging","days open"]);
  const idxCreated = findIdx(["fecha de creacion","fecha creacion","fecha de creación","created date","created"]);
  const idxClosed  = findIdx(["fecha de cierre","close date","closed date"]);
  const idxImporte = findIdx(["importe","monto","amount"]);
  const idxProb    = findIdx(["probabilidad","probability"]);
  const idxProd    = findIdx(["nombre del producto","producto"]);
  const idxPrecioT = findIdx(["precio total","total price"]);
  const idxCuenta  = findIdx(["nombre de la cuenta","cuenta","account name"]);

  if (idxOwner < 0 || idxStage < 0 || (idxAging < 0 && idxCreated < 0)) {
    dbg.push("Requisito mínimo: Comercial (Propietario de la oportunidad), Etapa y Antigüedad (o Fechas de creación/cierre).");
    const err: any = new Error("Detalle: no encontré columnas mínimas. Revisa el debug.");
    err.debug = dbg; throw err;
  }

  const out: any[] = [];
  for (let r = headerRow + 1; r < A.length; r++) {
    const row = A[r] || []; const isEmpty = row.every((v: any) => String(v).trim() === ""); if (isEmpty) continue;
    const get = (i: number) => (i >= 0 ? row[i] : "");
    const Comercial       = mapComercial(get(idxOwner));
    const Etapa           = String(get(idxStage) ?? "").trim();
    const Antiguedad      = idxAging >= 0 ? (() => { const v = get(idxAging); const n = Number(String(v).replace(/[^\d.-]/g, "")); return isFinite(n) ? n : null; })() : null;
    const Fecha_creacion  = idxCreated >= 0 ? parseDateCell(get(idxCreated)) : null;
    const Fecha_cierre    = idxClosed  >= 0 ? parseDateCell(get(idxClosed))  : null;
    const Importe         = idxImporte >= 0 ? toNumber(get(idxImporte)) : null;
    const Probabilidad    = idxProb    >= 0 ? toNumber(get(idxProb))    : null;
    const Producto        = idxProd    >= 0 ? String(get(idxProd) || "") : undefined;
    const Precio_total    = idxPrecioT >= 0 ? toNumber(get(idxPrecioT)) : null;
    const Cuenta          = idxCuenta  >= 0 ? String(get(idxCuenta) || "") : undefined;

    out.push({ Comercial, Etapa, Antiguedad, Fecha_creacion, Fecha_cierre, Importe, Probabilidad, Producto, Precio_total, Cuenta });
  }
  return { rows: out, sheetName, debug: dbg } as const;
}
function tryParseAnyDetail(wb: XLSX.WorkBook) {
  const errs: string[] = []; let lastDebug: string[] | null = null;
  for (const name of wb.SheetNames) {
    try { const ws = wb.Sheets[name]; if (!ws) continue; return parseDetailSheetRobust(ws, name); }
    catch (e: any) { if (e?.debug) lastDebug = e.debug; errs.push(`${name}: ${e?.message || e}`); }
  }
  const err: any = new Error("Detalle: ninguna hoja válida. Detalles: " + errs.join(" | "));
  if (lastDebug) err.debug = lastDebug; throw err;
}

// ====================== KPI Calcs =======================
function calcOffersFromPivot(model: any) {
  // Cuenta "Recuento de registros" en columnas cuya etapa contenga "proposal"
  const porComercial = model.rows.map((r: any) => {
    let offers = 0;
    for (const [stage, agg] of Object.entries(r.values)) {
      const st = norm(stage);
      if (st.includes("proposal")) {
        offers += (agg as any).count || 0;
      }
    }
    return { comercial: r.comercial, offers };
  });
  const total = porComercial.reduce((a: number, x: any) => a + x.offers, 0);
  return { total, porComercial };
}
function calcPipelineFromPivot(model: any) {
  const porComercial = model.rows.map((r: any) => {
    let sum = 0; for (const [stage, agg] of Object.entries(r.values)) { const st = norm(stage); if (OPEN_STAGES.some(s => st.includes(s))) sum += (agg as any).sum; }
    return { comercial: r.comercial, pipeline: sum };
  });
  const total = porComercial.reduce((a: number, x: any) => a + x.pipeline, 0);
  return { total, porComercial };
}
function calcWinRateFromPivot(model: any) {
  const porComercial = model.rows.map((r: any) => {
    let won = 0, lost = 0; for (const [stage, agg] of Object.entries(r.values)) { const st = norm(stage); if (st.includes("closed won") || st.includes("ganad")) won += (agg as any).count; else if (st.includes("closed lost") || st.includes("perdid")) lost += (agg as any).count; }
    const total = won + lost; const winRate = total ? (won / total) * 100 : 0;
    return { comercial: r.comercial, won, lost, total, winRate };
  });
  const tot = porComercial.reduce((a: any, c: any) => ({ comercial: "ALL", won: a.won + c.won, lost: a.lost + c.lost, total: a.total + c.total, winRate: 0 }), { comercial: "ALL", won: 0, lost: 0, total: 0, winRate: 0 });
  tot.winRate = tot.total ? (tot.won / tot.total) * 100 : 0;
  return { total: tot, porComercial };
}
function calcSalesCycleFromDetail(model: any) {
  const isClosed = (et: string) => { const s = norm(et); return s.includes("closed won") || s.includes("closed lost") || s.includes("ganad") || s.includes("perdid"); };
  const by = new Map<string, number[]>();
  model.rows.forEach((row: any) => {
    if (!isClosed(row.Etapa)) return;
    let days: number | null = row.Antiguedad ?? null;
    if (days == null) days = daysBetween(row.Fecha_creacion, row.Fecha_cierre) as any;
    if (days == null) return;
    const key = row.Comercial || "(Sin comercial)";
    if (!by.has(key)) by.set(key, []);
    (by.get(key) as number[]).push(days);
  });
  const porComercial = Array.from(by.entries()).map(([comercial, arr]) => ({ comercial, avgDays: arr.reduce((a, v) => a + v, 0) / arr.length, total: arr.length }))
    .sort((a, b) => a.comercial.localeCompare(b.comercial));
  const all = ([] as number[]).concat(...Array.from(by.values()));
  const totalAvgDays = all.length ? (all.reduce((a, v) => a + v, 0) / all.length) : 0;
  return { totalAvgDays, totalCount: all.length, porComercial };
}

// Cumplimiento de Meta (Anual) a partir del RESUMEN (pivot)
function calcAttainmentFromPivot(model: any) {
  const porComercial = model.rows.map((r: any) => {
    let wonCOP = 0;
    for (const [stage, agg] of Object.entries(r.values)) {
      const st = norm(stage);
      if (WON_STAGES.some(ws => st.includes(ws))) wonCOP += (agg as any).sum || 0;
    }
    const goal = metaAnual(r.comercial);
    const pct = goal > 0 ? (wonCOP / goal) * 100 : 0;
    return { comercial: r.comercial, wonCOP, goal, pct };
  });
  const totalWon = porComercial.reduce((a: number, x: any) => a + x.wonCOP, 0);
  const totalGoal = porComercial.reduce((a: number, x: any) => a + x.goal, 0);
  const totalPct = totalGoal > 0 ? (totalWon / totalGoal) * 100 : 0;
  return { total: { comercial: "ALL", wonCOP: totalWon, goal: totalGoal, pct: totalPct }, porComercial };
}

// ================= Fixtures (Demo) =================
const FIXTURE_PIVOT_AOA = [
  ["", "Qualification", "Qualification", "Needs Analysis", "Needs Analysis", "Proposal", "Proposal", "Closed Won", "Closed Won", "Closed Lost", "Closed Lost"],
  ["Propietario de oportunidad", "Suma de Precio total", "Recuento de registros", "Suma de Precio total", "Recuento de registros", "Suma de Precio total", "Recuento de registros", "Suma de Precio total", "Recuento de registros", "Suma de Precio total", "Recuento de registros"],
  ["Juan Garzón Linares", 10000000, 2, 25000000, 1, 0, 0, 8000000, 3, 7000000, 2],
  ["PABLO RODRIGUEZ RODRIGUEZ", 0, 0, 0, 0, 12000000, 1, 3000000, 1, 9000000, 4],
  ["TOTAL", 10000000, 2, 25000000, 1, 12000000, 1, 11000000, 4, 16000000, 6],
];
const FIXTURE_DETAIL_ROWS = [
  { "Propietario de oportunidad": "Juan Garzón Linares", "Etapa": "Closed Won", "Antigüedad": 25, "Importe": 12000000, "Probabilidad (%)": 100 },
  { "Propietario de oportunidad": "Juan Garzón Linares", "Etapa": "Closed Lost", "Antigüedad": 40, "Importe": 8000000,  "Probabilidad (%)": 0 },
  { "Propietario de oportunidad": "PABLO RODRIGUEZ RODRIGUEZ", "Etapa": "Closed Won", "Antigüedad": 30, "Importe": 6000000,  "Probabilidad (%)": 100 },
  { "Propietario de oportunidad": "PABLO RODRIGUEZ RODRIGUEZ", "Etapa": "Closed Lost", "Antigüedad": 60, "Importe": 2000000,  "Probabilidad (%)": 0 },
];

function wbFromAOA(aoa: any[][], name = "Hoja1") { const ws = XLSX.utils.aoa_to_sheet(aoa); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, name); return wb; }
function wbFromJSON(json: any[], name = "Hoja1") { const ws = XLSX.utils.json_to_sheet(json); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, name); return wb; }

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
   "KPI_OFFERS" | "KPI_VISITS"
 >("MENU");
  const [demo, setDemo] = useState(true);

  const [filePivotName, setFilePivotName] = useState("");
  const [fileDetailName, setFileDetailName] = useState("");
  const [fileVisitsName, setFileVisitsName] = useState("");
  const [pivot, setPivot] = useState<any>(null);
  const [detail, setDetail] = useState<any>(null);
  const [visits, setVisits] = useState<any>(null);
  const [offersModel, setOffersModel] = useState<any>(null);
  const [offersPeriod, setOffersPeriod] = useState<string>("");
  const [offersTarget, setOffersTarget] = useState<number>(5);
  const [selectedComercial, setSelectedComercial] = useState("ALL");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [winRateTarget, setWinRateTarget] = useState(30);
  const [cycleTarget, setCycleTarget] = useState(45);

const resetAll = () => {
  setFilePivotName("");
  setFileDetailName("");
  setFileVisitsName("");
  setPivot(null);
  setDetail(null);
  setVisits(null);
  setSelectedComercial("ALL");
  setError("");
  setInfo("");
  setWinRateTarget(30);
  setCycleTarget(45);
};
  
  const loadFixtures = () => {
    const wbPivot = wbFromAOA(FIXTURE_PIVOT_AOA, "Informe medicion KPI");
    const pv = tryParseAnyPivot(wbPivot); setPivot(pv); setFilePivotName("FIXTURE_PIVOT.xlsx");
    const wbDetail = wbFromJSON(FIXTURE_DETAIL_ROWS, "Detalle");
    const dt = tryParseAnyDetail(wbDetail); setDetail(dt); setFileDetailName("FIXTURE_DETAIL.xlsx");
    setInfo("Fixtures cargados (Resumen + Detalle)"); setError("");
  };

  React.useEffect(() => { if (demo && (!pivot || !detail)) { try { loadFixtures(); } catch {} } }, [demo]);

  const colorForWinRate = (valuePct: number) => valuePct >= winRateTarget ? "bg-green-500" : (valuePct >= winRateTarget * 0.8 ? "bg-yellow-400" : "bg-red-500");
  const colorForCycle = (days: number) => days <= cycleTarget ? "bg-green-500" : (days <= cycleTarget * 1.2 ? "bg-yellow-400" : "bg-red-500");

  async function onPivotFile(f: File) {
    setError(""); setInfo(prev => prev ? prev + "\n" : ""); setFilePivotName(f.name);
    try { const wb = await readWorkbookRobust(f); const model = tryParseAnyPivot(wb); setPivot(model); setInfo(prev => (prev + `Resumen OK • hoja: ${model.sheetName}`).trim()); }
    catch (e: any) { setPivot(null); setError(prev => (prev ? prev + "\n" : "") + `Resumen: ${e?.message || e}`); }
  }
  async function onDetailFile(f: File) {
    setError(""); setInfo(prev => prev ? prev + "\n" : ""); setFileDetailName(f.name);
    try { const wb = await readWorkbookRobust(f); const model = tryParseAnyDetail(wb); setDetail(model); 
    const off = buildOffersModelFromDetail(wb);
      setOffersModel(off);
      if (off.periods && off.periods.length) {
        setOffersPeriod(off.periods[off.periods.length - 1]); // último mes disponible
      }

    
    setInfo(prev => (prev + `Detalle OK • hoja: ${model.sheetName}`).trim()); }
    catch (e: any) { setDetail(null); let dbg = e?.debug ? "\n" + (e.debug).join("\n") : ""; setOffersModel(null); setError(prev => (prev ? prev + "\n" : "") + `Detalle: ${e?.message || e}${dbg}`); }
  }
  async function onVisitsFile(f: File) {
  setError("");
  setInfo(prev => prev ? prev + "\n" : "");
  setFileVisitsName(f.name);
  try {
    const wb = await readWorkbookRobust(f);
    const model = tryParseAnyVisits(wb);
    setVisits(model);
    setInfo(prev => (prev + `Visitas OK • hoja: ${model.sheetName}`).trim());
  } catch (e: any) {
    setVisits(null);
    setError(prev => (prev ? prev + "\n" : "") + `Visitas: ${e?.message || e}`);
  }
}

  
  const comercialesMenu = useMemo(() => FIXED_COMERCIALES, []);

  const pipeline = useMemo(() => pivot ? calcPipelineFromPivot(pivot) : { total: 0, porComercial: [] }, [pivot]);
  const winRate = useMemo(() => pivot ? calcWinRateFromPivot(pivot) : { total: { comercial: "ALL", won: 0, lost: 0, total: 0, winRate: 0 }, porComercial: [] }, [pivot]);
  const salesCycle = useMemo(() => detail ? calcSalesCycleFromDetail(detail) : { totalAvgDays: 0, totalCount: 0, porComercial: [] }, [detail]);
  const offers = useMemo(() => pivot ? calcOffersFromPivot(pivot) : { total: 0, porComercial: [] }, [pivot]);

  const visitsKPI = useMemo(() => {
    if (!visits) return { total: 0, porComercial: [] as any[] };
    const by = new Map<string, number>();
    for (const r of visits.rows) {
      const key = r.comercial || "(Sin comercial)";
      by.set(key, (by.get(key) || 0) + (Number(r.n) || 0));
    }
    const porComercial = Array.from(by.entries())
      .map(([comercial, count]) => ({ comercial, count }))
      .sort((a, b) => a.comercial.localeCompare(b.comercial));
    const total = porComercial.reduce((a, x) => a + x.count, 0);
    return { total, porComercial };
  }, [visits]);

  const BackBar = ({ title }: { title: string }) => (
    <header className="px-4 py-3 bg-white border-b sticky top-0 z-10">
      <div className="max-w-6xl mx-auto flex items-center justify-between gap-3">
        <h2 className="text-xl md:text-2xl font-bold">{title}</h2>
        <div className="flex gap-2">
          <button className="px-3 py-2 rounded border" onClick={() => setRoute("MENU")}>Volver al menú</button>
        </div>
      </div>
    </header>
  );

  const ScreenPipeline = () => {
    const data = useMemo(() => pipeline, [pipeline]);
    const selected = useMemo(() => {
      if (!pivot) return 0; if (selectedComercial === "ALL") return data.total; const row = data.porComercial.find(r => r.comercial === selectedComercial); return row ? row.pipeline : 0;
    }, [pivot, data, selectedComercial]);
    const max = useMemo(() => data.porComercial.reduce((m: number, x: any) => Math.max(m, x.pipeline), 0) || 1, [data]);
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Pipeline (COP)" />
        <main className="max-w-6xl mx-auto p-4 space-y-6">
          <section className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Comercial: {selectedComercial}</div>
            <div className="text-3xl font-bold mt-1">$ {fmtCOP(selected)}</div>
            <div className="text-xs text-gray-500 mt-1">Etapas: Qualification, Needs Analysis, Proposal, Negotiation</div>
          </section>
          {pivot && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="mb-3 font-semibold">Pipeline por comercial</div>
              <div className="space-y-2">
{data.porComercial.map((row: any) => {
  const pct = Math.round((row.pipeline / (max || 1)) * 100);
  return (
    <div key={row.comercial} className="text-sm">
      <div className="flex justify-between items-center">
        <span className="font-medium">{row.comercial}</span>
        <span>$ {fmtCOP(row.pipeline)}</span>
      </div>
      <div className="h-2 bg-gray-200 rounded">
        <div className="h-2 rounded bg-gray-700" style={{ width: pct + "%" }} />
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

  const ScreenWinRate = () => {
    const data = useMemo(() => winRate, [winRate]);
    const selected = useMemo(() => {
      if (!pivot) return 0; if (selectedComercial === "ALL") return data.total.winRate; return data.porComercial.find(r => r.comercial === selectedComercial)?.winRate || 0;
    }, [pivot, data, selectedComercial]);
    const max = useMemo(() => Math.max(data.total.winRate, ...(data.porComercial.map((r: any) => r.winRate))), [data]);
    const color = (v: number)=> v>=winRateTarget?"bg-green-500":(v>=winRateTarget*0.8?"bg-yellow-400":"bg-red-500");
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Tasa de Cierre (Win Rate)" />
        <main className="max-w-6xl mx-auto p-4 space-y-6">
          <section className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Comercial: {selectedComercial}</div>
            <div className="mt-2 flex items-end gap-3">
              <div className={`w-3 h-3 rounded-full ${color(selected)}`}></div>
              <div className="text-3xl font-bold">{Math.round(selected)}%</div>
            </div>
            <div className="text-xs text-gray-500 mt-1">Meta: {winRateTarget}% — Verde ≥ meta · Amarillo ≥ 80% · Rojo &lt; 80%</div>
          </section>
          {pivot && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="mb-3 font-semibold">Win Rate por comercial</div>
              <div className="space-y-2">
                {data.porComercial.map((row: any) => {
                  const pct = Math.round((row.winRate / (max || 1)) * 100);
                  return (
                    <div key={row.comercial} className="text-sm">
                      <div className="flex justify-between items-center">
                        <span className="font-medium">{row.comercial}</span>
                        <span className="flex items-center gap-2"><span className={`inline-block w-2 h-2 rounded-full ${color(row.winRate)}`}></span><span>{Math.round(row.winRate)}% ({row.won}/{row.total})</span></span>
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

const ScreenOffers = () => {
  const data = offersKPI;
  const selected = useMemo(() => {
    if (!offersModel) return 0;
    if (selectedComercial === "ALL") return data.total;
    const row = data.porComercial.find((r:any)=>r.comercial===selectedComercial);
    return row ? row.count : 0;
  }, [offersModel, data, selectedComercial]);

  const max = useMemo(() => data.porComercial.reduce((m:number,x:any)=>Math.max(m, x.count), 0) || 1, [data]);

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
                onChange={(e)=>setOffersPeriod(e.target.value)}
              >
                {(data.periods || []).map((p:string)=><option key={p} value={p}>{p}</option>)}
              </select>
            </div>
            <div className="text-sm text-gray-500">Meta mensual:
              <input
                type="number"
                className="ml-2 w-20 border rounded px-2 py-1 text-sm"
                value={offersTarget}
                onChange={(e)=>setOffersTarget(Math.max(0, Number(e.target.value||0)))}
              />
            </div>
          </div>

          <div className="mt-3 grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-3 bg-gray-100 rounded">
              <div className="text-xs text-gray-500">Ofertas del período</div>
              <div className="text-2xl font-bold">{data.total}</div>
            </div>
            <div className="p-3 bg-gray-100 rounded">
              <div className="text-xs text-gray-500">Del comercial seleccionado</div>
              <div className="text-2xl font-bold">{selected}</div>
            </div>
            <div className="p-3 bg-gray-100 rounded">
              <div className="text-xs text-gray-500">Meta mínima</div>
              <div className={`text-2xl font-bold ${selected >= offersTarget ? "text-green-600" : "text-red-600"}`}>
                {offersTarget}
              </div>
            </div>
          </div>
          <div className="text-xs text-gray-500 mt-2">
            Fuente: Archivo DETALLADO (una fila = una oferta). Requiere columnas: <em>Comercial</em> y <em>Fecha de oferta</em>.
          </div>
        </section>

        {offersModel && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">Ranking de ofertas por comercial ({data.period})</div>
            <div className="space-y-2">
{data.porComercial.map((row: any, i: number) => {
  const pctBar = Math.round((row.count / (max || 1)) * 100); // ancho relativo al top
  const st = offerStatus(row.count, offersTarget);
  const pctTarget = offersTarget > 0 ? ((row.count / offersTarget) * 100) : 0;
  const pctLabel = `${Math.round(pctTarget)}%`; // sin decimales

  return (
    <div key={row.comercial} className="text-sm">
      <div className="flex items-center justify-between gap-2">
        {/* IZQUIERDA: orden + nombre (sin color) */}
        <div className="font-medium">{i + 1}. {row.comercial}</div>

        {/* DERECHA: números en negro + punto de color */}
        <div className="flex items-center gap-2">
          <span className="tabular-nums text-gray-900">
            {pctLabel} ({row.count}/{offersTarget})
          </span>
          <span className={`inline-block w-2 h-2 rounded-full ${st.dot}`} />
        </div>
      </div>

      {/* Barra gris (como el resto de la app) */}
      <div className="h-2 bg-gray-200 rounded mt-1">
        <div
          className="h-2 rounded bg-gray-700"
          style={{ width: pctBar + "%" }}
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

const ScreenVisits = () => {
  const data = useMemo(() => visitsKPI, [visitsKPI]);
  const selected = useMemo(() => {
    if (!visits) return 0;
    if (selectedComercial === "ALL") return data.total;
    const row = data.porComercial.find((r: any) => r.comercial === selectedComercial);
    return row ? row.count : 0;
  }, [visits, data, selectedComercial]);
  const max = useMemo(() => data.porComercial.reduce((m: number, x: any) => Math.max(m, x.count), 0) || 1, [data]);

  return (
    <div className="min-h-screen bg-gray-50">
      <BackBar title="KPI • Visitas (conteo)" />
      <main className="max-w-6xl mx-auto p-4 space-y-6">
        <section className="p-4 bg-white rounded-xl border">
          <div className="text-sm text-gray-500">Comercial: {selectedComercial}</div>
          <div className="text-3xl font-bold mt-1">{selected} visitas</div>
          <div className="text-xs text-gray-500 mt-1">Fuente: archivo de Visitas (una fila = una visita, o columna “Visitas”).</div>
        </section>
        {visits && (
          <section className="p-4 bg-white rounded-xl border">
            <div className="mb-3 font-semibold">Visitas por comercial</div>
            <div className="space-y-2">
{data.porComercial.map((row: any, i: number) => {
  const pctBar = Math.round((row.count / (max || 1)) * 100); // ancho relativo al top
  const st = offerStatus(row.count, offersTarget);
  const pctTarget = offersTarget > 0 ? Math.round((row.count / offersTarget) * 100) : 0;

  return (
    <div key={row.comercial} className="text-sm">
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <span className={`inline-block w-2 h-2 rounded-full ${st.dot}`} />
          <span className="font-medium">{i + 1}. {row.comercial}</span>
        </div>
        <div className={`tabular-nums ${st.text}`}>
          {/* "hechas / meta (porcentaje)" */}
          {row.count} / {offersTarget} ({pctTarget}%)
        </div>
      </div>
      <div className="h-2 bg-gray-200 rounded mt-1">
        <div
          className={`h-2 rounded ${st.bar}`}
          style={{ width: pctBar + "%" }}
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

  const ScreenCycle = () => {
    const data = useMemo(() => salesCycle, [salesCycle]);
    const selected = useMemo(() => {
      if (!detail) return 0; if (selectedComercial === "ALL") return data.totalAvgDays; return data.porComercial.find(r => r.comercial === selectedComercial)?.avgDays || 0;
    }, [detail, data, selectedComercial]);
    const max = useMemo(() => Math.max(data.totalAvgDays, ...(data.porComercial.map((r: any) => r.avgDays))), [data]);
    const color = (d: number)=> d<=cycleTarget?"bg-green-500":(d<=cycleTarget*1.2?"bg-yellow-400":"bg-red-500");
    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Sales Cycle (días)" />
        <main className="max-w-6xl mx-auto p-4 space-y-6">
          <section className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Comercial: {selectedComercial}</div>
            <div className="mt-2 flex items-end gap-3"><div className={`w-3 h-3 rounded-full ${color(selected || 0)}`}></div><div className="text-3xl font-bold">{Math.round(selected || 0)} días</div></div>
            <div className="text-xs text-gray-500 mt-1">Verde ≤ meta ({cycleTarget} días) · Amarillo ≤ 120% meta · Rojo &gt; 120% meta</div>
          </section>
          {detail && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="mb-3 font-semibold">Sales Cycle por comercial</div>
              <div className="space-y-2">
                {data.porComercial.map((row: any) => {
                  const pct = Math.round(((row.avgDays || 0) / (max || 1)) * 100);
                  return (
                    <div key={row.comercial} className="text-sm">
                      <div className="flex justify-between items-center">
                        <span className="font-medium">{row.comercial}</span>
                        <span className="flex items-center gap-2"><span className={`inline-block w-2 h-2 rounded-full ${color(row.avgDays || 0)}`}></span><span>{Math.round(row.avgDays || 0)} días (n={row.total})</span></span>
                      </div>
                      <div className="h-2 bg-gray-200 rounded"><div className="h-2 rounded bg-gray-700" style={{ width: pct + "%" }} /></div>
                    </div>
                  );
                })}
              </div>
            </section>
          )}
          {detail && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="font-semibold mb-2">Debug de columnas (Detalle)</div>
              <pre className="text-xs bg-gray-50 p-2 rounded whitespace-pre-wrap">{detail.debug.join("\n")}</pre>
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


  const ScreenAttainment = () => {
    const data = useMemo(() => (pivot ? calcAttainmentFromPivot(pivot) : { total: { comercial: "ALL", wonCOP: 0, goal: 0, pct: 0 }, porComercial: [] }), [pivot]);
    const selected = useMemo(() => {
      if (!pivot) return { comercial: "ALL", wonCOP: 0, goal: 0, pct: 0 } as any;
      if (selectedComercial === "ALL") return data.total;
      const row = data.porComercial.find((r: any) => r.comercial === selectedComercial);
      return row || { comercial: selectedComercial, wonCOP: 0, goal: metaAnual(selectedComercial), pct: 0 };
    }, [pivot, data, selectedComercial]);
    const color = (p: number) => p >= 100 ? "bg-green-500" : (p >= 80 ? "bg-yellow-400" : "bg-red-500");
    const max = useMemo(() => Math.max(data.total.pct, ...(data.porComercial.map((r: any) => r.pct))), [data]);

    return (
      <div className="min-h-screen bg-gray-50">
        <BackBar title="KPI • Cumplimiento de Meta (Anual)" />
        <main className="max-w-6xl mx-auto p-4 space-y-6">
          <section className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Comercial: {selectedComercial}</div>
            <div className="mt-2 flex items-end gap-3">
              <div className={`w-3 h-3 rounded-full ${color(selected.pct)}`}></div>
              <div className="text-3xl font-bold">{Math.round(selected.pct)}%</div>
            </div>
            <div className="text-xs text-gray-500 mt-1">Meta anual = meta mensual × 12</div>
            <div className="text-xs text-gray-500">Cerrado (Closed Won): $ {fmtCOP(selected.wonCOP)} / Meta: $ {fmtCOP(selected.goal)}</div>
          </section>

          {pivot && (
            <section className="p-4 bg-white rounded-xl border">
              <div className="mb-3 font-semibold">Cumplimiento por comercial</div>
              <div className="space-y-2">
                {data.porComercial.map((row: any) => {
                  const pctW = Math.round((row.pct / (max || 1)) * 100);
                  return (
                    <div key={row.comercial} className="text-sm">
                      <div className="flex justify-between items-center">
                        <span className="font-medium">{row.comercial}</span>
                        <span className="flex items-center gap-2">
                          <span className={`inline-block w-2 h-2 rounded-full ${color(row.pct)}`}></span>
                          <span>{Math.round(row.pct)}% — $ {fmtCOP(row.wonCOP)} / $ {fmtCOP(row.goal)}</span>
                        </span>
                      </div>
                      <div className="h-2 bg-gray-200 rounded"><div className="h-2 rounded bg-gray-700" style={{ width: pctW + "%" }} /></div>
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

  // ========= Router principal: una pantalla a la vez =========
  if (route === "HOME") return <RouteHome onEnter={() => setRoute("MENU")} />;
  if (route === "KPI_PIPELINE") return <ScreenPipeline />;
  if (route === "KPI_WINRATE") return <ScreenWinRate />;
  if (route === "KPI_CYCLE") return <ScreenCycle />;
  if (route === "KPI_ATTAIN") return <ScreenAttainment />;
  if (route === "KPI_OFFERS") return <ScreenOffers />;
  if (route === "KPI_VISITS") return <ScreenVisits />;

  // ===== Menú ===== (route === "MENU")
  return (
    <div className="min-h-screen bg-gray-50">
      <header className="px-4 py-3 bg-white border-b sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-3">
          <h2 className="text-xl md:text-2xl font-bold">Menú principal</h2>
          <div className="flex items-center gap-2">
            <span className={`text-xs px-2 py-1 rounded-full border ${demo ? 'bg-green-50 border-green-300 text-green-700' : 'bg-gray-50 border-gray-300 text-gray-600'}`}>Modo Demo: {demo ? 'ON' : 'OFF'}</span>
            <button className="px-3 py-2 rounded border" onClick={() => { if (!demo) { setDemo(true); loadFixtures(); } else { setDemo(false); } }}>Toggle Demo</button>
            <button className="px-3 py-2 rounded border" onClick={resetAll}>Reiniciar</button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto p-4 space-y-6">
        {(error || info) && (
          <div className="space-y-2">
            {error && <div className="p-3 rounded border border-red-300 bg-red-50 text-sm text-red-700 whitespace-pre-wrap">{error}</div>}
            {info && <div className="p-3 rounded border border-blue-300 bg-blue-50 text-xs text-blue-700 whitespace-pre-wrap">{info}</div>}
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
          <div className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Meta Win Rate (%)</div>
            <input type="number" className="mt-2 w-24 border rounded px-2 py-1" value={winRateTarget} onChange={(e) => setWinRateTarget(Math.max(0, Math.min(100, Number(e.target.value) || 0)))} />
          </div>
          <div className="p-4 bg-white rounded-xl border">
            <div className="text-sm text-gray-500">Meta Sales Cycle (días)</div>
            <input type="number" className="mt-2 w-24 border rounded px-2 py-1" value={cycleTarget} onChange={(e) => setCycleTarget(Math.max(1, Number(e.target.value) || 1))} />
          </div>
        </section>

        {/* Cargar informes */}
        <section className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="p-4 bg-white rounded-xl border">
            <div className="font-semibold">Archivo RESUMEN (tabla dinámica)</div>
            <div className="text-xs text-gray-500 mb-2">Filas por Comercial, columnas por Etapa, métricas: Suma de Precio total / Recuento de registros</div>
            <input type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv" onChange={(e) => e.target.files && onPivotFile(e.target.files[0])} className="block text-sm" />
            <div className="text-xs text-gray-500 mt-1">{filePivotName || "Sin archivo"}</div>
            <div className="mt-3"><button className="px-3 py-2 rounded border" onClick={() => setRoute("KPI_PIPELINE") } disabled={!pivot}>Ir a Pipeline</button></div>
          </div>
          <div className="p-4 bg-white rounded-xl border">
            <div className="font-semibold">Archivo DETALLADO</div>
            <div className="text-xs text-gray-500 mb-2">Incluye: Propietario, Etapa, Antigüedad o Fechas, Importe, Probabilidad, Producto, Precio total, Cuenta</div>
            <input type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv" onChange={(e) => e.target.files && onDetailFile(e.target.files[0])} className="block text-sm" />
            <div className="text-xs text-gray-500 mt-1">{fileDetailName || "Sin archivo"}</div>
            <div className="mt-3"><button className="px-3 py-2 rounded border" onClick={() => setRoute("KPI_CYCLE") } disabled={!detail}>Ir a Sales Cycle</button></div>
          </div>
         <div className="p-4 bg-white rounded-xl border">
          <div className="font-semibold">Archivo VISITAS</div>
          <div className="text-xs text-gray-500 mb-2">
            Columnas sugeridas: <em>Comercial</em> / <em>Propietario</em>, <em>Fecha</em> (opcional), <em>Visitas</em> (opcional).
            Si no hay columna de Visitas, cada fila cuenta como 1.
          </div>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
            onChange={(e) => e.target.files && onVisitsFile(e.target.files[0])}
            className="block text-sm"
          />
          <div className="text-xs text-gray-500 mt-1">{fileVisitsName || "Sin archivo"}</div>
          <div className="mt-3">
            <button
              className="px-3 py-2 rounded border"
              onClick={() => setRoute("KPI_VISITS")}
              disabled={!pivot}
            >
              Ir a Visitas
            </button>
          </div>
        </div>
      </section>

        {/* Tarjetas de acceso a KPIs */}
        <section className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">📊 Pipeline (COP)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_PIPELINE")} disabled={!pivot}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">🎯 Tasa de Cierre (Win Rate)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_WINRATE")} disabled={!pivot}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">⏱️ Sales Cycle (días)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: DETALLADO</p>
            <button className="mt-auto px-3 py-2 rounded bg-black text-white disabled:opacity-40" onClick={() => setRoute("KPI_CYCLE")} disabled={!detail}>Ver KPI</button>
          </div>
          <div className="p-4 bg-white rounded-xl border flex flex-col">
            <div className="font-semibold">🏁 Cumplimiento de Meta (Anual)</div>
            <p className="text-xs text-gray-500 mt-1">Fuente: RESUMEN + Metas mensuales × 12</p>
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
              disabled={!pivot}
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
              disabled={!pivot}
            >
              Ver KPI
            </button>
          </div>
        </section>

        {/* Fixtures / Self‑Test */}
        <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <button className="px-3 py-2 rounded border bg-white" onClick={loadFixtures}>Cargar Fixtures</button>
          <button className="px-3 py-2 rounded border bg-white" onClick={() => {
            try {
              const wb1 = wbFromAOA(FIXTURE_PIVOT_AOA, "Informe medicion KPI");
              const arr = XLSX.write(wb1, { type: "array", bookType: "xlsx" });
              const wb2 = XLSX.read(new Uint8Array(arr), { type: "array", dense: true });
              const pv = tryParseAnyPivot(wb2);
              const p = calcPipelineFromPivot(pv);
              const expected = 10000000 + 25000000 + 12000000; if (p.total !== expected) throw new Error("Pipeline no coincide");
              const wr = calcWinRateFromPivot(pv);
              if (Math.round(wr.total.winRate) !== Math.round((4 / (4 + 6)) * 100)) throw new Error("Win Rate no coincide");
              const det = tryParseAnyDetail(wbFromJSON(FIXTURE_DETAIL_ROWS, "Detalle"));
              const sc = calcSalesCycleFromDetail(det);
              const expAvg = (25 + 40 + 30 + 60) / 4; if (Math.round(sc.totalAvgDays) !== Math.round(expAvg)) throw new Error("Sales Cycle no coincide");
              setInfo("SELF‑TEST: OK"); setError(""); setPivot(pv); setDetail(det);
            } catch (e: any) { setError("SELF‑TEST falló: " + (e?.message || e)); }
          }}>Self‑Test</button>
          <button className="px-3 py-2 rounded border bg-white" onClick={resetAll}>Reiniciar</button>
        </section>
      </main>
    </div>
  );
}
