let companyRows = [];
let sectorMap = {};
let amountByCodeYear = new Map();
let availableYears = [];
let chartInstance = null;
let cashChartInstance = null;

const REQUIRED_COLS = ["concept_code","period","statement","amount","node_type","agg_rule"];
const el = (id) => document.getElementById(id);

const fmtEUR = (n) => {
  if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
  return new Intl.NumberFormat("es-ES", { style:"currency", currency:"EUR", maximumFractionDigits:0 }).format(n);
};
const fmtPct = (n) => {
  if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
  return new Intl.NumberFormat("es-ES", { style:"percent", minimumFractionDigits:1, maximumFractionDigits:1 }).format(n);
};
const fmtNum = (n, digits=0) => {
  if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
  return new Intl.NumberFormat("es-ES", { maximumFractionDigits:digits, minimumFractionDigits:digits }).format(n);
};

function setStatus(msg, isError=false){
  el("status").textContent = msg;
  el("status").style.color = isError ? "#ffb4b4" : "var(--muted)";
}

function parseCsvFile(file, callback){
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    dynamicTyping: false,
    complete: (res) => callback(null, res.data),
    error: (err) => callback(err, null)
  });
}

function validateCompanyColumns(rows){
  const cols = rows.length ? Object.keys(rows[0]) : [];
  return REQUIRED_COLS.filter(c => !cols.includes(c));
}

function toNumber(v){
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  const s = String(v).trim();
  if (!s) return null;
  const normalized =
    (s.includes(",") && s.includes(".") && s.lastIndexOf(",") > s.lastIndexOf("."))
      ? s.replaceAll(".","").replace(",",".")
      : s.includes(",") ? s.replaceAll(".","").replace(",",".") : s;
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function isAnnualPeriod(p){ return /^\d{4}$/.test(String(p).trim()); }

function buildIndex(rows){
  amountByCodeYear = new Map();
  const years = new Set();

  for (const r of rows){
    const code = String(r.concept_code ?? "").trim();
    const period = String(r.period ?? "").trim();
    if (!isAnnualPeriod(period)) continue;

    const amt = toNumber(r.amount);
    if (amt === null) continue;

    years.add(period);
    const key = `${code}__${period}`;
    amountByCodeYear.set(key, (amountByCodeYear.get(key) ?? 0) + amt);
  }

  availableYears = Array.from(years).sort();
}

function getAmount(code, year){ return amountByCodeYear.get(`${code}__${String(year).trim()}`) ?? 0; }
function safeDiv(a,b){ if (b === 0) return null; return a / b; }

function deltaInfo(base, comp){
  if (base === null || comp === null) return {deltaPct:null, deltaAbs:null};
  const deltaAbs = base - comp;
  const deltaPct = comp === 0 ? null : (deltaAbs / Math.abs(comp));
  return {deltaPct, deltaAbs};
}

function sectorPill(kpiId, value, direction){
  const s = sectorMap[kpiId];
  if (!s || value === null || value === undefined || Number.isNaN(value)) return {cls:"", text:"Sector: N/D"};
  const min = s.min, med = s.media, max = s.max;
  if ([min,med,max].some(v => v === null || v === undefined || Number.isNaN(v))) return {cls:"", text:"Sector: N/D"};

  if (direction === "HIGH_BETTER"){
    if (value >= med) return {cls:"good", text:"Sector: por encima"};
    if (value >= min) return {cls:"warn", text:"Sector: en rango"};
    return {cls:"bad", text:"Sector: por debajo"};
  } else {
    if (value <= med) return {cls:"good", text:"Sector: por encima"};
    if (value <= max) return {cls:"warn", text:"Sector: en rango"};
    return {cls:"bad", text:"Sector: por debajo"};
  }
}

function formatValue(unit, v){
  if (v === null || v === undefined || Number.isNaN(v)) return "N/D";
  if (unit === "EUR") return fmtEUR(v);
  if (unit === "PCT") return fmtPct(v);
  if (unit === "DAYS") return `${fmtNum(v, 0)} días`;
  if (unit === "X") return `${fmtNum(v, 1)}x`;
  return fmtNum(v, 2);
}

function explainKpi(kpiId, baseVal){
  if (baseVal === null || baseVal === undefined || Number.isNaN(baseVal)) return "Dato no disponible.";
  if (kpiId === "CASH_NET") return (baseVal >= 0) ? "Entró caja este año." : "Salió caja este año.";
  if (kpiId === "EBITDA_M"){
    if (baseVal < 0) return "Estás perdiendo dinero por cada 100€ vendidos.";
    if (baseVal < 0.05) return "Margen muy ajustado: cualquier susto te afecta.";
    if (baseVal < 0.12) return "Margen normal: vigila costes y precios.";
    return "Buen margen: tienes colchón para invertir y absorber imprevistos.";
  }
  if (kpiId === "FM"){
    if (baseVal < 0) return "Colchón negativo: puedes tener tensiones si algo se retrasa.";
    return "Tienes colchón para absorber el día a día.";
  }
  if (kpiId === "DSO"){
    if (baseVal > 90) return "Cobras muy tarde: revisa condiciones y seguimiento de cobros.";
    if (baseVal > 60) return "Cobro algo lento: mejora el proceso de cobro.";
    return "Cobras razonablemente rápido.";
  }
  if (kpiId === "OXY"){
    if (baseVal < 15) return "Oxígeno bajo: cualquier parada te pone contra las cuerdas.";
    if (baseVal < 45) return "Oxígeno medio: vigila caja y costes fijos.";
    return "Oxígeno cómodo: buen margen de maniobra.";
  }
  if (kpiId === "ND_EB"){
    if (baseVal > 5) return "Deuda alta para tu capacidad de pago.";
    if (baseVal > 3) return "Deuda moderada: controla crecimiento de deuda.";
    return "Deuda razonable.";
  }
  return "Mira la tendencia: subir es bueno si refuerza tu margen y tu caja.";
}

// =========================
// KPI DEFINITIONS
// =========================
const KPI_DEFS = [
  { id:"SALES", cat:"VISTA GENERAL", name:"Ventas", unit:"EUR", direction:"HIGH_BETTER",
    help:"El tamaño del negocio: cuánto vendes.",
    calc: (y) => getAmount("PYG.MAIN.1", y)
  },
  { id:"EBITDA", cat:"VISTA GENERAL", name:"EBITDA", unit:"EUR", direction:"HIGH_BETTER",
    help:"Lo que genera la operativa antes de deuda e impuestos.",
    calc: (y) => getAmount("PYG.MAIN.A.1", y) - getAmount("PYG.MAIN.8", y)
  },
  { id:"EBITDA_M", cat:"VISTA GENERAL", name:"EBITDA / Ventas", unit:"PCT", direction:"HIGH_BETTER",
    help:"De cada 100€ que vendes, cuánto te queda.",
    calc: (y) => safeDiv(KPI_DEFS.find(k=>k.id==="EBITDA").calc(y), getAmount("PYG.MAIN.1", y))
  },
  { id:"CASH_NET", cat:"VISTA GENERAL", name:"Caja neta del período", unit:"EUR", direction:"HIGH_BETTER",
    help:"Si entra caja (+) o sale caja (-) en el año.",
    calc: (y) => getAmount("EFE.MAIN.E", y)
  },

  { id:"GM_M", cat:"OPERATIVO", name:"Margen bruto / Ventas", unit:"PCT", direction:"HIGH_BETTER",
    help:"Lo que queda tras el coste directo, antes de estructura.",
    calc: (y) => {
      const sales = getAmount("PYG.MAIN.1", y);
      const mg = getAmount("PYG.MAIN.1", y)+getAmount("PYG.MAIN.2", y)+getAmount("PYG.MAIN.3", y)+getAmount("PYG.MAIN.4", y);
      return safeDiv(mg, sales);
    }
  },
  { id:"FIXED", cat:"OPERATIVO", name:"Costes fijos", unit:"EUR", direction:"LOW_BETTER",
    help:"Estructura: lo que cuesta el negocio aunque no entre nadie.",
    calc: (y) => - (getAmount("PYG.MAIN.6", y) + getAmount("PYG.MAIN.7", y))
  },
  { id:"BEP", cat:"OPERATIVO", name:"Punto de equilibrio", unit:"EUR", direction:"LOW_BETTER",
    help:"Lo mínimo a vender para no perder dinero.",
    calc: (y) => {
      const mgPct = KPI_DEFS.find(k=>k.id==="GM_M").calc(y);
      const fixed = KPI_DEFS.find(k=>k.id==="FIXED").calc(y);
      if (mgPct === null || mgPct <= 0) return null;
      return fixed / mgPct;
    }
  },

  { id:"FM", cat:"LIQUIDEZ A CORTO", name:"Colchón a corto plazo (FM)", unit:"EUR", direction:"HIGH_BETTER",
    help:"Colchón del día a día: activo corriente menos pasivo corriente.",
    calc: (y) => getAmount("BAL.ACT.B", y) - getAmount("BAL.PNP.C", y)
  },
  { id:"DSO", cat:"LIQUIDEZ A CORTO", name:"Días en cobrar", unit:"DAYS", direction:"LOW_BETTER",
    help:"Cuánto tarda el dinero en volver desde clientes.",
    calc: (y) => {
      const sales = getAmount("PYG.MAIN.1", y);
      if (sales <= 0) return null;
      const cli = getAmount("BAL.ACT.B.III.1", y) + getAmount("BAL.ACT.B.III.2", y);
      const yPrev = String(Number(y) - 1);
      const cliPrev = availableYears.includes(yPrev)
        ? (getAmount("BAL.ACT.B.III.1", yPrev) + getAmount("BAL.ACT.B.III.2", yPrev))
        : null;
      const cliAvg = (cliPrev === null) ? cli : (cli + cliPrev) / 2;
      return (cliAvg / sales) * 365;
    }
  },
  { id:"OXY", cat:"LIQUIDEZ A CORTO", name:"Días de oxígeno", unit:"DAYS", direction:"HIGH_BETTER",
    help:"Cuántos días aguantas si mañana no entra nada.",
    calc: (y) => {
      const cash = getAmount("BAL.ACT.B.VII", y);
      const fixed = KPI_DEFS.find(k=>k.id==="FIXED").calc(y);
      const interest = - getAmount("PYG.MAIN.13", y);
      const burn = fixed + (Number.isFinite(interest) ? interest : 0);
      if (burn <= 0) return null;
      return cash / (burn / 365);
    }
  },

  { id:"ND_EB", cat:"NIVEL DE ENDEUDAMIENTO", name:"Deuda neta / EBITDA", unit:"X", direction:"LOW_BETTER",
    help:"Años para pagar deuda (si el EBITDA se mantuviera).",
    calc: (y) => {
      const debt = getAmount("BAL.PNP.B.II.2", y)+getAmount("BAL.PNP.C.III.2", y)+getAmount("BAL.PNP.B.II.1", y)+getAmount("BAL.PNP.C.III.1", y);
      const cash = getAmount("BAL.ACT.B.VII", y);
      const netDebt = debt - cash;
      const ebitda = KPI_DEFS.find(k=>k.id==="EBITDA").calc(y);
      if (ebitda <= 0) return null;
      return netDebt / ebitda;
    }
  },
  { id:"DEBT_GROSS", cat:"NIVEL DE ENDEUDAMIENTO", name:"Deuda bruta (bancaria)", unit:"EUR", direction:"LOW_BETTER",
    help:"Deuda bancaria total (corto + largo).",
    calc: (y) => getAmount("BAL.PNP.B.II.2", y)+getAmount("BAL.PNP.C.III.2", y)+getAmount("BAL.PNP.B.II.1", y)+getAmount("BAL.PNP.C.III.1", y)
  },
];

// =========================
// CASHFLOW helpers
// =========================
function inOutFromNet(net){
  return { in: net > 0 ? net : 0, out: net < 0 ? -net : 0, net };
}
function forceOut(net){
  const v = Math.abs(net);
  return { in: 0, out: v, net: -v };
}
function forceIn(net){
  const v = Math.abs(net);
  return { in: v, out: 0, net: v };
}

function efeLeafRowsForYear(year){
  const y = String(year).trim();
  return companyRows.filter(r =>
    String(r.statement).trim() === "EFE" &&
    String(r.period).trim() === y &&
    ["DETAIL","TOTAL"].includes(String(r.node_type).trim()) &&
    toNumber(r.amount) !== null
  );
}

function topLines(rows, prefix, n=6){
  const list = rows
    .filter(r => String(r.concept_code).trim().startsWith(prefix))
    .map(r => ({ name: (r.display_name||r.concept_code).toString().trim(), amt: toNumber(r.amount) }))
    .filter(x => x.amt !== null);
  list.sort((a,b)=> Math.abs(b.amt) - Math.abs(a.amt));
  return list.slice(0,n);
}

function wcChartLabel(net){
  if (net < 0) return "Sale por invertir en corriente";
  if (net > 0) return "Entra por desinvertir en corriente";
  return "Corriente (neutro)";
}
function actChartLabel(net){
  if (net < 0) return "Sale por mi negocio";
  if (net > 0) return "Entra por mi negocio";
  return "Mi negocio (neutro)";
}

function buildCashflowForYear(year){
  const rowsLeaf = efeLeafRowsForYear(year);
  const cashNet = getAmount("EFE.MAIN.E", year);

  // BLOQUE 1 – ACTIVIDAD (1 + 2)
  const resultado = getAmount("EFE.MAIN.1", year);
  const ajustes = getAmount("EFE.MAIN.2", year); // usamos el subtotal (incluye TODO)
  const actNet = resultado + ajustes;

  const b1 = {
    id:"ACT",
    name:"Por mi negocio (Actividad)",
    chartLabel: actChartLabel(actNet),
    ...inOutFromNet(actNet),
    detail: [
      (() => { const x=inOutFromNet(resultado); return {label:"Resultado (antes de impuestos)", in:x.in, out:x.out}; })(),
      (() => { const x=inOutFromNet(ajustes); return {label:"Ajustes (amortizaciones, provisiones, etc.)", in:x.in, out:x.out}; })()
    ]
  };

  // BLOQUE 2 – CORRIENTE (sub-total 3, incluye 3.e aunque sea TOTAL)
  const wcNet = getAmount("EFE.MAIN.3", year);
  const b2name = (wcNet < 0) ? "Inversión corriente" : (wcNet > 0) ? "Desinversión corriente" : "Corriente";

  const b2 = {
    id:"WC",
    name:b2name,
    chartLabel: wcChartLabel(wcNet),
    ...inOutFromNet(wcNet),
    detail: [
      (()=>{ const n=getAmount("EFE.MAIN.3.a", year); const x=inOutFromNet(n); return {label:"Existencias (stock)", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.3.b", year); const x=inOutFromNet(n); return {label:"Clientes (cobras / se retrasa)", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.3.c", year); const x=inOutFromNet(n); return {label:"Otros activos corrientes", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.3.d", year); const x=inOutFromNet(n); return {label:"Proveedores (te financian / pagas)", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.3.e", year); const x=inOutFromNet(n); return {label:"Otros pasivos corrientes", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.3.f", year); const x=inOutFromNet(n); return {label:"Otros activos/pasivos no corrientes", in:x.in, out:x.out}; })(),
    ]
  };

  // BLOQUE 3 – INTERESES E IMPUESTOS (sub-total 4)
  const otherOps = getAmount("EFE.MAIN.4", year);
  const b3 = {
    id:"TAXINT",
    name:"Intereses e impuestos",
    chartLabel: otherOps < 0 ? "Sale por intereses e impuestos" : "Entra por intereses e impuestos",
    ...inOutFromNet(otherOps),
    detail: [
      (()=>{ const n=getAmount("EFE.MAIN.4.a", year); const x=inOutFromNet(n); return {label:"Pagos de intereses", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.4.b", year); const x=inOutFromNet(n); return {label:"Cobros de dividendos", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.4.c", year); const x=inOutFromNet(n); return {label:"Cobros de intereses", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.4.d", year); const x=inOutFromNet(n); return {label:"Impuestos", in:x.in, out:x.out}; })(),
      (()=>{ const n=getAmount("EFE.MAIN.4.e", year); const x=inOutFromNet(n); return {label:"Otros pagos/cobros", in:x.in, out:x.out}; })(),
    ]
  };

  // BLOQUE 4 – INVERSIONES (sub-total 6, siempre sale en términos “de dueño”)
  const invNet = getAmount("EFE.MAIN.6", year);
  const b4 = {
    id:"INV",
    name:"Inversiones",
    chartLabel:"Sale por inversiones",
    ...forceOut(invNet),
    top: topLines(rowsLeaf, "EFE.MAIN.6", 6)
  };

  // BLOQUE 5 – DESINVERSIONES (sub-total 7, siempre entra)
  const divNet = getAmount("EFE.MAIN.7", year);
  const b5 = {
    id:"DIV",
    name:"Desinversiones",
    chartLabel:"Entra por desinversiones",
    ...forceIn(divNet),
    top: topLines(rowsLeaf, "EFE.MAIN.7", 6)
  };

  // BLOQUE 6 – FINANCIACIÓN (ajustada a tu CSV)
  const equityNet = getAmount("EFE.MAIN.9", year);
  const debtInAmt = getAmount("EFE.MAIN.10.a", year);          // a) Emisión
  const debtOutAmt = getAmount("EFE.MAIN.10.a.4.b", year);     // b) Devolución y amortización
  const divAmt = getAmount("EFE.MAIN.10.a.4.b.11", year);      // 11) Pagos por dividendos / remuneraciones

  const equity = inOutFromNet(equityNet);
  const debtIn = forceIn(debtInAmt);
  const debtOut = forceOut(debtOutAmt);
  const divi = forceOut(divAmt);

  const finIn = equity.in + debtIn.in;
  const finOut = equity.out + debtOut.out + divi.out;

  const b6 = {
    id:"FIN",
    name:"Financiación",
    chartLabel:"Financiación",
    in: finIn,
    out: finOut,
    net: finIn - finOut,
    finDetail: [
      {label:"Entra por financiación (préstamos / emisiones)", in: debtIn.in, out: 0},
      {label:"Sale por devolución (amortización / cancelación)", in: 0, out: debtOut.out},
      {label:"Sale por dividendos / remuneraciones", in: 0, out: divi.out},
      {label:"Patrimonio (ampliaciones/otros)", in: equity.in, out: equity.out},
    ]
  };

  const blocks = [b1,b2,b3,b4,b5,b6];

  // Ajuste para cuadrar visualmente columnas (NO cambia el neto real)
  const totalIn = blocks.reduce((s,b)=> s + b.in, 0);
  const totalOut = blocks.reduce((s,b)=> s + b.out, 0);
  const diff = totalIn - totalOut;

  let balanceRow = null;
  if (Math.abs(diff) > 0.01){
    if (diff > 0){
      balanceRow = { id:"BAL", name:"Caja que se queda (para cuadrar)", chartLabel:"Caja que se queda",
        in:0, out:diff, net:-diff,
        note:"Es ahorro de caja. Se muestra en SALE para cuadrar columnas." };
    } else {
      balanceRow = { id:"BAL", name:"Caja que uso (para cuadrar)", chartLabel:"Caja que uso",
        in:(-diff), out:0, net:(-diff),
        note:"Has consumido caja. Se muestra en ENTRA para cuadrar columnas." };
    }
  }

  return { year, cashNet, blocks, balanceRow };
}

// =========================
// CONSOLIDACIÓN (QA) – ajustada a TU CSV
// =========================
const CRITICAL_SPECS = [
  // PyG
  {type:"exact", code:"PYG.MAIN.1", label:"Ventas"},
  {type:"exact", code:"PYG.MAIN.A.1", label:"Resultado operativo base"},
  {type:"exact", code:"PYG.MAIN.8", label:"Amortización/ajustes EBITDA"},
  {type:"exact", code:"PYG.MAIN.2", label:"Partida margen bruto 2"},
  {type:"exact", code:"PYG.MAIN.3", label:"Partida margen bruto 3"},
  {type:"exact", code:"PYG.MAIN.4", label:"Partida margen bruto 4"},
  {type:"exact", code:"PYG.MAIN.6", label:"Costes fijos 6"},
  {type:"exact", code:"PYG.MAIN.7", label:"Costes fijos 7"},
  {type:"exact", code:"PYG.MAIN.13", label:"Intereses (PyG)"},

  // Balance
  {type:"exact", code:"BAL.ACT.B", label:"Activo corriente"},
  {type:"exact", code:"BAL.PNP.C", label:"Pasivo corriente"},
  {type:"exact", code:"BAL.ACT.B.VII", label:"Caja (efectivo y equivalentes)"},
  {type:"exact", code:"BAL.ACT.B.III.1", label:"Clientes (1)"},
  {type:"exact", code:"BAL.ACT.B.III.2", label:"Clientes (2)"},
  {type:"exact", code:"BAL.PNP.B.II.1", label:"Deuda LP (1)"},
  {type:"exact", code:"BAL.PNP.B.II.2", label:"Deuda LP (2)"},
  {type:"exact", code:"BAL.PNP.C.III.1", label:"Deuda CP (1)"},
  {type:"exact", code:"BAL.PNP.C.III.2", label:"Deuda CP (2)"},

  // EFE
  {type:"exact", code:"EFE.MAIN.E", label:"Caja neta del período"},
  {type:"exact", code:"EFE.MAIN.1", label:"Resultado (EFE)"},
  {type:"exact", code:"EFE.MAIN.2", label:"Ajustes al resultado (EFE)"},
  {type:"exact", code:"EFE.MAIN.3", label:"Cambios en capital corriente (EFE)"},
  {type:"exact", code:"EFE.MAIN.4", label:"Intereses/impuestos/otros explotación (EFE)"},
  {type:"exact", code:"EFE.MAIN.6", label:"Pagos por inversiones (EFE)"},
  {type:"exact", code:"EFE.MAIN.7", label:"Cobros por desinversiones (EFE)"},
  {type:"exact", code:"EFE.MAIN.9", label:"Patrimonio (EFE)"},
  {type:"exact", code:"EFE.MAIN.10.a", label:"Emisión deuda (EFE)"},
  {type:"exact", code:"EFE.MAIN.10.a.4.b", label:"Devolución/amortización deuda (EFE)"},
  {type:"exact", code:"EFE.MAIN.10.a.4.b.11", label:"Dividendos/remuneraciones (EFE)"},
];

function hasCodeOrPrefix(spec){
  const codes = companyRows.map(r => String(r.concept_code||"").trim());
  if (spec.type === "exact") return codes.includes(spec.code);
  return codes.some(c => c.startsWith(spec.code));
}

function pillHTML(type, text){
  const cls = type === "OK" ? "good" : (type === "WARN" ? "warn" : "bad");
  return `<span class="pill ${cls}">${text}</span>`;
}

function updateQA(){
  const body = el("qaBody");
  const summary = el("qaSummary");

  if (!companyRows.length || !availableYears.length){
    body.innerHTML = `<div class="qaEmpty">Aún no hay datos. Sube el CSV de empresa.</div>`;
    summary.textContent = "—";
    return;
  }

  const baseY = el("baseYear").value || availableYears[availableYears.length-1];

  // 1) Años
  const yearsOK = availableYears.length >= 2;
  const yearsRow = {
    title:"Años detectados",
    desc: yearsOK ? `OK: ${availableYears.join(", ")}` : "Necesitas mínimo 2 años (YYYY) para comparar.",
    status: yearsOK ? pillHTML("OK","OK") : pillHTML("BAD","FALTA")
  };

  // 2) Códigos críticos
  const missing = CRITICAL_SPECS.filter(s => !hasCodeOrPrefix(s));
  const codesOK = missing.length === 0;
  const codesRow = {
    title:"Códigos críticos",
    desc: codesOK ? "OK: están las partidas clave para KPIs + EFE." : `Faltan ${missing.length} partidas necesarias.`,
    status: codesOK ? pillHTML("OK","OK") : pillHTML("BAD","FALTA")
  };

  // 3) Cuadre EFE (bloques vs EFE.MAIN.E)
  const cf = buildCashflowForYear(baseY);
  const netBlocks = cf.blocks.reduce((s,b)=> s + (b.net ?? (b.in - b.out)), 0);
  const cashNet = cf.cashNet;
  const diff = netBlocks - cashNet;
  const tol = 1; // 1€ tolerancia

  let cashStatus = "OK";
  let cashDesc = `OK: Σ(bloques) = ${fmtEUR(netBlocks)} y Caja neta (EFE.MAIN.E) = ${fmtEUR(cashNet)}.`;
  if (Math.abs(diff) > tol){
    cashStatus = "WARN";
    cashDesc = `NO CUADRA: Σ(bloques) = ${fmtEUR(netBlocks)} pero EFE.MAIN.E = ${fmtEUR(cashNet)} (diferencia ${fmtEUR(diff)}).`;
  }
  const cashRow = {
    title:"Cuadre de caja (EFE)",
    desc: cashDesc,
    status: cashStatus === "OK" ? pillHTML("OK","OK") : pillHTML("WARN","REVISAR")
  };

  // 4) Sector
  const sectorOn = el("sectorToggle").checked;
  const sectorLoaded = Object.keys(sectorMap).length > 0;
  const sectorRow = {
    title:"Sector",
    desc: sectorLoaded ? (sectorOn ? "Comparativa sector ON." : "Sector cargado pero OFF (puedes activarlo).")
                      : "No hay CSV sector cargado (opcional).",
    status: sectorLoaded ? pillHTML("OK", sectorOn ? "ON" : "OFF") : pillHTML("WARN","N/D")
  };

  const okCount = [yearsRow, codesRow, cashRow].filter(r => r.status.includes("OK")).length;
  summary.textContent = `Base ${baseY} · Checks OK: ${okCount}/3 · Filas: ${fmtNum(companyRows.length,0)}`;

  body.innerHTML = `
    <div class="qaRow">
      <div class="t">${yearsRow.title}</div>
      <div class="d">${yearsRow.desc}</div>
      <div class="p">${yearsRow.status}</div>
    </div>
    <div class="qaRow">
      <div class="t">${codesRow.title}</div>
      <div class="d">${codesRow.desc}</div>
      <div class="p">${codesRow.status}</div>
    </div>
    <div class="qaRow">
      <div class="t">${cashRow.title}</div>
      <div class="d">${cashRow.desc}</div>
      <div class="p">${cashRow.status}</div>
    </div>
    <div class="qaRow">
      <div class="t">${sectorRow.title}</div>
      <div class="d">${sectorRow.desc}</div>
      <div class="p">${sectorRow.status}</div>
    </div>
    ${!codesOK ? `
      <div class="qaCodes">
        <b>Partidas que faltan:</b><br/>
        ${missing.map(m => `• ${m.label} <span style="opacity:.8">(${m.type==="exact"?"=":"prefijo "} ${m.code})</span>`).join("<br/>")}
      </div>
    ` : ``}
  `;
}

// =========================
// Render KPIs + CASHFLOW
// =========================
function render(){
  renderKpis();
  renderCashflow();
  updateQA();
}

function renderKpis(){
  const baseY = el("baseYear").value;
  const compY = el("compYear").value;
  const sectorOn = el("sectorToggle").checked;

  const categories = ["VISTA GENERAL","OPERATIVO","LIQUIDEZ A CORTO","NIVEL DE ENDEUDAMIENTO"];
  const root = el("kpis");
  root.innerHTML = "";

  for (const cat of categories){
    const section = document.createElement("div");
    section.innerHTML = `<div class="sectionTitle">${cat}</div>`;
    const grid = document.createElement("div");
    grid.className = "kpiGrid";

    const kpis = KPI_DEFS.filter(k => k.cat === cat);
    for (const k of kpis){
      const baseVal = k.calc(baseY);
      const compVal = k.calc(compY);
      const {deltaPct, deltaAbs} = deltaInfo(baseVal, compVal);
      const pill = sectorOn ? sectorPill(k.id, baseVal, k.direction) : {cls:"", text:"Sector: OFF"};

      const card = document.createElement("div");
      card.className = "kpiCard";
      card.innerHTML = `
        <div class="kpiTop">
          <div>
            <div class="kpiName">${k.name}</div>
            <div class="kpiHelp">${k.help}</div>
          </div>
        </div>
        <div class="kpiValue">${formatValue(k.unit, baseVal)}</div>
        <div class="kpiSub">${compY}: ${formatValue(k.unit, compVal)}</div>
        <div class="kpiMeta">
          <span class="pill ${pill.cls}">${pill.text}</span>
          <span class="delta">
            Variación:
            <b>${deltaPct === null ? "N/D" : fmtPct(deltaPct)}</b>
            ${deltaAbs === null ? "" : ` (${k.unit==="EUR" ? fmtEUR(deltaAbs) : fmtNum(deltaAbs,2)})`}
          </span>
        </div>
      `;
      card.addEventListener("click", () => openKpiModal(k, baseY, compY, sectorOn));
      grid.appendChild(card);
    }

    section.appendChild(grid);
    root.appendChild(section);
  }
}

// =========================
// Modal KPI
// =========================
function openKpiModal(kpi, baseY, compY, sectorOn){
  const baseVal = kpi.calc(baseY);
  el("modalTitle").textContent = kpi.name;
  el("modalSub").textContent = explainKpi(kpi.id, baseVal);

  const years = availableYears.slice(-6);
  const series = years.map(y => kpi.calc(y));

  const s = sectorMap[kpi.id];
  const hasSector = sectorOn && s && [s.min,s.media,s.max].every(v => v !== null && v !== undefined && !Number.isNaN(v));
  const pill = sectorOn ? sectorPill(kpi.id, baseVal, kpi.direction) : {cls:"", text:"Sector: OFF"};

  el("modalTable").innerHTML = `
    <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:10px; margin-top:8px;">
      <div class="card" style="padding:10px;">
        <div style="color:var(--muted); font-size:12px;">${baseY}</div>
        <div style="font-weight:900; font-size:16px; margin-top:4px;">${formatValue(kpi.unit, baseVal)}</div>
      </div>
      <div class="card" style="padding:10px;">
        <div style="color:var(--muted); font-size:12px;">${compY}</div>
        <div style="font-weight:900; font-size:16px; margin-top:4px;">${formatValue(kpi.unit, kpi.calc(compY))}</div>
      </div>
      <div class="card" style="padding:10px;">
        <div style="color:var(--muted); font-size:12px;">Sector</div>
        <div style="margin-top:8px;"><span class="pill ${pill.cls}">${pill.text}</span></div>
      </div>
    </div>
  `;

  if (chartInstance){ chartInstance.destroy(); chartInstance = null; }
  const ctx = el("modalChart").getContext("2d");
  const datasets = [];

  if (hasSector){
    datasets.push({ label:"Sector (mín)", data: years.map(()=>s.min), borderWidth:0, pointRadius:0, tension:.2 });
    datasets.push({ label:"Sector (máx)", data: years.map(()=>s.max), borderWidth:0, pointRadius:0, tension:.2, fill:"-1" });
    datasets.push({ label:"Sector (media)", data: years.map(()=>s.media), borderDash:[4,4], pointRadius:0, tension:.2 });
  }
  datasets.push({ label:kpi.name, data: series, borderWidth:2, pointRadius:3, tension:.25 });

  chartInstance = new Chart(ctx, {
    type:"line",
    data:{ labels: years, datasets },
    options:{
      responsive:true,
      plugins:{
        legend:{ display:true, labels:{ color:"#cbd5ff" } },
        tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${formatValue(kpi.unit, c.parsed.y)}` } }
      },
      scales:{
        x:{ ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } }
      }
    }
  });

  el("modal").classList.remove("hidden");
}

el("modalClose").addEventListener("click", ()=> el("modal").classList.add("hidden"));
el("modal").addEventListener("click", (e)=> { if(e.target.id==="modal") el("modal").classList.add("hidden"); });

// =========================
// CASHFLOW render
// =========================
function renderCashflow(){
  if (!availableYears.length){
    el("cashSub").textContent = "Sube el CSV para ver este bloque.";
    el("cashCompare").disabled = true;
    return;
  }

  const baseY = el("baseYear").value;
  const compY = el("compYear").value;

  el("cashCompare").disabled = false;
  const compare = el("cashCompare").checked;

  const base = buildCashflowForYear(baseY);
  const comp = compare ? buildCashflowForYear(compY) : null;

  if (!compare){
    el("cashSub").textContent =
      `Año ${baseY}. Caja neta del período: ${fmtEUR(base.cashNet)} → ${base.cashNet >= 0 ? "Entró caja" : "Salió caja"}.`;
  } else {
    el("cashSub").textContent =
      `Comparativa ${baseY} vs ${compY}. Caja neta: ${fmtEUR(base.cashNet)} vs ${fmtEUR(comp.cashNet)}.`;
  }

  const labels = compare
    ? [`Entra ${baseY}`, `Sale ${baseY}`, `Entra ${compY}`, `Sale ${compY}`]
    : ["Entra dinero", "Sale dinero"];

  const rowsBase = base.balanceRow ? [...base.blocks, base.balanceRow] : [...base.blocks];
  const rowsComp = comp ? (comp.balanceRow ? [...comp.blocks, comp.balanceRow] : [...comp.blocks]) : [];

  const byId = new Map();
  for (const r of rowsBase) byId.set(r.id, {base:r});
  for (const r of rowsComp){
    const cur = byId.get(r.id) || {};
    cur.comp = r;
    byId.set(r.id, cur);
  }

  const datasets = [];
  for (const [id, pack] of byId.entries()){
    const b = pack.base;
    const c = pack.comp;
    const label = compare ? (b ? b.chartLabel : c.chartLabel) : (b ? b.chartLabel : id);
    const data = compare
      ? [(b?.in || 0), (b?.out || 0), (c?.in || 0), (c?.out || 0)]
      : [(b?.in || 0), (b?.out || 0)];
    datasets.push({ label, data, borderWidth: 1 });
  }

  if (cashChartInstance){ cashChartInstance.destroy(); cashChartInstance = null; }
  const ctx = el("cashChart").getContext("2d");
  cashChartInstance = new Chart(ctx, {
    type:"bar",
    data:{ labels, datasets },
    options:{
      responsive:true,
      plugins:{
        legend:{ display:true, labels:{ color:"#cbd5ff" } },
        tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${fmtEUR(c.parsed.y)}` } }
      },
      scales:{
        x:{ stacked:true, ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ stacked:true, ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } }
      }
    }
  });

  const makeNum = (v) => `<td class="num">${fmtEUR(v)}</td>`;
  let html = `
    <table>
      <thead>
        <tr>
          <th>Concepto</th>
          <th class="num">Entra ${baseY}</th>
          <th class="num">Sale ${baseY}</th>
          <th class="num">Neto ${baseY}</th>
          ${compare ? `<th class="num">Entra ${compY}</th><th class="num">Sale ${compY}</th><th class="num">Neto ${compY}</th>` : ``}
        </tr>
      </thead>
      <tbody>
  `;

  const baseTableRows = base.blocks.concat(base.balanceRow ? [base.balanceRow] : []);
  const compById = new Map((comp ? comp.blocks.concat(comp.balanceRow ? [comp.balanceRow] : []) : []).map(x => [x.id, x]));

  for (const b of baseTableRows){
    const net = b.in - b.out;
    const c = compare ? compById.get(b.id) : null;
    const cNet = c ? (c.in - c.out) : null;

    html += `
      <tr class="cashRow" data-id="${b.id}">
        <td><b>${b.name}</b></td>
        ${makeNum(b.in)}${makeNum(b.out)}${makeNum(net)}
        ${compare ? `${makeNum(c?.in||0)}${makeNum(c?.out||0)}${makeNum(cNet||0)}` : ``}
      </tr>
      <tr class="detailRow" data-detail="${b.id}" style="display:none;">
        <td colspan="${compare ? 7 : 4}">
          <div class="detailBox">
            ${renderBlockDetail(b)}
          </div>
        </td>
      </tr>
    `;
  }

  html += `</tbody></table>`;
  el("cashTable").innerHTML = html;

  el("cashTable").onclick = (e) => {
    const tr = e.target.closest(".cashRow");
    if (!tr) return;
    const id = tr.getAttribute("data-id");
    const detail = el("cashTable").querySelector(`tr[data-detail="${id}"]`);
    if (!detail) return;
    detail.style.display = (detail.style.display === "none") ? "" : "none";
  };
}

function renderBlockDetail(b){
  if (b.detail && b.detail.length){
    const rows = b.detail.map(x =>
      `<div>${x.label}</div><div class="r">${fmtEUR(x.in||0)}</div><div class="r">${fmtEUR(x.out||0)}</div>`
    ).join("");

    return `
      <div class="mini">
        <div class="h">Detalle</div><div class="h r">Entra</div><div class="h r">Sale</div>
        ${rows}
      </div>
    `;
  }

  if (b.top){
    const rows = b.top.map(x => {
      const inn = x.amt > 0 ? x.amt : 0;
      const out = x.amt < 0 ? -x.amt : 0;
      return `<div>${x.name}</div><div class="r">${fmtEUR(inn)}</div><div class="r">${fmtEUR(out)}</div>`;
    }).join("");

    return `
      <div class="mini">
        <div class="h">Top movimientos</div><div class="h r">Entra</div><div class="h r">Sale</div>
        ${rows || `<div>Sin detalle</div><div class="r">-</div><div class="r">-</div>`}
      </div>
    `;
  }

  if (b.finDetail){
    const rows = b.finDetail.map(x =>
      `<div>${x.label}</div><div class="r">${fmtEUR(x.in||0)}</div><div class="r">${fmtEUR(x.out||0)}</div>`
    ).join("");

    return `
      <div class="mini">
        <div class="h">Detalle financiación</div><div class="h r">Entra</div><div class="h r">Sale</div>
        ${rows}
      </div>
    `;
  }

  if (b.note){
    return `<div style="color:var(--muted); font-size:12px;">${b.note}</div>`;
  }

  return `<div style="color:var(--muted); font-size:12px;">Sin detalle.</div>`;
}

// =========================
// Events
// =========================
el("cashCompare").addEventListener("change", () => { renderCashflow(); updateQA(); });

el("fileCompany").addEventListener("change", (e)=>{
  const file = e.target.files?.[0];
  if (!file) return;

  setStatus("Leyendo CSV empresa…");
  parseCsvFile(file, (err, rows) => {
    if (err){ setStatus("Error leyendo CSV empresa.", true); return; }

    companyRows = rows;
    const missing = validateCompanyColumns(companyRows);
    if (missing.length){ setStatus(`Faltan columnas en CSV empresa: ${missing.join(", ")}`, true); return; }

    buildIndex(companyRows);
    if (availableYears.length < 2){ setStatus("Necesito al menos 2 años (periodo anual YYYY) para comparar.", true); updateQA(); return; }

    el("baseYear").disabled = false;
    el("compYear").disabled = false;

    el("baseYear").innerHTML = availableYears.map(y => `<option value="${y}">${y}</option>`).join("");
    el("compYear").innerHTML = availableYears.map(y => `<option value="${y}">${y}</option>`).join("");

    el("baseYear").value = availableYears[availableYears.length - 1];
    el("compYear").value = availableYears[availableYears.length - 2];

    el("sectorToggle").disabled = Object.keys(sectorMap).length === 0;
    el("sectorToggle").checked = Object.keys(sectorMap).length > 0;
    el("sectorStatus").textContent = el("sectorToggle").checked ? "ON" : "OFF";

    el("cashCompare").disabled = false;
    el("cashCompare").checked = false;

    setStatus(`CSV empresa OK. Años detectados: ${availableYears.join(", ")}`);
    render();
  });
});

el("fileSector").addEventListener("change", (e)=>{
  const file = e.target.files?.[0];
  if (!file) return;

  setStatus("Leyendo CSV sector…");
  parseCsvFile(file, (err, rows) => {
    if (err){ setStatus("Error leyendo CSV sector.", true); return; }

    sectorMap = {};
    for (const r of rows){
      const kpi = String(r.KPI ?? r.kpi ?? "").trim();
      if (!kpi) continue;
      const min = toNumber(r.Min ?? r.min);
      const media = toNumber(r.Media ?? r.media);
      const max = toNumber(r.Max ?? r.max);
      sectorMap[kpi] = {min, media, max};
    }

    const hasAny = Object.keys(sectorMap).length > 0;
    el("sectorToggle").disabled = !hasAny;
    el("sectorToggle").checked = hasAny;
    el("sectorStatus").textContent = hasAny ? "ON" : "OFF";

    setStatus(hasAny ? "CSV sector OK. Comparativa sector activada." : "CSV sector sin datos válidos.", !hasAny);
    render();
  });
});

el("sectorToggle").addEventListener("change", ()=>{
  el("sectorStatus").textContent = el("sectorToggle").checked ? "ON" : "OFF";
  render();
});
el("baseYear").addEventListener("change", render);
el("compYear").addEventListener("change", render);

// Primer pintado QA vacío
updateQA();
// ===== Tabs (arriba) =====
(function initTabs(){
  const tabs = Array.from(document.querySelectorAll(".tab"));

  // OJO: aquí metemos OPERATIVO
  const panes = {
    home: document.getElementById("tab-home"),
    operativo: document.getElementById("tab-operativo"),
    cash: document.getElementById("tab-cash"),
    profit: document.getElementById("tab-profit"),
  };

  function activate(name){
    // activar botones
    tabs.forEach(t => t.classList.toggle("active", t.dataset.tab === name));

    // activar paneles
    Object.entries(panes).forEach(([k, pane]) => {
      if (!pane) return;
      pane.classList.toggle("active", k === name);
    });

    // Forzar resize de charts (evita “gráficos raros”)
    setTimeout(()=>{
      try { if (window.cashChartInstance) window.cashChartInstance.resize(); } catch(e){}
      try { if (window.chartInstance) window.chartInstance.resize(); } catch(e){}
      // Si existe el render específico de Operativo, lo llamamos al entrar
      try {
        if (name === "operativo" && typeof window.__renderOperativoNow === "function") {
          window.__renderOperativoNow();
        }
      } catch(e){}
    }, 50);
  }

  tabs.forEach(t => t.addEventListener("click", ()=> activate(t.dataset.tab)));

  // Arranque por defecto
  activate("home");
})();


// ===== Resumen en 20 segundos (Home) =====
(function initOwnerSummary(){
  const $ = (id) => document.getElementById(id);

  function getKpi(id){
    return (typeof KPI_DEFS !== "undefined") ? KPI_DEFS.find(k => k.id === id) : null;
  }
  function kpiVal(id, year){
    const k = getKpi(id);
    if (!k) return null;
    const v = k.calc(year);
    return (v === null || v === undefined || Number.isNaN(v)) ? null : v;
  }

  function badgeHTML(kind, text){
    const cls = kind === "good" ? "good" : (kind === "warn" ? "warn" : "bad");
    return `<span class="badge ${cls}">${text}</span>`;
  }

  function setSumCard(elId, kind, stateText, lineText){
    const box = $(elId);
    if (!box) return;
    box.querySelector(".sumState").innerHTML = badgeHTML(kind, stateText);
    box.querySelector(".sumLine").textContent = lineText;
  }

  function pctPoint(deltaPct){
    if (deltaPct === null || deltaPct === undefined || Number.isNaN(deltaPct)) return "N/D";
    // deltaPct viene en formato ratio (0.02 = +2%)
    const pp = deltaPct * 100;
    const sign = pp >= 0 ? "+" : "";
    return `${sign}${fmtNum(pp, 1)} pp`;
  }

  function arrow(delta){
    if (delta === null || delta === undefined || Number.isNaN(delta)) return "•";
    return delta > 0 ? "↑" : (delta < 0 ? "↓" : "→");
  }

  // Evaluaciones (semáforos)
  function evalProfit(ebitdaM, ebitda, baseY, compY){
    if (ebitdaM === null || ebitda === null) return {kind:"warn", state:"N/D", line:"Faltan datos para calcular rentabilidad."};

    const ebitdaMComp = kpiVal("EBITDA_M", compY);
    const improves = (ebitdaMComp !== null) ? (ebitdaM >= ebitdaMComp) : null;

    if (ebitda < 0 || ebitdaM < 0.05) {
      return {kind:"bad", state:"❌ Crítico", line:"Operativamente estás perdiendo. Prioridad: margen y costes."};
    }
    if (ebitdaM >= 0.12 && (improves === null || improves)) {
      return {kind:"good", state:"✅ Bien", line:"Buen margen: puedes invertir sin asfixiarte."};
    }
    return {kind:"warn", state:"⚠️ Ajustado", line:"Margen sensible: una caída de ventas te muerde."};
  }

  function evalCash(cashNet, oxy){
    if (cashNet === null || oxy === null) return {kind:"warn", state:"N/D", line:"Faltan datos para evaluar la caja."};

    if (oxy < 30) return {kind:"bad", state:"❌ Tensión", line:"Oxígeno bajo: prioridad cobro/caja ya."};
    if (cashNet > 0 && oxy >= 45) return {kind:"good", state:"✅ Cómodo", line:"Entra caja y tienes colchón."};
    return {kind:"warn", state:"⚠️ Vigila", line:"Caja sensible: vigila cobros y gastos fijos."};
  }

  function evalRisk(ndEb, ebitda){
    if (ndEb === null || ebitda === null) return {kind:"warn", state:"N/D", line:"Faltan datos para evaluar riesgo."};
    if (ebitda <= 0) return {kind:"bad", state:"❌ Tensión", line:"Con EBITDA ≤ 0 no hay margen de seguridad para deuda."};

    if (ndEb <= 3) return {kind:"good", state:"✅ Sano", line:"Deuda sostenible para tu capacidad de pago."};
    if (ndEb <= 5) return {kind:"warn", state:"⚠️ Exigente", line:"Deuda exigente: cuidado con inversiones y caja."};
    return {kind:"bad", state:"❌ Peligro", line:"Deuda peligrosa: estás muy justo para pagar."};
  }

  // Top cambios (3+3)
  function buildTopChanges(baseY, compY){
    const items = [
      {id:"EBITDA", label:"EBITDA", type:"EUR", better:"HIGH"},
      {id:"EBITDA_M", label:"Margen EBITDA", type:"PP", better:"HIGH"},
      {id:"CASH_NET", label:"Caja neta", type:"EUR", better:"HIGH"},
      {id:"DSO", label:"Días de cobro", type:"DAYS", better:"LOW"},
      {id:"OXY", label:"Días de oxígeno", type:"DAYS", better:"HIGH"},
      {id:"ND_EB", label:"Deuda neta / EBITDA", type:"X", better:"LOW"},
    ];

    const rows = [];
    for (const it of items){
      const b = kpiVal(it.id, baseY);
      const c = kpiVal(it.id, compY);
      if (b === null || c === null) continue;

      const delta = b - c; // base - comp
      // score para ordenar: si LOW es mejor, invertimos el delta
      const effectiveDelta = (it.better === "LOW") ? (-delta) : delta;

      // peso por tipo
      let score = Math.abs(effectiveDelta);
      if (it.type === "EUR") score = Math.abs(effectiveDelta) / 100000; // 100k€ = 1 punto
      if (it.type === "DAYS") score = Math.abs(effectiveDelta);         // 1 día = 1 punto
      if (it.type === "PP") score = Math.abs(effectiveDelta) * 100;     // 0.01 (=1pp) => 1 punto
      if (it.type === "X") score = Math.abs(effectiveDelta) * 10;       // 0.1x => 1 punto

      const improved = effectiveDelta > 0;
      rows.push({ ...it, b, c, delta, improved, score });
    }

    rows.sort((a,b)=> b.score - a.score);

    const improvements = rows.filter(r=> r.improved).slice(0,3);
    const worsens = rows.filter(r=> !r.improved).slice(0,3);

    function fmtRow(r){
      const d = r.delta;
      const arr = arrow(d);

      let tail = "";
      if (r.type === "EUR") tail = fmtEUR(d);
      else if (r.type === "DAYS") tail = `${fmtNum(d,0)} días`;
      else if (r.type === "PP") {
        // aquí delta es ratio (ej 0.02 = +2%), lo convertimos a pp
        tail = pctPoint(d);
      } else if (r.type === "X") {
        const sign = d >= 0 ? "+" : "";
        tail = `${sign}${fmtNum(d,1)}x`;
      } else tail = fmtNum(d,2);

      // Para LOW-better, si baja es mejora, pero delta será negativo. Mostramos el signo real.
      return `<li><b>${r.label}</b> ${arr} ${tail}</li>`;
    }

    const html = `
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
        <div>
          <div style="font-weight:900; margin-bottom:6px;">Mejoras</div>
          <ul class="sumList">${improvements.length ? improvements.map(fmtRow).join("") : "<li>N/D</li>"}</ul>
        </div>
        <div>
          <div style="font-weight:900; margin-bottom:6px;">Empeoramientos</div>
          <ul class="sumList">${worsens.length ? worsens.map(fmtRow).join("") : "<li>N/D</li>"}</ul>
        </div>
      </div>
    `;
    return html;
  }

  // Dónde está el dinero
  function buildWhereMoney(baseY){
    const cash = (typeof getAmount === "function") ? getAmount("BAL.ACT.B.VII", baseY) : null;
    const debtGross = kpiVal("DEBT_GROSS", baseY);

    let netDebt = null;
    if (debtGross !== null && cash !== null) netDebt = debtGross - cash;

    const oxy = kpiVal("OXY", baseY);

    return `
      <div style="display:grid; grid-template-columns: 1fr; gap:10px;">
        <div>💰 <b>Caja hoy:</b> ${cash === null ? "N/D" : fmtEUR(cash)}</div>
        <div>🏦 <b>Deuda neta:</b> ${netDebt === null ? "N/D" : fmtEUR(netDebt)}</div>
        <div>🫁 <b>Días de oxígeno:</b> ${oxy === null ? "N/D" : `${fmtNum(oxy,0)} días`}</div>
      </div>
      <div style="margin-top:10px; color:var(--muted); font-size:12px;">
        Caja = lo que tienes. Deuda neta = deuda bruta menos caja. Oxígeno = cuántos días aguantas sin ingresar.
      </div>
    `;
  }

  function renderOwnerSummary(){
    const hint = $("sumHint");
    const sumChanges = $("sumChanges")?.querySelector(".sumWideBody");
    const sumWhere = $("sumWhere")?.querySelector(".sumWideBody");

    if (!companyRows || !companyRows.length || !availableYears || availableYears.length < 2){
      if (hint) hint.textContent = "Sube el CSV de empresa.";
      setSumCard("sumProfit","warn","—","Sube datos para ver el estado.");
      setSumCard("sumCash","warn","—","Sube datos para ver el estado.");
      setSumCard("sumRisk","warn","—","Sube datos para ver el estado.");
      if (sumChanges) sumChanges.textContent = "—";
      if (sumWhere) sumWhere.textContent = "—";
      return;
    }

    const baseY = document.getElementById("baseYear").value || availableYears[availableYears.length-1];
    const compY = document.getElementById("compYear").value || availableYears[availableYears.length-2];
    const sectorOn = document.getElementById("sectorToggle").checked;

    if (hint) hint.textContent = `Base ${baseY} vs ${compY} · Sector ${sectorOn ? "ON" : "OFF"}`;

    const ebitdaM = kpiVal("EBITDA_M", baseY);
    const ebitda = kpiVal("EBITDA", baseY);
    const cashNet = kpiVal("CASH_NET", baseY);
    const oxy = kpiVal("OXY", baseY);
    const ndEb = kpiVal("ND_EB", baseY);

    const p = evalProfit(ebitdaM, ebitda, baseY, compY);
    const c = evalCash(cashNet, oxy);
    const r = evalRisk(ndEb, ebitda);

    setSumCard("sumProfit", p.kind, p.state, p.line);
    setSumCard("sumCash", c.kind, c.state, c.line);
    setSumCard("sumRisk", r.kind, r.state, r.line);

    if (sumChanges) sumChanges.innerHTML = buildTopChanges(baseY, compY);
    if (sumWhere) sumWhere.innerHTML = buildWhereMoney(baseY);
  }

  // Enganchamos el resumen al render() existente sin romper nada
  if (typeof render === "function"){
    const _render = render;
    render = function(){
      _render();
      renderOwnerSummary();
    };
  }

  // Primer pintado por si ya hay datos cargados
  try { renderOwnerSummary(); } catch(e){}
})();
// ===== Visual Upgrade V4 (EFE perfecto + Macros PRO + NO toca modales KPI) =====
(function visualsUpgradeV4(){
  // --------- Formatos ---------
  // Etiquetas MACRO: valores en millones SIN "M€"
  function fmtM(n, d=1){
    if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
    return fmtNum(n / 1e6, d);
  }
  // Ejes MACRO: sí muestran unidad para contexto
  function fmtAxisM(n){
    if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
    return `${fmtNum(n / 1e6, 0)}M€`;
  }
  // EFE: mantiene M€ en cápsula (como te gusta)
  function fmtMM(n){
    if (n === null || n === undefined || Number.isNaN(n)) return "N/D";
    const sign = n < 0 ? "-" : "";
    const v = Math.abs(n) / 1e6;
    return `${sign}${fmtNum(v, 1)}M€`;
  }

  function clamp(v, min, max){ return Math.min(max, Math.max(min, v)); }

  // --------- Etiquetas EFE (mantener) ---------
  function shortEfeLabel(label){
    const s = String(label || "").toLowerCase();
    if (s.includes("mi negocio")) return "Negocio";
    if (s.includes("invertir en corriente") || s.includes("corriente")) return "Corriente";
    if (s.includes("intereses") || s.includes("impuestos")) return "Intereses/Imp.";
    if (s.includes("inversiones")) return "Inversiones";
    if (s.includes("desinversiones")) return "Desinv.";
    if (s.includes("financi")) return "Financiación";
    if (s.includes("caja que se queda")) return "Caja se queda";
    if (s.includes("caja que uso")) return "Caja uso";
    return String(label || "").replace(/^Entra\s+por\s+/i,"").replace(/^Sale\s+por\s+/i,"").trim();
  }

  // ============================================================
  // 1) Plugin MACROS: texto pequeño, sin cápsula, evita solapes
  //    - Muestra TODOS los años
  //    - Si se pisa: intenta alejar; si no, se omite ese año
  // ============================================================
  const macroLabelsPluginV2 = {
    id: "macroLabelsPluginV2",
    afterDatasetsDraw(chart, args, opts){
      const { ctx, chartArea } = chart;
      if (!chartArea) return;

      const padR = (opts?.padRight ?? 22);
      const padT = (opts?.padTop ?? 12);

      // Para anti-solape: guardamos cajas ya dibujadas
      const placed = [];
      function collides(box){
        return placed.some(b =>
          box.x < b.x + b.w && box.x + box.w > b.x &&
          box.y < b.y + b.h && box.y + box.h > b.y
        );
      }
      function place(box){ placed.push(box); }

      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.shadowColor = "rgba(0,0,0,.55)";
      ctx.shadowBlur = 5;

      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;

        const mode = ds._labelMode || "M";
        const isLine = ds.type === "line";
        const isBar  = !ds.type || ds.type === "bar";

        meta.data.forEach((el, i) => {
          const v = ds.data[i];
          if (v === null || v === undefined || Number.isNaN(v)) return;

          // texto principal/secundario
          let t1 = "", t2 = "";
          if (mode === "M") t1 = fmtM(v, 1);
          else if (mode === "M0") t1 = fmtM(v, 0);
          else if (mode === "EBITDA_M_WITH_MARGIN"){
            const sales = opts?.salesSeries?.[i] ?? null;
            const m = (sales && sales !== 0) ? (v / sales) : null;
            t1 = fmtM(v, 1);
            t2 = (m === null) ? "" : `${fmtNum(m*100,1)}%`; // aquí sí porcentaje
          } else {
            t1 = fmtM(v, 1);
          }

          // posición base
          const p = el.tooltipPosition();
          let x = clamp(p.x, chartArea.left + 10, chartArea.right - padR);
          let y = p.y;

          // tamaño fuente (ligeramente más pequeño)
          const font1 = "11px system-ui";
          const font2 = "10px system-ui";

          // candidatos de posición (para no solapar)
          const candidates = [];

          if (isLine){
            // arriba (preferido), luego más arriba, luego abajo
            candidates.push({dy:-12}, {dy:-22}, {dy:+14});
          } else if (isBar){
            // centro (preferido), si barra pequeña: arriba
            const props = el.getProps(["y","base","x"], true);
            const h = Math.abs(props.base - props.y);
            if (h < 18) candidates.push({dy:-12}, {dy:-22}, {dy:+14});
            else candidates.push({y:(props.y + props.base)/2}, {dy:-12}, {dy:+14});
          }

          let drawn = false;

          for (const cand of candidates){
            let yy = (cand.y !== undefined) ? cand.y : (y + (cand.dy || 0));
            yy = clamp(yy, chartArea.top + padT, chartArea.bottom - 10);

            // bounding box aproximada (con 1 o 2 líneas)
            ctx.font = font1;
            const w1 = ctx.measureText(t1).width;
            ctx.font = font2;
            const w2 = t2 ? ctx.measureText(t2).width : 0;
            const w = Math.max(w1, w2);
            const hBox = t2 ? 24 : 14;

            const box = { x: x - w/2 - 2, y: yy - hBox/2, w: w + 4, h: hBox };

            // si choca, probamos siguiente
            if (collides(box)) continue;

            // dibujar
            ctx.font = font1;
            ctx.fillStyle = "rgba(255,255,255,.92)";
            ctx.fillText(t1, x, t2 ? (yy - 6) : yy);

            if (t2){
              ctx.font = font2;
              ctx.fillStyle = "rgba(255,255,255,.78)";
              ctx.fillText(t2, x, yy + 8);
            }

            place(box);
            drawn = true;
            break;
          }

          // si no se pudo colocar sin chocar, se omite ese año (preferible a “cutre”)
          if (!drawn) return;
        });
      });

      ctx.restore();
    }
  };

  // ============================================================
  // 2) Plugin EFE: cápsula centrada con nombre + importe (mantener)
  // ============================================================
  const cashStackLabelPlugin = {
    id: "cashStackLabelPlugin",
    afterDatasetsDraw(chart){
      const { ctx, chartArea } = chart;
      if (!chartArea) return;

      const yScale = chart.scales?.y;
      const range = Math.abs((yScale?.max ?? 1) - (yScale?.min ?? 0)) || 1;
      const threshold = range * 0.08;

      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;

        meta.data.forEach((el, i) => {
          const v = ds.data[i];
          if (!v || Number.isNaN(v)) return;
          if (Math.abs(v) < threshold) return;

          const props = el.getProps(["x","y","base"], true);
          let x = clamp(props.x, chartArea.left + 14, chartArea.right - 14);
          let y = clamp((props.y + props.base) / 2, chartArea.top + 16, chartArea.bottom - 16);

          const name = shortEfeLabel(ds.label);
          const val = fmtMM(v);

          const font1 = "12px system-ui";
          const font2 = "11px system-ui";
          ctx.font = font2;
          const wName = ctx.measureText(name).width;
          ctx.font = font1;
          const wVal = ctx.measureText(val).width;
          const w = Math.max(wName, wVal);
          const h = 30;

          const bx = x - (w/2) - 8;
          const by = y - (h/2);
          const bw = w + 16;
          const bh = h;

          ctx.fillStyle = "rgba(0,0,0,.35)";
          ctx.strokeStyle = "rgba(255,255,255,.12)";
          ctx.lineWidth = 1;

          const r = 8;
          ctx.beginPath();
          ctx.moveTo(bx + r, by);
          ctx.lineTo(bx + bw - r, by);
          ctx.quadraticCurveTo(bx + bw, by, bx + bw, by + r);
          ctx.lineTo(bx + bw, by + bh - r);
          ctx.quadraticCurveTo(bx + bw, by + bh, bx + bw - r, by + bh);
          ctx.lineTo(bx + r, by + bh);
          ctx.quadraticCurveTo(bx, by + bh, bx, by + bh - r);
          ctx.lineTo(bx, by + r);
          ctx.quadraticCurveTo(bx, by, bx + r, by);
          ctx.closePath();
          ctx.fill();
          ctx.stroke();

          ctx.fillStyle = "rgba(255,255,255,.86)";
          ctx.font = font2;
          ctx.fillText(name, x, y - 7);

          ctx.fillStyle = "rgba(255,255,255,.95)";
          ctx.font = font1;
          ctx.fillText(val, x, y + 9);
        });
      });

      ctx.restore();
    }
  };

  // ============================================================
  // 3) Macros HOME: Ventas/Deuda línea; EBITDA/BEP barra
  //    - Leyenda con forma: la forma ya la da Chart.js, texto solo el nombre
  // ============================================================
  let homeSalesEbitdaChart = null;
  let homeSalesBepChart = null;
  let homeDebtEbitdaChart = null;

  function netDebtYear(year){
    const cash = (typeof getAmount === "function") ? getAmount("BAL.ACT.B.VII", year) : null;
    const gross = (typeof KPI_DEFS !== "undefined")
      ? (KPI_DEFS.find(k=>k.id==="DEBT_GROSS")?.calc(year) ?? null)
      : null;
    if (cash === null || gross === null) return null;
    return gross - cash;
  }

  function scalesMacro(){
    return {
      x:{ ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } },
      y:{
        ticks:{ color:"#aab6d6", maxTicksLimit: 4, callback:(v)=> fmtAxisM(v) },
        grid:{ color:"rgba(255,255,255,.06)" }
      }
    };
  }

  function legendTextOnly(){
    const original = Chart.defaults.plugins.legend.labels.generateLabels;
    return function(c){
      const labels = original(c);
      labels.forEach(l => {
        const ds = c.data.datasets[l.datasetIndex];
        l.text = ds.label; // solo nombre (la forma ya se ve en el icono)
      });
      return labels;
    };
  }

  function renderHomeMacroCharts(){
    if (!companyRows?.length || !availableYears?.length) return;

    const years = availableYears.slice(-6);

    const sales = years.map(y => (typeof getAmount === "function") ? getAmount("PYG.MAIN.1", y) : null);
    const ebitda = years.map(y => KPI_DEFS.find(k=>k.id==="EBITDA")?.calc(y) ?? null);
    const bep = years.map(y => KPI_DEFS.find(k=>k.id==="BEP")?.calc(y) ?? null);
    const nd = years.map(y => netDebtYear(y));

    // 1) Ventas (línea) vs EBITDA (barra) + etiqueta EBITDA con %
    const c1 = document.getElementById("homeChartSalesEbitda");
    if (c1){
      if (homeSalesEbitdaChart) homeSalesEbitdaChart.destroy();

      homeSalesEbitdaChart = new Chart(c1.getContext("2d"), {
        type:"bar",
        data:{
          labels: years,
          datasets:[
            { label:"EBITDA", data: ebitda, borderWidth:1, _labelMode:"EBITDA_M_WITH_MARGIN" }, // barra
            { label:"Ventas", data: sales, type:"line", borderWidth:2, pointRadius:3, tension:.25, _labelMode:"M" } // línea
          ]
        },
        options:{
          responsive:true,
          layout:{ padding:{ right: 30, top: 14 } },
          plugins:{
            legend:{ labels:{ color:"#cbd5ff", generateLabels: legendTextOnly() } },
            tooltip:{
              callbacks:{
                label:(c)=>{
                  const v = c.parsed.y;
                  if (c.dataset.label === "EBITDA"){
                    const s = sales[c.dataIndex];
                    const m = (s && s !== 0) ? (v/s) : null;
                    return `EBITDA: ${fmtEUR(v)} (${m===null?"N/D":fmtPct(m)})`;
                  }
                  return `${c.dataset.label}: ${fmtEUR(v)}`;
                }
              }
            },
            macroLabelsPluginV2:{ salesSeries: sales, padRight: 30, padTop: 14 }
          },
          scales: scalesMacro()
        },
        plugins:[macroLabelsPluginV2]
      });
    }

    // 2) Ventas (línea) vs BEP (barra)
    const c2 = document.getElementById("homeChartSalesBep");
    if (c2){
      if (homeSalesBepChart) homeSalesBepChart.destroy();

      homeSalesBepChart = new Chart(c2.getContext("2d"), {
        type:"bar",
        data:{
          labels: years,
          datasets:[
            { label:"Punto de equilibrio", data: bep, borderWidth:1, _labelMode:"M" }, // barra
            { label:"Ventas", data: sales, type:"line", borderWidth:2, pointRadius:3, tension:.25, _labelMode:"M" } // línea
          ]
        },
        options:{
          responsive:true,
          layout:{ padding:{ right: 30, top: 14 } },
          plugins:{
            legend:{ labels:{ color:"#cbd5ff", generateLabels: legendTextOnly() } },
            tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${fmtEUR(c.parsed.y)}` } },
            macroLabelsPluginV2:{ padRight: 30, padTop: 14 }
          },
          scales: scalesMacro()
        },
        plugins:[macroLabelsPluginV2]
      });
    }

    // 3) Deuda neta (línea) vs EBITDA (barra)
    const c3 = document.getElementById("homeChartDebtEbitda");
    if (c3){
      if (homeDebtEbitdaChart) homeDebtEbitdaChart.destroy();

      homeDebtEbitdaChart = new Chart(c3.getContext("2d"), {
        type:"bar",
        data:{
          labels: years,
          datasets:[
            { label:"EBITDA", data: ebitda, borderWidth:1, _labelMode:"M" }, // barra
            { label:"Deuda neta", data: nd, type:"line", borderWidth:2, pointRadius:3, tension:.25, _labelMode:"M" } // línea
          ]
        },
        options:{
          responsive:true,
          layout:{ padding:{ right: 30, top: 14 } },
          plugins:{
            legend:{ labels:{ color:"#cbd5ff", generateLabels: legendTextOnly() } },
            tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${fmtEUR(c.parsed.y)}` } },
            macroLabelsPluginV2:{ padRight: 30, padTop: 14 }
          },
          scales: scalesMacro()
        },
        plugins:[macroLabelsPluginV2]
      });
    }
  }

  // ============================================================
  // 4) EFE: recrea chart con plugin cápsula (mantener)
  // ============================================================
  if (typeof renderCashflow === "function"){
    const _rcf = renderCashflow;
    renderCashflow = function(){
      _rcf();
      try{
        if (!cashChartInstance) return;
        if (cashChartInstance.config._cashLabelV4) return;

        const canvas = document.getElementById("cashChart");
        const cfg = cashChartInstance.config;
        const data = JSON.parse(JSON.stringify(cfg.data));
        const options = JSON.parse(JSON.stringify(cfg.options || {}));

        options.layout = options.layout || {};
        options.layout.padding = options.layout.padding || {};
        options.layout.padding.top = Math.max(options.layout.padding.top || 0, 18);
        options.layout.padding.right = Math.max(options.layout.padding.right || 0, 24);

        cashChartInstance.destroy();
        cashChartInstance = new Chart(canvas.getContext("2d"), {
          type:"bar",
          data,
          options,
          plugins:[cashStackLabelPlugin]
        });
        cashChartInstance.config._cashLabelV4 = true;
      }catch(e){}
    };
  }

  // ============================================================
  // 5) Hook render(): SOLO repinta macros + EFE (NO toca modales KPI)
  // ============================================================
  if (typeof render === "function"){
    const _render = render;
    render = function(){
      _render();
      try { renderHomeMacroCharts(); } catch(e){}
    };
  }

  try { renderHomeMacroCharts(); } catch(e){}
})();
// ===== Plugin GLOBAL de etiquetas (formato KPI) - NO toca macros ni EFE =====
(function addGlobalLabelsV2(){
  function clamp(v, min, max){ return Math.min(max, Math.max(min, v)); }

  function isMoneyScale(scale){
    const mx = Math.max(Math.abs(scale?.max ?? 0), Math.abs(scale?.min ?? 0));
    return mx >= 100000; // a partir de 100k asumimos € (EBITDA, ventas, etc.)
  }

  function isRatioScale(scale){
    // ratios típicos 0..1 (margen bruto, EBITDA/ventas, etc.)
    const mx = scale?.max ?? 0;
    const mn = scale?.min ?? 0;
    return (mn >= -0.5 && mx <= 2.0);
  }

  function fmtMoneyM(value){
    // millones con 1 decimal + M€
    return `${fmtNum(value / 1e6, 1)}M€`;
  }

  function fmtPercent1(value){
    // value viene como 0,653 -> 65,3%
    return `${fmtNum(value * 100, 1)}%`;
  }

  function fmtDefault(value){
    // fallback general (días, ratios x, etc.)
    // si es casi entero, lo dejamos entero
    const rounded = Math.round(value);
    if (Math.abs(value - rounded) < 1e-9) return fmtNum(rounded, 0);
    return fmtNum(value, 1);
  }

  // Anti-solape simple
  function collides(a, b){
    return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
  }

  const globalLabelsPlugin = {
    id: "globalLabelsPlugin",
    afterDatasetsDraw(chart){
      const id = chart?.canvas?.id || "";

      // No tocar los 3 macros (ya van con su plugin)
      if (id.startsWith("homeChart")) return;
      // No tocar EFE (ya va con cápsulas)
      if (id === "cashChart") return;

      const { ctx, chartArea } = chart;
      if (!chartArea) return;

      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.font = "12px system-ui";
      ctx.fillStyle = "rgba(255,255,255,.92)";
      ctx.shadowColor = "rgba(0,0,0,.55)";
      ctx.shadowBlur = 5;

      const placed = [];

      // Control densidad: si hay muchos puntos, no ensuciamos
      const n = chart.data.labels?.length || 0;
      const allowAll = n <= 14;
      const every = n <= 24 ? 2 : 3;

      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;

        const scaleId = ds.yAxisID || "y";
        const scale = chart.scales?.[scaleId];
        const money = isMoneyScale(scale);
        const ratio = isRatioScale(scale);

        const isLine = (ds.type === "line" || meta.type === "line");
        const isBar  = (ds.type === "bar" || meta.type === "bar" || !ds.type);

        meta.data.forEach((el, i) => {
          const v = ds.data?.[i];
          if (v === null || v === undefined || Number.isNaN(v)) return;

          if (!allowAll && (i % every !== 0) && i !== n-1) return;

          let text = "";
          if (money) text = fmtMoneyM(v);
          else if (ratio) text = fmtPercent1(v);
          else text = fmtDefault(v);

          if (!text) return;

          const p = el.tooltipPosition();
          let x = clamp(p.x, chartArea.left + 10, chartArea.right - 10);
          let y = p.y;

          const candidates = [];
          if (isLine){
            candidates.push({dy:-14}, {dy:-24}, {dy:+14});
          } else if (isBar){
            try{
              const props = el.getProps(["y","base"], true);
              const h = Math.abs(props.base - props.y);
              if (h >= 18) candidates.push({y:(props.y + props.base)/2});
              candidates.push({dy:-14}, {dy:+14});
            }catch(e){
              candidates.push({dy:-14}, {dy:+14});
            }
          } else {
            candidates.push({dy:-14}, {dy:+14});
          }

          let done = false;

          for (const c of candidates){
            let yy = (c.y !== undefined) ? c.y : (y + (c.dy || 0));
            yy = clamp(yy, chartArea.top + 12, chartArea.bottom - 12);

            const w = ctx.measureText(text).width;
            const box = { x: x - w/2 - 3, y: yy - 8, w: w + 6, h: 16 };

            if (placed.some(b => collides(box, b))) continue;

            ctx.fillText(text, x, yy);
            placed.push(box);
            done = true;
            break;
          }

          if (!done) return;
        });
      });

      ctx.restore();
    }
  };

  // Registrar 1 sola vez
  if (!Chart.registry.plugins.get("globalLabelsPlugin")) {
    Chart.register(globalLabelsPlugin);
  }

  // repintado
  try { if (typeof render === "function") render(); } catch(e){}
})();
// ===== Operativo (PyG oficial) - Waterfall pro + etiquetas EBITDA% + Drill-down =====
(function operativoModule(){
  if (window.__operativoReady) return;
  window.__operativoReady = true;

  const $ = (id) => document.getElementById(id);

  // Fallbacks por si algo no está en window
  const fmtNumSafe = (typeof fmtNum === "function")
    ? fmtNum
    : (n, d=0) => new Intl.NumberFormat("es-ES",{maximumFractionDigits:d,minimumFractionDigits:d}).format(n);

  const fmtPctSafe = (typeof fmtPct === "function")
    ? fmtPct
    : (n) => new Intl.NumberFormat("es-ES",{style:"percent",minimumFractionDigits:1,maximumFractionDigits:1}).format(n);

  const fmtEURSafe = (typeof fmtEUR === "function")
    ? fmtEUR
    : (n) => new Intl.NumberFormat("es-ES",{style:"currency",currency:"EUR",maximumFractionDigits:0}).format(n);

  function fmtM(v, d=1){
    if (v === null || v === undefined || Number.isNaN(v)) return "N/D";
    const sign = v < 0 ? "-" : "";
    return `${sign}${fmtNumSafe(Math.abs(v)/1e6, d)}M€`;
  }
  function fmtAxisM(v){
    if (v === null || v === undefined || Number.isNaN(v)) return "N/D";
    return `${fmtNumSafe(v/1e6, 0)}M€`;
  }

  function getBaseYear(){
    const node = $("baseYear");
    if (node && node.value) return parseInt(node.value, 10);
    return (window.availableYears?.length) ? +window.availableYears[window.availableYears.length-1] : 2023;
  }
  function getCompYear(){
    const node = $("compYear");
    if (node && node.value) return parseInt(node.value, 10);
    const ys = window.availableYears || [];
    return ys.length >= 2 ? +ys[ys.length-2] : getBaseYear();
  }

  function isOperativoVisible(){
    const pane = $("tab-operativo");
    return !!(pane && pane.classList.contains("active"));
  }

  // IMPORTANTÍSIMO: coherencia con HOME => usamos getAmount (index)
  function sumCodeYear(code, year){
    if (typeof getAmount === "function") return getAmount(code, year);
    // fallback si algo fallara
    try {
      return (window.amountByCodeYear?.get(`${code}__${String(year).trim()}`)) ?? 0;
    } catch(e){
      return 0;
    }
  }

  function nameOf(code){
    // buscamos en companyRows si existe
    try{
      const rows = (typeof companyRows !== "undefined" && Array.isArray(companyRows)) ? companyRows : (window.companyRows || []);
      const r = rows.find(x => String(x.concept_code||"").trim() === code);
      return (r?.display_name || r?.normalized_name || code);
    }catch(e){
      return code;
    }
  }

  // KPI refs
  function kpi(id){
    const list = (typeof KPI_DEFS !== "undefined") ? KPI_DEFS : (window.KPI_DEFS || []);
    return list.find(x => x.id === id) || null;
  }

  // ------------------------
  // PyG oficial (estructura “dueño”)
  // ------------------------
  const OP_WF_OWNER = [
    { key:"SALES", label:["Ventas"], codes:["PYG.MAIN.1"], drill:["PYG.MAIN.1.a","PYG.MAIN.1.b"], totalStyle:true },
    { key:"ACT_ADJ", label:["Ajustes","actividad"], codes:["PYG.MAIN.2","PYG.MAIN.3"], drill:["PYG.MAIN.2","PYG.MAIN.3"] },
    { key:"OTHER_INC", label:["Otros","ingresos"], codes:["PYG.MAIN.5"], drill:["PYG.MAIN.5.a","PYG.MAIN.5.b"] },
    { key:"COGS", label:["Aprov.","(coste dir.)"], codes:["PYG.MAIN.4"], drill:["PYG.MAIN.4.a","PYG.MAIN.4.b","PYG.MAIN.4.c","PYG.MAIN.4.d"] },
    { key:"GROSS", label:["Margen","bruto"], isTotal:true, buildFrom:["SALES","ACT_ADJ","OTHER_INC","COGS"], totalStyle:true },

    { key:"STAFF", label:["Personal"], codes:["PYG.MAIN.6"], drill:["PYG.MAIN.6.a","PYG.MAIN.6.b","PYG.MAIN.6.c"] },
    { key:"OPEX", label:["Otros","gastos"], codes:["PYG.MAIN.7"], drill:["PYG.MAIN.7.a","PYG.MAIN.7.b","PYG.MAIN.7.c","PYG.MAIN.7.d"] },
    { key:"EBITDA", label:["EBITDA"], isTotal:true, buildFrom:["GROSS","STAFF","OPEX"], totalStyle:true },
  ];

  const OP_WF_CONT = [
    { key:"DA", label:["Amort."], codes:["PYG.MAIN.8"], drill:["PYG.MAIN.8"] },
    { key:"ADJ_CONT", label:["Ajustes","contables"], codes:["PYG.MAIN.9","PYG.MAIN.10"], drill:["PYG.MAIN.9","PYG.MAIN.10"] },
    { key:"IMPAIR", label:["Deterioros"], codes:["PYG.MAIN.11"], drill:["PYG.MAIN.11.a","PYG.MAIN.11.b"] },
    { key:"EBIT", label:["Rdo.","explot."], isTotal:true, codes:["PYG.MAIN.A.1"], drill:["PYG.MAIN.A.1"], totalStyle:true }
  ];

  function computeBlockValue(block, year, cache){
    if (block.isTotal){
      if (block.codes?.length){
        return block.codes.reduce((acc,c)=> acc + sumCodeYear(c, year), 0);
      }
      let s = 0;
      for (const k of (block.buildFrom || [])) s += (cache[k] ?? 0);
      return s;
    }
    return (block.codes || []).reduce((acc,c)=> acc + sumCodeYear(c, year), 0);
  }

  // ------------------------
  // Waterfall series (offset + value + color)
  // ------------------------
  function buildWaterfallSeries(year, contableOn){
    const blocks = contableOn ? OP_WF_OWNER.concat(OP_WF_CONT) : OP_WF_OWNER.slice();

    const cache = {};
    blocks.forEach(b => cache[b.key] = computeBlockValue(b, year, cache));

    const labels = [];
    const offsets = [];
    const values = [];
    const colors = [];
    const meta = [];
    const mapByKey = {};

    let cum = 0;

    for (let i=0;i<blocks.length;i++){
      const b = blocks[i];
      const v = cache[b.key] ?? 0;

      labels.push(b.label); // array -> multiline label en Chart.js

      if (b.isTotal){
        offsets.push(0);
        values.push(v);
        cum = v;
      } else {
        offsets.push(cum);
        values.push(v);
        cum += v;
      }

      // Colores “tipo PyG”
      // total: azul; subida: verde; bajada: naranja
      const isTotalStyle = !!b.totalStyle;
      if (isTotalStyle) colors.push("rgba(96,165,250,.75)");
      else if (v >= 0) colors.push("rgba(34,197,94,.68)");
      else colors.push("rgba(245,158,11,.78)");

      const packed = { ...b, value: v };
      meta.push(packed);
      mapByKey[b.key] = v;
    }

    return { labels, offsets, values, colors, meta, mapByKey };
  }

  // Etiquetas en waterfall (cápsula solo valor)
  const waterfallLabelPlugin = {
    id: "opWaterfallLabelPlugin",
    afterDatasetsDraw(chart){
      const { ctx, chartArea } = chart;
      if (!chartArea) return;

      const ds = chart.data.datasets?.[1];
      const meta = chart.getDatasetMeta(1);
      if (!ds || !meta) return;

      const yScale = chart.scales?.y;
      const range = Math.abs((yScale?.max ?? 1) - (yScale?.min ?? 0)) || 1;
      const threshold = range * 0.07;

      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.shadowColor = "rgba(0,0,0,.55)";
      ctx.shadowBlur = 6;

      meta.data.forEach((bar, i) => {
        const v = ds.data[i];
        if (v === null || v === undefined || Number.isNaN(v)) return;
        if (Math.abs(v) < threshold) return;

        const props = bar.getProps(["x","y","base"], true);
        const x = props.x;
        const y = (props.y + props.base)/2;

        const t = fmtM(v, 1);

        ctx.font = "12px system-ui";
        const w = ctx.measureText(t).width;

        const bx = x - (w/2) - 8;
        const by = y - 14;
        const bw = w + 16;
        const bh = 28;

        ctx.fillStyle = "rgba(0,0,0,.30)";
        ctx.strokeStyle = "rgba(255,255,255,.12)";
        ctx.lineWidth = 1;

        const r = 9;
        ctx.beginPath();
        ctx.moveTo(bx + r, by);
        ctx.lineTo(bx + bw - r, by);
        ctx.quadraticCurveTo(bx + bw, by, bx + bw, by + r);
        ctx.lineTo(bx + bw, by + bh - r);
        ctx.quadraticCurveTo(bx + bw, by + bh, bx + bw - r, by + bh);
        ctx.lineTo(bx + r, by + bh);
        ctx.quadraticCurveTo(bx, by + bh, bx, by + bh - r);
        ctx.lineTo(bx, by + r);
        ctx.quadraticCurveTo(bx, by, bx + r, by);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = "rgba(255,255,255,.95)";
        ctx.fillText(t, x, y);
      });

      ctx.restore();
    }
  };

  // Conectores entre barras (estilo waterfall clásico)
  const waterfallConnectorPlugin = {
    id: "opWaterfallConnectorPlugin",
    afterDatasetsDraw(chart){
      const cfg = chart?.config || {};
      const wfData = cfg._wfData;
      if (!wfData) return;

      const meta = chart.getDatasetMeta(1); // dataset de valores
      const bars = meta?.data || [];
      if (!bars.length) return;

      const yScale = chart.scales?.y;
      if (!yScale) return;

      const offsets = wfData.offsets || [];
      const values = wfData.values || [];

      const ctx = chart.ctx;
      ctx.save();
      ctx.strokeStyle = "rgba(255,255,255,.18)";
      ctx.lineWidth = 2;

      for (let i=0; i<bars.length-1; i++){
        const bar = bars[i];
        const next = bars[i+1];
        if (!bar || !next) continue;

        const p1 = bar.getProps(["x","width"], true);
        const p2 = next.getProps(["x","width"], true);

        const endVal = (offsets[i] ?? 0) + (values[i] ?? 0);
        const nextStart = offsets[i+1] ?? 0;

        const y1 = yScale.getPixelForValue(endVal);
        const y2 = yScale.getPixelForValue(nextStart);

        const x1 = p1.x + (p1.width ? p1.width/2 : 0);
        const x2 = p2.x - (p2.width ? p2.width/2 : 0);

        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
      }

      ctx.restore();
    }
  };

  // Etiqueta especial para EBITDA: importe + % debajo (en el gráfico “La película”)
  const ebitdaTwoLineLabelPlugin = {
    id: "opEbitdaTwoLineLabelPlugin",
    afterDatasetsDraw(chart){
      // solo para el canvas de “La película”
      if ((chart?.canvas?.id || "") !== "opChartSalesEbitda") return;

      const { ctx, chartArea } = chart;
      if (!chartArea) return;

      // dataset 0 = EBITDA (barra), dataset 1 = Ventas (línea)
      const dsE = chart.data.datasets?.[0];
      const dsS = chart.data.datasets?.[1];
      const metaE = chart.getDatasetMeta(0);
      if (!dsE || !dsS || !metaE) return;

      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.shadowColor = "rgba(0,0,0,.55)";
      ctx.shadowBlur = 6;

      metaE.data.forEach((bar, i) => {
        const ebitda = dsE.data[i];
        const sales = dsS.data[i];

        if (ebitda === null || ebitda === undefined || Number.isNaN(ebitda)) return;

        const pct = (sales && sales !== 0) ? (ebitda / sales) : null;

        const props = bar.getProps(["x","y","base"], true);
        const x = props.x;
        const y = (props.y + props.base) / 2;

        const t1 = fmtM(ebitda, 1);
        const t2 = (pct === null) ? "N/D" : fmtPctSafe(pct);

        ctx.font = "12px system-ui";
        ctx.fillStyle = "rgba(255,255,255,.95)";
        ctx.fillText(t1, x, y - 7);

        ctx.font = "11px system-ui";
        ctx.fillStyle = "rgba(255,255,255,.78)";
        ctx.fillText(t2, x, y + 10);
      });

      ctx.restore();
    }
  };

  // ------------------------
  // Drill-down panel
  // ------------------------
  let opDetailChart = null;

  function renderDetail(metaItem, year){
    $("opDetailTitle").textContent = metaItem.label?.join ? metaItem.label.join(" ") : String(metaItem.label);
    $("opDetailSub").textContent = `Año ${year} · ${fmtM(metaItem.value)} · Desglose y evolución.`;

    const drill = metaItem.drill || metaItem.codes || [];
    const rows = drill.map(code => ({
      code,
      name: nameOf(code),
      value: sumCodeYear(code, year)
    })).filter(x => x.value !== 0 || x.code === "PYG.MAIN.A.1");

    rows.sort((a,b)=> Math.abs(b.value) - Math.abs(a.value));

    $("opDetailTable").innerHTML = `
      <table class="tbl">
        <thead>
          <tr><th>Partida</th><th style="text-align:right;">${year}</th></tr>
        </thead>
        <tbody>
          ${rows.map(r=>`
            <tr>
              <td>${r.name}</td>
              <td style="text-align:right;">${fmtM(r.value)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    `;

    const years = (window.availableYears || []).slice(-4).map(Number);
    const contableOn = !!$("opToggleContable")?.checked;

    // Evolución en gráfico
    const evoVals = years.map(y => {
      const wfMap = buildWaterfallSeries(y, contableOn).mapByKey || {};
      if (wfMap[metaItem.key] !== undefined) return wfMap[metaItem.key];
      // fallback suma de códigos
      return (metaItem.codes || []).reduce((acc,c)=> acc + sumCodeYear(c, y), 0);
    });

    if (opDetailChart){ opDetailChart.destroy(); opDetailChart = null; }
    const ctx = $("opDetailChart")?.getContext("2d");
    if (ctx){
      opDetailChart = new Chart(ctx, {
        type:"bar",
        data:{
          labels: years,
          datasets:[{
            label: metaItem.label?.join ? metaItem.label.join(" ") : metaItem.label,
            data: evoVals,
            borderWidth:1,
            backgroundColor: evoVals.map(v => v >= 0 ? "rgba(34,197,94,.65)" : "rgba(245,158,11,.78)")
          }]
        },
        options:{
          responsive:true,
          plugins:{
            legend:{ display:false },
            tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${fmtEURSafe(c.parsed.y)}` } }
          },
          scales:{
            x:{ ticks:{ color:"#aab6d6" }, grid:{ display:false } },
            y:{ ticks:{ color:"#aab6d6", maxTicksLimit:4, callback:(v)=>fmtAxisM(v) }, grid:{ color:"rgba(255,255,255,.08)" } }
          }
        }
      });
    }

    const evoRows = rows.map(r => ({
      name: r.name,
      vals: years.map(y => sumCodeYear(r.code, y))
    }));

    $("opDetailEvo").innerHTML = `
      <table class="tbl">
        <thead>
          <tr>
            <th>Partida</th>
            ${years.map(y=>`<th style="text-align:right;">${y}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${evoRows.map(er=>`
            <tr>
              <td>${er.name}</td>
              ${er.vals.map(v=>`<td style="text-align:right;">${fmtM(v)}</td>`).join("")}
            </tr>`).join("")}
        </tbody>
      </table>
    `;
  }

  // ------------------------
  // Resumen 4 años
  // ------------------------
  function renderSummary(contableOn){
    const years = (window.availableYears || []).slice(-4).map(Number);

    const kEB = kpi("EBITDA");
    const kBEP = kpi("BEP");

    const rows = years.map(y => {
      const wf = buildWaterfallSeries(y, contableOn);
      const map = {};
      wf.meta.forEach(m => map[m.key] = m.value);

      const sales = map.SALES ?? sumCodeYear("PYG.MAIN.1", y);
      const gross = map.GROSS ?? 0;
      const staff = map.STAFF ?? sumCodeYear("PYG.MAIN.6", y);
      const opex  = map.OPEX  ?? sumCodeYear("PYG.MAIN.7", y);

      const ebitda = kEB ? kEB.calc(y) : (map.EBITDA ?? 0);
      const bep = kBEP ? kBEP.calc(y) : null;
      const safety = (bep !== null && !Number.isNaN(bep)) ? (sales - bep) : null;

      const grossPct = (sales && sales !== 0) ? (gross / sales) : null;
      const ebitdaPct = (sales && sales !== 0) ? (ebitda / sales) : null;

      return { y, sales, gross, grossPct, staff, opex, ebitda, ebitdaPct, bep, safety };
    });

    $("opSummaryTable").innerHTML = `
      <table class="tbl">
        <thead>
          <tr>
            <th>Año</th>
            <th style="text-align:right;">Ventas</th>
            <th style="text-align:right;">Margen bruto</th>
            <th style="text-align:right;">MB %</th>
            <th style="text-align:right;">Personal</th>
            <th style="text-align:right;">Otros gastos</th>
            <th style="text-align:right;">EBITDA</th>
            <th style="text-align:right;">EBITDA %</th>
            <th style="text-align:right;">Punto equilibrio</th>
            <th style="text-align:right;">Margen seguridad</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r=>`
            <tr>
              <td>${r.y}</td>
              <td style="text-align:right;">${fmtM(r.sales)}</td>
              <td style="text-align:right;">${fmtM(r.gross)}</td>
              <td style="text-align:right;">${r.grossPct==null?"N/D":fmtPctSafe(r.grossPct)}</td>
              <td style="text-align:right;">${fmtM(r.staff)}</td>
              <td style="text-align:right;">${fmtM(r.opex)}</td>
              <td style="text-align:right;">${fmtM(r.ebitda)}</td>
              <td style="text-align:right;">${r.ebitdaPct==null?"N/D":fmtPctSafe(r.ebitdaPct)}</td>
              <td style="text-align:right;">${r.bep==null?"N/D":fmtM(r.bep)}</td>
              <td style="text-align:right;">${r.safety==null?"N/D":fmtM(r.safety)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    `;
  }

  // ------------------------
  // Charts
  // ------------------------
  let opChart1 = null;
  let opChart2 = null;
  let wfChart = null;

  function renderTopCharts(){
    const years = (window.availableYears || []).slice(-6).map(Number);
    const sales = years.map(y => sumCodeYear("PYG.MAIN.1", y));

    const kEB = kpi("EBITDA");
    const kBEP = kpi("BEP");

    const ebitda = years.map(y => kEB ? kEB.calc(y) : null);
    const bep = years.map(y => kBEP ? kBEP.calc(y) : null);

    // safety note
    const baseY = getBaseYear();
    const salesBase = sumCodeYear("PYG.MAIN.1", baseY);
    const bepBase = kBEP ? kBEP.calc(baseY) : null;
    const safety = (bepBase !== null && !Number.isNaN(bepBase)) ? (salesBase - bepBase) : null;

    $("opSafetyNote").textContent =
      (safety === null) ? "Margen de seguridad: N/D"
      : (safety >= 0 ? `Margen de seguridad: te sobra ${fmtM(safety)}` : `Margen de seguridad: te falta ${fmtM(Math.abs(safety))}`);

    // 1) Ventas vs EBITDA
    const c1 = $("opChartSalesEbitda");
    if (c1){
      if (opChart1) opChart1.destroy();
      opChart1 = new Chart(c1.getContext("2d"), {
        type:"bar",
        data:{
          labels: years,
          datasets:[
            { label:"EBITDA", data: ebitda, borderWidth:1 }, // barra
            { label:"Ventas", data: sales, type:"line", borderWidth:2, pointRadius:3, tension:.25 } // línea
          ]
        },
        options:{
          responsive:true,
          layout:{ padding:{ right: 26, top: 16 } },
          plugins:{
            legend:{ labels:{ color:"#cbd5ff" } },
            tooltip:{
              callbacks:{
                label:(c)=>{
                  const v = c.parsed.y;
                  if (c.dataset.label === "EBITDA"){
                    const s = sales[c.dataIndex];
                    const m = (s && s !== 0) ? (v/s) : null;
                    return `EBITDA: ${fmtEURSafe(v)} (${m===null?"N/D":fmtPctSafe(m)})`;
                  }
                  return `${c.dataset.label}: ${fmtEURSafe(v)}`;
                }
              }
            }
          },
          scales:{
            x:{ ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } },
            y:{ ticks:{ color:"#aab6d6", maxTicksLimit:4, callback:(v)=>fmtAxisM(v) }, grid:{ color:"rgba(255,255,255,.06)" } }
          }
        },
        plugins:[ebitdaTwoLineLabelPlugin]
      });
    }

    // 2) Ventas vs BEP
    const c2 = $("opChartSalesBep");
    if (c2){
      if (opChart2) opChart2.destroy();
      opChart2 = new Chart(c2.getContext("2d"), {
        type:"bar",
        data:{
          labels: years,
          datasets:[
            { label:"Punto de equilibrio", data: bep, borderWidth:1 },
            { label:"Ventas", data: sales, type:"line", borderWidth:2, pointRadius:3, tension:.25 }
          ]
        },
        options:{
          responsive:true,
          layout:{ padding:{ right: 26, top: 16 } },
          plugins:{
            legend:{ labels:{ color:"#cbd5ff" } },
            tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${fmtEURSafe(c.parsed.y)}` } }
          },
          scales:{
            x:{ ticks:{ color:"#aab6d6" }, grid:{ color:"rgba(255,255,255,.06)" } },
            y:{ ticks:{ color:"#aab6d6", maxTicksLimit:4, callback:(v)=>fmtAxisM(v) }, grid:{ color:"rgba(255,255,255,.06)" } }
          }
        }
      });
    }
  }

  function renderWaterfall(yearShown, contableOn){
    const wf = buildWaterfallSeries(yearShown, contableOn);
    const c = $("opChartWaterfall");
    if (!c) return;

    if (wfChart) wfChart.destroy();

    wfChart = new Chart(c.getContext("2d"), {
      type:"bar",
      data:{
        labels: wf.labels,
        datasets:[
          { label:"Offset", data: wf.offsets, backgroundColor:"rgba(0,0,0,0)", borderWidth:0, stack:"w" },
          { label:"PyG", data: wf.values, backgroundColor: wf.colors, borderWidth:1, stack:"w" }
        ]
      },
      options:{
        responsive:true,
        layout:{ padding:{ right: 26, top: 16, bottom: 6 } },
        plugins:{
          legend:{ display:false },
          tooltip:{
            callbacks:{
              label:(ctx)=>{
                const idx = ctx.dataIndex;
                const m = wf.meta[idx];
                return `${(m.label?.join ? m.label.join(" ") : m.label)}: ${fmtEURSafe(m.value)}`;
              }
            }
          }
        },
        scales:{
          x:{
            ticks:{
              color:"#aab6d6",
              maxRotation:0,
              minRotation:0,
              autoSkip:false
            },
            grid:{ display:false }
          },
          y:{
            ticks:{ color:"#aab6d6", maxTicksLimit:5, callback:(v)=>fmtAxisM(v) },
            grid:{ color:"rgba(255,255,255,.06)" }
          }
        },
        onClick:(evt, elements)=>{
          if (!elements || !elements.length) return;
          const pick = elements.find(e => e.datasetIndex === 1) || elements[0];
          const idx = pick.index;
          const metaItem = wf.meta[idx];
          renderDetail(metaItem, yearShown);
        }
      },
      plugins:[waterfallLabelPlugin, waterfallConnectorPlugin]
    });

    // guardamos data para el plugin de conectores
    wfChart.config._wfData = { offsets: wf.offsets, values: wf.values };
  }

  function renderOperativo(){
    if (!isOperativoVisible()) return;
    if (!window.availableYears || !window.availableYears.length) return;

    const baseY = getBaseYear();
    const compY = getCompYear();

    const sel = $("opYearSelect");
    if (sel && sel.options.length === 0){
      const opts = [];
      opts.push({ y: baseY, label: `Base (${baseY})` });
      if (compY !== baseY) opts.push({ y: compY, label: `Comparativa (${compY})` });
      for (const y of window.availableYears.map(Number)){
        if (y !== baseY && y !== compY) opts.push({ y, label: String(y) });
      }
      sel.innerHTML = opts.map(o=> `<option value="${o.y}">${o.label}</option>`).join("");
      sel.value = String(baseY);
    }

    const yearShown = sel ? parseInt(sel.value, 10) : baseY;
    const contableOn = !!$("opToggleContable")?.checked;

    renderTopCharts();
    renderWaterfall(yearShown, contableOn);
    renderSummary(contableOn);
  }

  // Exponemos un “render now” para initTabs()
  window.__renderOperativoNow = function(){
    try{ renderOperativo(); }catch(e){}
  };

  // listeners
  $("opYearSelect")?.addEventListener("change", ()=> renderOperativo());
  $("opToggleContable")?.addEventListener("change", ()=> renderOperativo());

  // Hook a tu render global (cuando cambian años/sector/csv)
  if (!window.__operativoRenderWrapped && typeof window.render === "function"){
    window.__operativoRenderWrapped = true;
    const _render = window.render;
    window.render = function(){
      _render();
      try{ renderOperativo(); }catch(e){}
    };
  }

  // primer intento
  try{ renderOperativo(); }catch(e){}
})();
