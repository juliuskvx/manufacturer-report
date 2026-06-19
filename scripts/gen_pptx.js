const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

const d = JSON.parse(fs.readFileSync("output/manufacturer_data.json", "utf8"));
const reportMonth = d.reportMonth;
const asOf = d.asOf;

fs.appendFileSync(process.env.GITHUB_ENV || "/dev/null", `REPORT_DATE=${d.reportDate}\n`);

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.title = `EuroAir Intel — Manufacturer Intelligence Report ${reportMonth}`;

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  navy:    "0D1F3C",
  blue:    "1B3F6E",
  mid:     "2474B5",
  gold:    "C9A84C",
  white:   "FFFFFF",
  offW:    "F8FAFC",
  slate:   "64748B",
  lgray:   "E2E8F0",
  text:    "1E293B",
  ice:     "D6E8FA",
  boeing:  "0033A0",
  airbus:  "005B8E",
  bLight:  "4B9CD3",
  aLight:  "00AEEF",
  green:   "16A34A",
  red:     "DC2626",
};

const serif = "Cambria";
const sans  = "Calibri";

// ── Helpers ───────────────────────────────────────────────────────────────────
function footer(slide) {
  slide.addShape(pptx.shapes.RECTANGLE, {x:0, y:5.38, w:10, h:0.245, fill:{color:C.navy}});
  slide.addText("Boeing.com · Airbus.com · Yahoo Finance  ·  EuroAir Intel Manufacturer Report  ·  " + reportMonth,
    {x:0.3, y:5.39, w:9.4, h:0.22, fontSize:8, color:C.ice, fontFace:sans, align:"center", valign:"middle"});
}

function header(slide, title, sub) {
  slide.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:10, h:0.78, fill:{color:C.navy}});
  slide.addShape(pptx.shapes.RECTANGLE, {x:0, y:0.78, w:10, h:0.04, fill:{color:C.gold}});
  slide.addText(title, {x:0.4, y:0.05, w:7.0, h:0.7, fontSize:21, bold:true, color:C.white, fontFace:serif, valign:"middle"});
  if (sub) slide.addText(sub, {x:7.2, y:0.05, w:2.6, h:0.7, fontSize:10, color:C.ice, fontFace:sans, valign:"middle", align:"right"});
  footer(slide);
}

function card(slide, x, y, w, h, fill, shadow) {
  slide.addShape(pptx.shapes.RECTANGLE, {x, y, w, h,
    fill:{color: fill || C.white},
    shadow: shadow !== false ? {type:"outer", blur:5, offset:2, angle:45, color:"94A3B8", opacity:0.2} : undefined
  });
}

function statCard(slide, x, y, w, label, value, unit, accentCol, dark) {
  const h = 1.45;
  card(slide, x, y, w, h, dark ? accentCol : C.white);
  slide.addShape(pptx.shapes.RECTANGLE, {x, y, w:w, h:0.05, fill:{color:accentCol}});
  slide.addText(label.toUpperCase(), {x:x+0.18, y:y+0.14, w:w-0.36, h:0.28, fontSize:8.5, color: dark ? C.ice : C.slate, fontFace:sans, charSpacing:0.8});
  slide.addText(value, {x:x+0.18, y:y+0.44, w:w-0.36, h:0.65, fontSize: String(value).length > 8 ? 22 : 28, bold:true, color: dark ? C.gold : C.text, fontFace:serif, valign:"middle"});
  if (unit) slide.addText(unit, {x:x+0.18, y:y+1.16, w:w-0.36, h:0.24, fontSize:9, color: dark ? C.aLight : C.slate, fontFace:sans});
}

function num(n) { return Number(n).toLocaleString(); }
function chg(n) { const v = Number(n); return (v >= 0 ? "▲ +" : "▼ ") + Math.abs(v).toFixed(2) + "%"; }
function chgColor(n) { return Number(n) >= 0 ? C.green : C.red; }

const B  = d.boeing;
const A  = d.airbus;
const Bs = B.stock || {};
const As = A.stock || {};

// ── SLIDE 1: Title ─────────────────────────────────────────────────────────────
const s1 = pptx.addSlide();
s1.background = {color:C.navy};

// Left panel
s1.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:5.8, h:5.625, fill:{color:C.navy}});
// Right panel — dark blue
s1.addShape(pptx.shapes.RECTANGLE, {x:5.8, y:0, w:4.2, h:5.625, fill:{color:C.blue}});
// Gold vertical divider
s1.addShape(pptx.shapes.RECTANGLE, {x:5.73, y:0, w:0.14, h:5.625, fill:{color:C.gold}});

// Tag line
s1.addText("EUROAIR INTEL", {x:0.55, y:0.55, w:5.0, h:0.35, fontSize:11, color:C.gold, fontFace:sans, charSpacing:4, bold:true});
// Main title
s1.addText("Manufacturer", {x:0.55, y:0.95, w:5.0, h:0.9,  fontSize:50, bold:true, color:C.white, fontFace:serif});
s1.addText("Intelligence", {x:0.55, y:1.82, w:5.0, h:0.9,  fontSize:50, bold:true, color:C.white, fontFace:serif});
s1.addText("Report",       {x:0.55, y:2.69, w:5.0, h:0.9,  fontSize:50, bold:true, color:C.aLight, fontFace:serif});
s1.addText(reportMonth,    {x:0.55, y:3.72, w:5.0, h:0.4,  fontSize:14, color:C.ice, fontFace:sans});
s1.addText("Boeing  vs  Airbus",  {x:0.55, y:4.22, w:5.0, h:0.38, fontSize:13, color:C.gold, fontFace:serif, italic:true});

// Right panel content
s1.addText("INSIDE THIS REPORT", {x:6.0, y:0.7, w:3.6, h:0.3, fontSize:9, color:C.gold, fontFace:sans, charSpacing:2, bold:true});
const sections = ["Head-to-Head Scorecard", "Monthly Deliveries 2026", "Deliveries by Aircraft Model", "Order Backlog by Programme", "Stock Performance (30 days)", "Orders & Book-to-Bill Intelligence", "Data Archive (Excel attached)"];
sections.forEach((s, i) => {
  s1.addText(`${i+1}.  ${s}`, {x:6.0, y:1.1+i*0.55, w:3.7, h:0.42, fontSize:10.5, color: i===6 ? C.gold : C.ice, fontFace:sans, italic: i===6});
});

// Bottom strip
s1.addShape(pptx.shapes.RECTANGLE, {x:0, y:5.18, w:10, h:0.445, fill:{color:"060E1E"}});
s1.addText(`Data as of ${asOf}  ·  Sources: Boeing.com · Airbus.com · Yahoo Finance  ·  Published 15th of each month`, {x:0.4, y:5.22, w:9.2, h:0.35, fontSize:9, color:C.slate, fontFace:sans, align:"center"});

// ── SLIDE 2: Head-to-Head Scorecard ───────────────────────────────────────────
const s2 = pptx.addSlide();
s2.background = {color:C.offW};
header(s2, "Head-to-Head Scorecard", reportMonth);

// Boeing / Airbus column headers
s2.addShape(pptx.shapes.RECTANGLE, {x:0.3,  y:0.92, w:4.3, h:0.46, fill:{color:C.boeing}});
s2.addShape(pptx.shapes.RECTANGLE, {x:5.4,  y:0.92, w:4.3, h:0.46, fill:{color:C.airbus}});
s2.addText("✈  BOEING", {x:0.3,  y:0.92, w:4.3, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:serif, align:"center", valign:"middle"});
s2.addText("✈  AIRBUS", {x:5.4,  y:0.92, w:4.3, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:serif, align:"center", valign:"middle"});
s2.addText("METRIC", {x:4.63, y:0.92, w:0.74, h:0.46, fontSize:7, color:C.slate, fontFace:sans, align:"center", valign:"middle"});

const metrics = [
  {label:"Deliveries YTD",   b:num(B.deliveries_ytd),   a:num(A.deliveries_ytd),   bw: B.deliveries_ytd > A.deliveries_ytd},
  {label:"Gross Orders YTD", b:num(B.orders_ytd),       a:num(A.orders_ytd),       bw: B.orders_ytd > A.orders_ytd},
  {label:"Total Backlog",    b:num(B.backlog)+" ac",    a:num(A.backlog)+" ac",    bw: B.backlog > A.backlog},
  {label:"Backlog Runway",   b:B.backlog_years+"y",     a:A.backlog_years+"y",     bw: B.backlog_years > A.backlog_years},
];

metrics.forEach((m, i) => {
  const y = 1.48 + i * 0.9;
  const bFill = m.bw ? "EEF3FF" : C.white;
  const aFill = !m.bw ? "E6F4FF" : C.white;

  card(s2, 0.3, y, 4.3, 0.75, bFill);
  card(s2, 5.4, y, 4.3, 0.75, aFill);

  // Gold winner bar
  if (m.bw)  s2.addShape(pptx.shapes.RECTANGLE, {x:0.3, y, w:0.07, h:0.75, fill:{color:C.gold}});
  if (!m.bw) s2.addShape(pptx.shapes.RECTANGLE, {x:9.63, y, w:0.07, h:0.75, fill:{color:C.gold}});

  s2.addText(m.b, {x:0.5, y:y+0.1, w:4.0, h:0.55, fontSize:22, bold:true, color: m.bw ? C.boeing : C.slate, fontFace:serif, valign:"middle"});
  s2.addText(m.a, {x:5.55, y:y+0.1, w:4.0, h:0.55, fontSize:22, bold:true, color: !m.bw ? C.airbus : C.slate, fontFace:serif, valign:"middle"});

  // Centre metric label
  s2.addText(m.label, {x:4.63, y:y+0.2, w:0.74, h:0.35, fontSize:7, color:C.slate, fontFace:sans, align:"center", wrap:true});
  // Winner star
  s2.addText(m.bw ? "★" : "★", {x: m.bw ? 4.4 : 4.85, y:y+0.18, w:0.22, h:0.35, fontSize:12, color:C.gold, fontFace:sans, align:"center"});
});

// ── SLIDE 3: Monthly Deliveries Chart ─────────────────────────────────────────
const s3 = pptx.addSlide();
s3.background = {color:C.white};
header(s3, "Monthly Deliveries 2026 YTD — Boeing vs Airbus", asOf);

s3.addImage({path:"output/chart_monthly.png", x:0.25, y:0.9, w:9.5, h:4.38});

// YTD totals callout bar
s3.addShape(pptx.shapes.RECTANGLE, {x:0.25, y:5.0, w:4.65, h:0.32, fill:{color:C.boeing}});
s3.addShape(pptx.shapes.RECTANGLE, {x:5.1,  y:5.0, w:4.65, h:0.32, fill:{color:C.airbus}});
s3.addText(`Boeing YTD: ${num(B.deliveries_ytd)} aircraft delivered`, {x:0.25, y:5.0, w:4.65, h:0.32, fontSize:10, bold:true, color:C.white, fontFace:sans, align:"center", valign:"middle"});
s3.addText(`Airbus YTD: ${num(A.deliveries_ytd)} aircraft delivered`, {x:5.1,  y:5.0, w:4.65, h:0.32, fontSize:10, bold:true, color:C.white, fontFace:sans, align:"center", valign:"middle"});

// ── SLIDE 4: Model Breakdown ───────────────────────────────────────────────────
const s4 = pptx.addSlide();
s4.background = {color:C.offW};
header(s4, "Deliveries by Aircraft Model — 2026 YTD", asOf);
s4.addImage({path:"output/chart_models.png", x:0.25, y:0.9, w:9.5, h:4.38});

// ── SLIDE 5: Backlog ───────────────────────────────────────────────────────────
const s5 = pptx.addSlide();
s5.background = {color:C.white};
header(s5, "Order Backlog by Programme", asOf);
s5.addImage({path:"output/chart_backlog.png", x:0.25, y:0.9, w:9.5, h:3.9});

// Backlog coverage callout cards
[[B, "Boeing", C.boeing, 0.25], [A, "Airbus", C.airbus, 5.1]].forEach(([obj, label, col, x]) => {
  card(s5, x, 4.88, 4.65, 0.44, col);
  s5.addText(`${label} backlog = ${obj.backlog_years} years of production  ·  ${num(obj.backlog)} aircraft unfilled`,
    {x:x+0.15, y:4.88, w:4.35, h:0.44, fontSize:10, bold:true, color:C.white, fontFace:sans, align:"center", valign:"middle"});
});

// ── SLIDE 6: Stock Performance ─────────────────────────────────────────────────
const s6 = pptx.addSlide();
s6.background = {color:C.white};
header(s6, "Stock Performance — BA (NYSE) vs AIR.PA (Euronext)", Bs.as_of || asOf);
s6.addImage({path:"output/chart_stock.png", x:0.25, y:0.9, w:9.5, h:3.88});

// Stock stat cards
[[Bs, "Boeing  (BA)", C.boeing, "$", 0.25], [As, "Airbus  (AIR.PA)", C.airbus, "€", 5.1]].forEach(([st, label, col, cur, x]) => {
  card(s6, x, 4.86, 4.65, 0.46, col);
  s6.addText(`${label}`, {x:x+0.18, y:4.88, w:2.0, h:0.4, fontSize:11, color:C.white, fontFace:serif, valign:"middle", bold:true});
  s6.addText(`${cur}${st.price}`, {x:x+2.1, y:4.88, w:1.3, h:0.4, fontSize:14, bold:true, color:C.gold, fontFace:serif, valign:"middle"});
  const pos = Number(st.change_pct) >= 0;
  s6.addText(chg(st.change_pct), {x:x+3.35, y:4.88, w:1.15, h:0.4, fontSize:11, bold:true, color: pos ? "90EE90" : "FFB3B3", fontFace:sans, valign:"middle", align:"right"});
});

// 52-week range bars
[[Bs, C.boeing, "$", 0.25], [As, C.airbus, "€", 5.1]].forEach(([st, col, cur, x]) => {
  if (!st.week52_low || !st.week52_high) return;
  // Small 52w range indicator above the cards
  s6.addText(`52w: ${cur}${st.week52_low} — ${cur}${st.week52_high}`, {x:x+0.18, y:4.72, w:4.3, h:0.16, fontSize:7.5, color:C.slate, fontFace:sans});
});

// ── SLIDE 7: Orders Intelligence ──────────────────────────────────────────────
const s7 = pptx.addSlide();
s7.background = {color:C.navy};
header(s7, "Orders & Book-to-Bill Intelligence", reportMonth);

[[B, "BOEING", C.boeing, C.bLight, 0.25], [A, "AIRBUS", C.airbus, C.aLight, 5.1]].forEach(([obj, label, col, accent, x]) => {
  card(s7, x, 0.92, 4.65, 4.38, col, false);
  // Header stripe
  s7.addShape(pptx.shapes.RECTANGLE, {x, y:0.92, w:4.65, h:0.52, fill:{color: col === C.boeing ? "001F6B" : "003D5C"}});
  s7.addText(`✈  ${label}`, {x:x+0.2, y:0.92, w:4.25, h:0.52, fontSize:16, bold:true, color:C.white, fontFace:serif, valign:"middle"});

  const btb = (obj.orders_ytd / Math.max(obj.deliveries_ytd, 1)).toFixed(2);
  const items = [
    ["Gross Orders YTD",    num(obj.orders_ytd),       "firm orders booked"],
    ["Deliveries YTD",      num(obj.deliveries_ytd),   "aircraft handed over"],
    ["Book-to-Bill Ratio",  btb + "x",                 btb >= 1 ? "orders exceed deliveries ✓" : "deliveries exceed orders"],
    ["Total Backlog",       num(obj.backlog),           "aircraft on order"],
    ["Backlog Runway",      obj.backlog_years + " yrs", "at current delivery pace"],
  ];

  items.forEach(([lbl, val, sub], i) => {
    const y = 1.57 + i * 0.73;
    s7.addText(lbl, {x:x+0.22, y:y, w:4.2, h:0.24, fontSize:9, color:C.ice, fontFace:sans});
    s7.addText(val, {x:x+0.22, y:y+0.24, w:4.2, h:0.35, fontSize:18, bold:true, color:C.gold, fontFace:serif});
    s7.addText(sub, {x:x+0.22, y:y+0.59, w:4.2, h:0.16, fontSize:7.5, color:accent, fontFace:sans, italic:true});
  });
});

// ── SLIDE 8: Key Metrics Summary Cards ────────────────────────────────────────
const s8 = pptx.addSlide();
s8.background = {color:C.offW};
header(s8, "At a Glance — Key Metrics", reportMonth);

// 8 stat cards in 2 rows of 4
const cards8 = [
  {label:"Boeing Deliveries YTD", value:num(B.deliveries_ytd), unit:"aircraft", col:C.boeing},
  {label:"Airbus Deliveries YTD", value:num(A.deliveries_ytd), unit:"aircraft", col:C.airbus},
  {label:"Boeing Orders YTD",     value:num(B.orders_ytd),     unit:"gross orders", col:C.boeing},
  {label:"Airbus Orders YTD",     value:num(A.orders_ytd),     unit:"gross orders", col:C.airbus},
  {label:"Boeing Backlog",        value:num(B.backlog),         unit:"aircraft unfilled", col:C.boeing},
  {label:"Airbus Backlog",        value:num(A.backlog),         unit:"aircraft unfilled", col:C.airbus},
  {label:"BA Stock",              value:`$${Bs.price}`,         unit:chg(Bs.change_pct)+" today", col: Number(Bs.change_pct)>=0 ? C.green : C.red},
  {label:"AIR.PA Stock",          value:`€${As.price}`,         unit:chg(As.change_pct)+" today", col: Number(As.change_pct)>=0 ? C.green : C.red},
];

cards8.forEach((c, i) => {
  const col = i % 4;
  const row = Math.floor(i / 4);
  const x = 0.25 + col * 2.42;
  const y = 0.98 + row * 1.65;
  statCard(s8, x, y, 2.28, c.label, c.value, c.unit, c.col, false);
});

// ── SLIDE 9: Back Cover ────────────────────────────────────────────────────────
const s9 = pptx.addSlide();
s9.background = {color:C.navy};

s9.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:4.0, h:5.625, fill:{color:C.blue}});
s9.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:0.1,  h:5.625, fill:{color:C.gold}});

// Left: summary stats
s9.addText("Summary", {x:0.25, y:0.55, w:3.5, h:0.65, fontSize:28, bold:true, color:C.white, fontFace:serif});
s9.addText(reportMonth, {x:0.25, y:1.22, w:3.5, h:0.3,  fontSize:10, color:C.ice, fontFace:sans});

const summaryRows = [
  ["Boeing deliveries YTD",  num(B.deliveries_ytd)],
  ["Airbus deliveries YTD",  num(A.deliveries_ytd)],
  ["Boeing orders YTD",      num(B.orders_ytd)],
  ["Airbus orders YTD",      num(A.orders_ytd)],
  ["Boeing backlog",         num(B.backlog)+" ac"],
  ["Airbus backlog",         num(A.backlog)+" ac"],
  [`BA  (${Bs.as_of||"—"})`, `$${Bs.price}`],
  [`AIR.PA (${As.as_of||"—"})`, `€${As.price}`],
];

summaryRows.forEach(([lbl, val], i) => {
  const y = 1.65 + i * 0.47;
  s9.addText(lbl, {x:0.25, y, w:2.05, h:0.32, fontSize:9, color:C.ice, fontFace:sans});
  s9.addText(val, {x:2.32, y, w:1.55, h:0.32, fontSize:9, bold:true, color:C.gold, fontFace:sans, align:"right"});
});

// Right: closing
s9.addText("Boeing\nvs\nAirbus", {x:4.3, y:0.9, w:5.3, h:2.2, fontSize:48, bold:true, color:C.white, fontFace:serif, lineSpacingMultiple:1.2, align:"center"});
s9.addText("Monthly intelligence briefing on commercial aircraft orders, deliveries, backlog, and stock performance. Data from official manufacturer releases and Yahoo Finance.", {x:4.3, y:3.3, w:5.3, h:1.0, fontSize:11, color:C.ice, fontFace:sans, lineSpacingMultiple:1.35, align:"center"});
s9.addShape(pptx.shapes.RECTANGLE, {x:4.3, y:4.45, w:5.3, h:0.05, fill:{color:C.gold}});
s9.addText("EuroAir Intel  ·  Published monthly on the 15th  ·  aviotiongeek@gmail.com", {x:4.3, y:4.6, w:5.3, h:0.3, fontSize:9.5, color:C.slate, fontFace:sans, align:"center"});

// ── Write ──────────────────────────────────────────────────────────────────────
fs.mkdirSync("output", {recursive:true});
const fname = `output/Manufacturer_Intelligence_Report_${d.reportDate}.pptx`;
pptx.writeFile({fileName: fname})
  .then(() => console.log("DONE: " + fname))
  .catch(e => { console.error("ERROR: " + e.message); process.exit(1); });
