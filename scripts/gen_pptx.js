const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

const d = JSON.parse(fs.readFileSync("output/manufacturer_data.json", "utf8"));
const reportMonth = d.reportMonth;
const asOf = d.asOf;

fs.appendFileSync(process.env.GITHUB_ENV || "/dev/null", `REPORT_DATE=${d.reportDate}\n`);

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.title = `EuroAir Intel — Manufacturer Intelligence Report ${reportMonth}`;

// ── Palette ──────────────────────────────────────────────────────────────────
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
  boeing:  "0033A0",   // Boeing brand blue
  airbus:  "00AEEF",   // Airbus brand cyan
  boeingL: "4B9CD3",
  airbusD: "00205B",
};

const serif = "Cambria";
const sans  = "Calibri";

// ── Helpers ──────────────────────────────────────────────────────────────────
function header(slide, title, sub) {
  slide.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:10, h:0.8, fill:{color:C.navy}});
  slide.addText(title, {x:0.4, y:0.05, w:7.2, h:0.7, fontSize:20, bold:true, color:C.white, fontFace:serif, valign:"middle"});
  if (sub) slide.addText(sub, {x:7.3, y:0.05, w:2.5, h:0.7, fontSize:10, color:C.ice, fontFace:sans, valign:"middle", align:"right"});
  slide.addText("Boeing & Airbus Official Data · EuroAir Intel", {x:0, y:5.35, w:10, h:0.28, fontSize:9, color:C.slate, fontFace:sans, align:"center"});
}

function card(slide, x, y, w, h, col) {
  slide.addShape(pptx.shapes.RECTANGLE, {x, y, w, h, fill:{color:col||C.white},
    shadow:{type:"outer", blur:3, offset:2, angle:45, color:"CBD5E1", opacity:0.25}});
}

function num(n) { return Number(n).toLocaleString(); }
function pct(n) {
  const v = Number(n);
  return (v >= 0 ? "+" : "") + v.toFixed(2) + "%";
}

// ── Data shortcuts ────────────────────────────────────────────────────────────
const B  = d.boeing;
const A  = d.airbus;
const Bs = B.stock || {};
const As = A.stock || {};

// ── SLIDE 1: Title ────────────────────────────────────────────────────────────
const s1 = pptx.addSlide();
s1.background = {color:C.navy};

// Split background — Boeing left, Airbus right
s1.addShape(pptx.shapes.RECTANGLE, {x:5.0, y:0, w:5.0, h:5.625, fill:{color:C.airbusD}});
s1.addShape(pptx.shapes.RECTANGLE, {x:4.92, y:0, w:0.16, h:5.625, fill:{color:C.gold}});

// Boeing side
s1.addText("Boeing", {x:0.55, y:0.7, w:4.2, h:1.0, fontSize:46, bold:true, color:C.white, fontFace:serif});
s1.addText("vs", {x:0.55, y:1.65, w:4.2, h:0.6, fontSize:22, color:C.gold, fontFace:serif, italic:true});
s1.addText("Airbus", {x:0.55, y:2.2, w:4.2, h:1.0, fontSize:46, bold:true, color:C.airbus, fontFace:serif});
s1.addText("Manufacturer\nIntelligence Report", {x:0.55, y:3.3, w:4.2, h:0.9, fontSize:13, color:C.ice, fontFace:sans, lineSpacingMultiple:1.3});
s1.addText(reportMonth, {x:0.55, y:4.3, w:4.2, h:0.4, fontSize:11, color:C.slate, fontFace:sans});

// Airbus side
s1.addText("Orders · Deliveries\nBacklog · Stock", {x:5.2, y:1.6, w:4.5, h:1.5, fontSize:22, bold:true, color:C.white, fontFace:serif, align:"center", lineSpacingMultiple:1.5});
s1.addText("Monthly intelligence briefing\nEuroAir Intel", {x:5.2, y:4.0, w:4.5, h:0.8, fontSize:11, color:C.ice, fontFace:sans, align:"center", lineSpacingMultiple:1.4});

// Bottom bar
s1.addShape(pptx.shapes.RECTANGLE, {x:0, y:5.15, w:10, h:0.475, fill:{color:C.blue}});
s1.addText(`DATA AS OF ${asOf.toUpperCase()} · SOURCES: BOEING.COM · AIRBUS.COM · YAHOO FINANCE`, {x:0.4, y:5.2, w:9, h:0.35, fontSize:9, color:C.ice, fontFace:sans, charSpacing:1.5});

// ── SLIDE 2: Scorecard ────────────────────────────────────────────────────────
const s2 = pptx.addSlide();
s2.background = {color:C.offW};
header(s2, "Head-to-Head Scorecard — " + reportMonth, asOf);

// Column headers
s2.addShape(pptx.shapes.RECTANGLE, {x:0.3,  y:0.88, w:4.45, h:0.38, fill:{color:C.boeing}});
s2.addShape(pptx.shapes.RECTANGLE, {x:5.25, y:0.88, w:4.45, h:0.38, fill:{color:C.airbusD}});
s2.addText("BOEING", {x:0.3,  y:0.88, w:4.45, h:0.38, fontSize:13, bold:true, color:C.white, fontFace:serif, align:"center", valign:"middle"});
s2.addText("AIRBUS", {x:5.25, y:0.88, w:4.45, h:0.38, fontSize:13, bold:true, color:C.white, fontFace:serif, align:"center", valign:"middle"});

const rows = [
  {label:"Deliveries YTD",     b:num(B.deliveries_ytd),   a:num(A.deliveries_ytd),   winner: B.deliveries_ytd > A.deliveries_ytd ? "b" : "a"},
  {label:"Gross Orders YTD",   b:num(B.orders_ytd),       a:num(A.orders_ytd),       winner: B.orders_ytd > A.orders_ytd ? "b" : "a"},
  {label:"Total Backlog",      b:num(B.backlog)+" ac",    a:num(A.backlog)+" ac",    winner: B.backlog > A.backlog ? "b" : "a"},
  {label:"Backlog Coverage",   b:B.backlog_years+"y",     a:A.backlog_years+"y",     winner: A.backlog_years > B.backlog_years ? "a" : "b"},
];

rows.forEach((row, i) => {
  const y = 1.36 + i * 0.88;
  // Boeing cell
  card(s2, 0.3, y, 4.45, 0.75, row.winner==="b" ? "EEF4FF" : C.white);
  if (row.winner==="b") s2.addShape(pptx.shapes.RECTANGLE, {x:0.3, y, w:0.06, h:0.75, fill:{color:C.gold}});
  s2.addText(row.label, {x:0.5, y:y+0.04, w:4.1, h:0.3, fontSize:9.5, color:C.slate, fontFace:sans});
  s2.addText(row.b, {x:0.5, y:y+0.34, w:4.1, h:0.36, fontSize:20, bold:true, color:C.boeing, fontFace:serif});
  // Airbus cell
  card(s2, 5.25, y, 4.45, 0.75, row.winner==="a" ? "E6F7FF" : C.white);
  if (row.winner==="a") s2.addShape(pptx.shapes.RECTANGLE, {x:9.64, y, w:0.06, h:0.75, fill:{color:C.gold}});
  s2.addText(row.label, {x:5.35, y:y+0.04, w:4.1, h:0.3, fontSize:9.5, color:C.slate, fontFace:sans});
  s2.addText(row.a, {x:5.35, y:y+0.34, w:4.1, h:0.36, fontSize:20, bold:true, color:C.airbusD, fontFace:serif});
  // Centre label
  s2.addText(row.label.toUpperCase(), {x:4.76, y:y+0.22, w:0.48, h:0.3, fontSize:6.5, color:C.slate, fontFace:sans, align:"center", wrap:true});
});

// ── SLIDE 3: Monthly Deliveries Chart ─────────────────────────────────────────
const s3 = pptx.addSlide();
s3.background = {color:C.white};
header(s3, "Monthly Deliveries 2026 YTD — Boeing vs Airbus", asOf);
s3.addImage({path:"output/chart_monthly.png", x:0.2, y:0.88, w:9.6, h:4.42});

// ── SLIDE 4: Model Breakdown ──────────────────────────────────────────────────
const s4 = pptx.addSlide();
s4.background = {color:C.offW};
header(s4, "Deliveries by Aircraft Model — 2026 YTD", asOf);
s4.addImage({path:"output/chart_models.png", x:0.2, y:0.88, w:9.6, h:4.42});

// ── SLIDE 5: Backlog Deep Dive ────────────────────────────────────────────────
const s5 = pptx.addSlide();
s5.background = {color:C.white};
header(s5, "Order Backlog by Programme", asOf);
s5.addImage({path:"output/chart_backlog.png", x:0.2, y:0.88, w:9.6, h:4.0});

// Backlog coverage callout
s5.addShape(pptx.shapes.RECTANGLE, {x:0.3, y:4.95, w:4.55, h:0.45, fill:{color:C.boeing}});
s5.addText(`Boeing backlog = ${B.backlog_years} years of production`, {x:0.3, y:4.95, w:4.55, h:0.45, fontSize:10, bold:true, color:C.white, fontFace:sans, align:"center", valign:"middle"});
s5.addShape(pptx.shapes.RECTANGLE, {x:5.15, y:4.95, w:4.55, h:0.45, fill:{color:C.airbusD}});
s5.addText(`Airbus backlog = ${A.backlog_years} years of production`, {x:5.15, y:4.95, w:4.55, h:0.45, fontSize:10, bold:true, color:C.white, fontFace:sans, align:"center", valign:"middle"});

// ── SLIDE 6: Stock Performance Chart ─────────────────────────────────────────
const s6 = pptx.addSlide();
s6.background = {color:C.white};
header(s6, "Stock Performance — BA (NYSE) vs AIR.PA (Euronext)", Bs.as_of || asOf);
s6.addImage({path:"output/chart_stock.png", x:0.2, y:0.88, w:9.6, h:4.1});

// Stock cards below chart
[[Bs, "Boeing (BA)", C.boeing, 0.3], [As, "Airbus (AIR.PA)", C.airbusD, 5.2]].forEach(([st, label, col, x]) => {
  const changePos = Number(st.change_pct) >= 0;
  card(s6, x, 5.02, 4.5, 0.5, C.white);
  s6.addShape(pptx.shapes.RECTANGLE, {x, y:5.02, w:0.06, h:0.5, fill:{color:col}});
  const currency = st.currency === "EUR" ? "€" : "$";
  s6.addText(`${label}:  ${currency}${st.price}  `, {x:x+0.15, y:5.07, w:2.5, h:0.38, fontSize:13, bold:true, color:col, fontFace:serif, valign:"middle"});
  s6.addText(pct(st.change_pct), {x:x+2.65, y:5.07, w:1.7, h:0.38, fontSize:13, bold:true, color:changePos?"008000":"CC0000", fontFace:serif, valign:"middle"});
});

// ── SLIDE 7: Orders Intelligence ─────────────────────────────────────────────
const s7 = pptx.addSlide();
s7.background = {color:C.navy};
header(s7, "Orders Intelligence — " + reportMonth, asOf);

// Two big panels
[[B, "BOEING", C.boeing, C.boeingL, 0.3], [A, "AIRBUS", C.airbusD, C.airbus, 5.2]].forEach(([obj, label, col, accent, x]) => {
  card(s7, x, 0.95, 4.5, 4.45, col);
  s7.addText(label, {x:x+0.2, y:1.05, w:4.1, h:0.5, fontSize:18, bold:true, color:C.white, fontFace:serif});
  const items = [
    ["YTD Orders",    num(obj.orders_ytd)],
    ["YTD Deliveries",num(obj.deliveries_ytd)],
    ["Total Backlog", num(obj.backlog)+" aircraft"],
    ["Book-to-Bill",  (obj.orders_ytd / Math.max(obj.deliveries_ytd,1)).toFixed(2)+"x"],
    ["Backlog Runway",obj.backlog_years+" years"],
  ];
  items.forEach(([lbl, val], i) => {
    const y = 1.65 + i * 0.7;
    s7.addText(lbl, {x:x+0.25, y:y+0.02, w:4.0, h:0.24, fontSize:9, color:C.ice, fontFace:sans});
    s7.addText(val, {x:x+0.25, y:y+0.26, w:4.0, h:0.38, fontSize:17, bold:true, color:C.gold, fontFace:serif});
  });
});

// ── SLIDE 8: Summary ──────────────────────────────────────────────────────────
const s8 = pptx.addSlide();
s8.background = {color:C.navy};

s8.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:4.2, h:5.625, fill:{color:C.blue}});
s8.addShape(pptx.shapes.RECTANGLE, {x:0, y:0, w:0.08, h:5.625, fill:{color:C.gold}});

s8.addText("Summary", {x:0.25, y:0.55, w:3.7, h:0.65, fontSize:26, bold:true, color:C.white, fontFace:serif});
s8.addText(reportMonth, {x:0.25, y:1.22, w:3.7, h:0.3, fontSize:10, color:C.ice, fontFace:sans});

const summaryRows = [
  ["Boeing deliveries YTD",  num(B.deliveries_ytd)],
  ["Airbus deliveries YTD",  num(A.deliveries_ytd)],
  ["Boeing orders YTD",      num(B.orders_ytd)],
  ["Airbus orders YTD",      num(A.orders_ytd)],
  ["Boeing backlog",         num(B.backlog)+" ac"],
  ["Airbus backlog",         num(A.backlog)+" ac"],
  [`BA stock (${Bs.as_of||"—"})`, `$${Bs.price} (${pct(Bs.change_pct)})`],
  [`AIR.PA (${As.as_of||"—"})`,   `€${As.price} (${pct(As.change_pct)})`],
];

summaryRows.forEach(([lbl, val], i) => {
  const y = 1.65 + i * 0.47;
  s8.addText(lbl, {x:0.25, y, w:2.0, h:0.3, fontSize:9.5, color:C.ice, fontFace:sans});
  s8.addText(val, {x:2.3,  y, w:1.85, h:0.3, fontSize:9.5, bold:true, color:C.gold, fontFace:sans});
});

// Right: big closing statement
s8.addText("Manufacturer\nIntelligence", {x:4.5, y:1.1, w:5.2, h:1.8, fontSize:34, bold:true, color:C.white, fontFace:serif, lineSpacingMultiple:1.3});
s8.addText("Monthly briefing on Boeing and Airbus commercial aircraft orders, deliveries, backlog, and stock performance. Data sourced from official manufacturer reports and Yahoo Finance.", {x:4.5, y:3.2, w:5.2, h:1.1, fontSize:11, color:C.ice, fontFace:sans, lineSpacingMultiple:1.35});
s8.addShape(pptx.shapes.RECTANGLE, {x:4.5, y:4.45, w:5.2, h:0.06, fill:{color:C.gold}});
s8.addText("EuroAir Intel · Published monthly on the 15th", {x:4.5, y:4.6, w:5.2, h:0.3, fontSize:10, color:C.slate, fontFace:sans});

// ── Write ─────────────────────────────────────────────────────────────────────
fs.mkdirSync("output", {recursive:true});
const fname = `output/Manufacturer_Intelligence_Report_${d.reportDate}.pptx`;
pptx.writeFile({fileName: fname})
  .then(() => console.log("DONE: " + fname))
  .catch(e => { console.error("ERROR: " + e.message); process.exit(1); });
