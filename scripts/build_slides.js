/* 報告スライド（v4）を生成する。文章の控え docs/slides_text.md は、生成後に scripts/dump_slide_text.py で PPTX から書き出す。
 * 数値は reports/slide_data.json（scripts/export_slide_data.py が作る）からだけ読む。
 * デザインの決まり: 白背景（表紙と結論だけ紺）・Arial・色は紺/青/橙の 3 色まで・タイトルは言いたいことを 1 文で。
 * v4 では文字を減らすため、タイトルを 1 行にし、グラフのページの説明は 1〜2 項目に絞って注目点を図に直接書き、
 * 背景・評価のやり方・設計の判断・限界・結論・付録 D を図か表に置き換えた。要旨・問い・安定性の 3 枚は v3 の形のまま。
 * 使い方: uv run python scripts/export_slide_data.py
 *        NODE_PATH=<pptxgenjs のある node_modules> node scripts/build_slides.js
 *        uv run --with python-pptx python scripts/dump_slide_text.py
 * 出力:   reports/slides/ett_oil_temperature_poc.pptx
 */
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const OUT_DIR = path.join(ROOT, "reports", "slides");
fs.mkdirSync(OUT_DIR, { recursive: true });
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "slide_data.json"), "utf8"));
const WEEK = D.week_median_ETTh2_h6;

const NAVY = "1F4E79", BLUE = "2E75B6", ORANGE = "C55A11";
const INK = "262626", BODY = "3A3A3A", MUTED = "6E6E6E", GRAY = "A6A6A6", LIGHT = "D9D9D9", WHITE = "FFFFFF", TINT = "D9D9D9";
const FONT = "Arial";
const W = 10, H = 5.625, M = 0.5, CW = W - 2 * M;

const f2 = (v) => Number(v).toFixed(2);
const f1 = (v) => Number(v).toFixed(1);
const cut = (a, b) => Math.round((1 - b / a) * 100);
const mae = D.mae, raw = D.mae_raw, det = D.detection_ETTh2_h6, thr2 = D.thresholds.ETTh2.q95;
const lg = det.lgbm;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.author = "Kutsunugi Hijiri";
pres.title = "変圧器油温の先読み PoC";
let page = 0;

function slide({ dark = false } = {}) {
  const s = pres.addSlide();
  s.background = { color: dark ? NAVY : WHITE };
  page += 1;
  s.addText(String(page), { x: W - M - 0.6, y: H - 0.4, w: 0.6, h: 0.24, fontFace: FONT, fontSize: 10, color: dark ? TINT : MUTED, align: "right", valign: "bottom", margin: 0, isTextBox: true });
  return s;
}
function text(s, t, x, y, w, h, { size = 14, color = BODY, bold = false, align = "left", valign = "top" } = {}) {
  s.addText(t, { x, y, w, h, fontFace: FONT, fontSize: size, color, bold, align, valign, margin: 0, isTextBox: true, lineSpacingMultiple: 1.08 });
}
function title(s, t, { color = NAVY, size = 24, y = 0.36 } = {}) {
  text(s, t, M, y, CW, 0.55, { size, color, bold: true });
}
function appendixLabel(s, t) {
  text(s, t, M, 0.14, CW, 0.26, { size: 11, color: MUTED, bold: true });
}
function note(s, t, { dark = false } = {}) {
  s.addText(t, { x: M, y: H - 0.62, w: CW - 0.8, h: 0.4, fontFace: FONT, fontSize: 11, color: dark ? TINT : MUTED, valign: "bottom", margin: 0, isTextBox: true, lineSpacingMultiple: 1.05 });
}
function hline(s, x, y, w, color = LIGHT, pt = 0.75) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: pt } });
}
function bullets(s, items, { x = M, y = 1.3, w = CW, size = 16, gap = 0.22, color = BODY, emphasis = [0] } = {}) {
  const lineHeight = size * 1.24 / 72;
  let top = y;
  items.forEach((item, index) => {
    const lines = String(item).split("\n").reduce((total, part) => {
      const units = Array.from(part).reduce((sum, ch) => sum + (/[^\x00-\x7F]/.test(ch) ? 1 : 0.56), 0);
      return total + Math.max(1, Math.ceil(units * size / (w * 72)));
    }, 0);
    const h = lines * lineHeight + 0.05;
    if (index > 0) hline(s, x, top - gap / 2, w, LIGHT, 0.5);
    text(s, String(item), x, top, w, h, { size, color: emphasis.includes(index) ? NAVY : color, bold: emphasis.includes(index) });
    top += h + gap;
  });
}
function table(s, rows, { x = M, y = 1.3, w = CW, colW, size = 16, head = true, rowH = 0.5, bold = [], cellStyle } = {}) {
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const isHead = head && ri === 0;
    const extra = cellStyle ? cellStyle(ri, ci) || {} : {};
    return {
      text: String(c),
      options: {
        fontFace: FONT, fontSize: isHead ? Math.max(12, size - 2) : size, color: isHead ? NAVY : (bold.includes(ci) ? NAVY : BODY), bold: isHead || bold.includes(ci),
        align: "left", valign: "middle", margin: [7, 8, 7, 6], fill: isHead ? "F2F2F2" : WHITE,
        border: [{ type: "none" }, { type: "none" }, { pt: isHead ? 1 : 0.5, color: isHead ? NAVY : LIGHT }, { type: "none" }],
        ...extra,
      },
    };
  }));
  s.addTable(data, { x, y, w, colW, rowH, autoPage: false });
  hline(s, x, y, w, head ? NAVY : GRAY, head ? 1 : 0.75);
}

// ---------- グラフ
function plot(ch) { return { x: ch.x + ch.layout.x * ch.w, y: ch.y + ch.layout.y * ch.h, w: ch.layout.w * ch.w, h: ch.layout.h * ch.h }; }
function yOf(ch, v) { const p = plot(ch); return p.y + p.h * (1 - (v - ch.min) / (ch.max - ch.min)); }
function xOf(ch, i) { const p = plot(ch); return p.x + p.w * (i + 0.5) / ch.labels.length; }
function barChart(s, ch) {
  ch.min = ch.min || 0;
  s.addChart(pres.charts.BAR, [{ name: "value", labels: ch.labels, values: ch.values }], {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, barDir: "col", chartColors: ch.colors, barGapWidthPct: ch.gap || 110,
    valAxisMinVal: ch.min, valAxisMaxVal: ch.max, valAxisHidden: true, valGridLine: { style: "none" },
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: ch.catSize || 13, catAxisLabelColor: BODY, catAxisLineShow: true, catAxisLineColor: GRAY,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FONT, dataLabelFontSize: ch.valSize || 16, dataLabelFontBold: true, dataLabelColor: INK, dataLabelFormatCode: ch.fmt || "0.00",
    showLegend: false, layout: ch.layout,
  });
  return ch;
}
function lineChart(s, ch) {
  s.addChart(pres.charts.LINE, ch.series.map((se) => ({ name: se.name, labels: ch.labels, values: se.values })), {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, chartColors: ch.series.map((se) => se.color), lineSize: ch.lineSize || 2.25, lineDataSymbol: ch.symbol || "none", lineDataSymbolSize: 7,
    valAxisMinVal: ch.min, valAxisMaxVal: ch.max, valAxisMajorUnit: ch.step, valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12, valAxisLabelColor: MUTED, valAxisLabelFormatCode: ch.fmt || "0", valAxisLineShow: false,
    valAxisHidden: !!ch.hideVal,
    valGridLine: ch.hideGrid ? { style: "none" } : { color: "EBEBEB", size: 0.75 }, catGridLine: { style: "none" },
    catAxisLabelPos: "none", catAxisMajorTickMark: "none", catAxisLineColor: GRAY,
    showLegend: false, layout: ch.layout,
  });
  if (!ch.noXLabels) {
    const p = plot(ch), step = ch.freq || 1;
    ch.labels.forEach((lb, i) => {
      if (i % step) return;
      text(s, ch.xLabel ? ch.xLabel(lb, i) : lb, xOf(ch, i) - 0.4, p.y + p.h + 0.05, 0.8, 0.22, { size: ch.xSize || 12, color: MUTED, align: "center" });
    });
  }
  return ch;
}
function endLabels(s, ch, items, { w = 1.5, size = 13, minGap = 0.26 } = {}) {
  const p = plot(ch);
  const pos = items.map((it) => ({ ...it, y: yOf(ch, it.value) })).sort((a, b) => a.y - b.y);
  for (let i = 1; i < pos.length; i++) if (pos[i].y - pos[i - 1].y < minGap) pos[i].y = pos[i - 1].y + minGap;
  pos.forEach((it) => text(s, it.text, p.x + p.w + 0.08, it.y - 0.13, w, 0.26, { size, color: it.color, bold: !!it.bold, valign: "middle" }));
}

// ---------- 2 行のタイトルと、枠の高さに合わせて間隔を決める箇条書き（要旨・問い・安定性で使う。v3 と同じ）
function titleTwoLines(s, t, { color = NAVY, size = 24, y = 0.34, h = 0.95 } = {}) {
  s.addText(t, { x: M, y, w: CW, h, fontFace: FONT, fontSize: size, bold: true, color, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.08 });
}
function bulletsFit(s, items, { x = M, y = 1.5, w = CW, h = 3.3, size = 20, gap = 12, color = BODY, emphasis = [0] } = {}) {
  const lineHeight = size * 1.24 / 72;
  const heights = items.map((item) => String(item).split("\n").reduce((total, part) => {
    const units = Array.from(part).reduce((sum, char) => sum + (/[^\x00-\x7F]/.test(char) ? 1 : 0.56), 0);
    return total + Math.max(1, Math.ceil(units * size / (w * 72)));
  }, 0) * lineHeight + 0.05);
  const spacing = Math.max(0.06, Math.min(gap / 72 + 0.08, (Math.min(h, 4.72 - y) - heights.reduce((sum, height) => sum + height, 0)) / Math.max(1, items.length - 1)));
  let top = y;
  items.forEach((item, index) => {
    if (index > 0) hline(s, x, top - spacing / 2, w, LIGHT, 0.5);
    text(s, String(item), x, top, w, heights[index], { size, color: emphasis.includes(index) ? NAVY : color, bold: emphasis.includes(index) });
    top += heights[index] + spacing;
  });
}
function chartLabel(s, t, x, y, w = 3.5) { text(s, t, x, y, w, 0.22, { size: 11, color: MUTED }); }

// ========== 1. 表紙
{
  const s = slide({ dark: true });
  text(s, "変圧器の油温を6時間先まで予測し、\n高温になる前に知らせられるか", M, 1.3, 8.8, 1.5, { size: 30, color: WHITE, bold: true });
  text(s, "ETT公開データを使った油温予測の検証報告", M, 2.95, 8.8, 0.4, { size: 16, color: TINT });
  hline(s, M, 4.25, CW, TINT, 0.75);
  text(s, "沓脱 聖　｜　2026年9月", M, 4.48, 8.8, 0.35, { size: 15, color: TINT });
}

// ========== 2. 要旨
{
  const s = slide();
  const d = det.lgbm, r6 = raw.ETTh2["6"];
  titleTwoLines(s, `設備2では6時間先の誤差を${cut(r6.ridge, r6.lgbm)}%減らし、\n高温${d.events}区間中${d.detected}区間を事前に検知した`);
  table(s, [
    ["精度", `6時間先のMAEは${f2(mae.ETTh2["6"].lgbm)}℃（最良の比較手法は${f2(mae.ETTh2["6"].ridge)}℃）`],
    ["事前検知", `平均${f1(d.lead_h)}時間前に検知。通知${d.notices}件のうち誤通知は${d.false_notices}件`],
    ["限界", "24時間先と設備1は改善が小さく、短い高温は見逃す"],
  ], { y: 1.55, colW: [1.7, 7.3], size: 17, rowH: 0.6, head: false, bold: [0] });
  text(s, "提案：", M + 0.08, 3.72, 1.5, 0.4, { size: 17, color: NAVY, bold: true });
  text(s, "現場へ通知せずに試行し、\n通知候補の数と外れた数を確かめる", M + 1.78, 3.68, CW - 1.78, 0.85, { size: 20, color: NAVY, bold: true });
  note(s, "設備2（ETTh2）・テスト期間 2017年11月〜2018年6月・主な評価（予測時点までの観測だけを使う）の条件。区間と通知は別の数え方（定義は付録D）");
}

// ========== 3. 背景 — 「超えてから気づく／超える前に知らせる」のイメージ図（実データではない）
{
  const s = slide();
  title(s, "超えてから気づく監視を、超える前に知らせる監視へ");
  // イメージ図（実データではない）: 滑らかに上がって閾値を超える曲線
  const n = 24, thr = 46;
  const curve = Array.from({ length: n }, (_, i) => 36 + 13 * Math.exp(-Math.pow((i - 15) / 4.2, 2)));
  const ch = lineChart(s, {
    x: M, y: 1.2, w: 5.9, h: 3.45, labels: curve.map((_, i) => String(i)), min: 30, max: 52, step: 5, hideVal: true, hideGrid: true, noXLabels: true,
    series: [{ name: "油温", values: curve, color: INK }, { name: "閾値", values: curve.map(() => thr), color: ORANGE }],
    layout: { x: 0.02, y: 0.06, w: 0.8, h: 0.8 }, lineSize: 2.5,
  });
  const p = plot(ch);
  const cross = curve.findIndex((v) => v > thr);           // 閾値を超える最初の時刻
  const early = cross - 4;                                  // 予測で分かる時刻（イメージ）
  text(s, "閾値", p.x + p.w + 0.08, yOf(ch, thr) - 0.13, 0.8, 0.26, { size: 13, color: ORANGE, bold: true, valign: "middle" });
  text(s, "油温", p.x + p.w + 0.08, yOf(ch, curve[n - 1]) - 0.13, 0.8, 0.26, { size: 13, color: INK, bold: true, valign: "middle" });
  // 閾値監視: 超えた時点で気づく（点の左上、閾値の上に置く）
  s.addShape(pres.shapes.OVAL, { x: xOf(ch, cross) - 0.09, y: yOf(ch, curve[cross]) - 0.09, w: 0.18, h: 0.18, fill: { color: ORANGE }, line: { color: WHITE, width: 1 } });
  text(s, "閾値監視：超えてから気づく", xOf(ch, cross) - 3.3, yOf(ch, 49.2), 3.1, 0.3, { size: 13, color: ORANGE, bold: true, align: "right" });
  // 予測: 数時間前に分かる（点の左上、閾値の下に置く）
  s.addShape(pres.shapes.OVAL, { x: xOf(ch, early) - 0.09, y: yOf(ch, curve[early]) - 0.09, w: 0.18, h: 0.18, fill: { color: BLUE }, line: { color: WHITE, width: 1 } });
  text(s, "予測：\n超える前に知らせる", M + 0.1, yOf(ch, 43.2), xOf(ch, early) - M - 0.2, 0.52, { size: 13, color: BLUE, bold: true });
  // 2 つの時刻の差（曲線の下）
  const ay = yOf(ch, 33.3);
  s.addShape(pres.shapes.LINE, { x: xOf(ch, early), y: ay, w: xOf(ch, cross) - xOf(ch, early), h: 0, line: { color: BLUE, width: 1.5, beginArrowType: "triangle", endArrowType: "triangle" } });
  text(s, "数時間の余裕", (xOf(ch, early) + xOf(ch, cross)) / 2 - 0.8, ay + 0.06, 1.6, 0.26, { size: 12, color: BLUE, bold: true, align: "center" });
  text(s, "時間 →", p.x + p.w - 0.8, p.y + p.h + 0.05, 0.8, 0.22, { size: 11, color: MUTED, align: "right" });
  bullets(s, [
    "油温は、劣化や事故の\nリスクと密接に関係する",
    "これまでは経験則の\n閾値監視が中心",
    "負荷の変動や外部要因で、\n温度の動きが複雑に",
  ], { x: 6.75, y: 1.45, w: 2.75, size: 15, gap: 0.3, emphasis: [] });
  note(s, "ご相談時に伺った背景より。図はイメージで、実データではない");
}

// ========== 4. 問い
{
  const s = slide();
  titleTwoLines(s, "問いは、予測の精度・高温の事前検知・\n外れる理由の3つに絞った");
  table(s, [
    ["問い", "測るもの", "事前に決めた目標"],
    ["予測の精度", "平均絶対誤差（MAE）", "最良の比較手法より\nMAEで10%以上改善"],
    ["高温の事前検知", "検知率・通知数・先行時間", "通知が平均1日1件以下で\n検知率50%以上"],
    ["外れる理由と効く情報", "誤差の傾向と特徴量の効果", "目標は置かず、改善点を探す"],
  ], { y: 1.55, colW: [2.5, 3.0, 3.5], size: 16, rowH: 0.62, bold: [0] });
  note(s, "高温は、学習期間の油温の上位5%で定義した。故障や危険を示す温度ではない");
}

// ========== 5. データ — 事実と、それを設計にどう使ったか
{
  const s = slide();
  const di = D.diurnal_ETTh2;
  const amps = Object.values(di.amplitude);
  title(s, `油温は1日で${Math.round(Math.min(...amps))}〜${Math.round(Math.max(...amps))}℃動き、負荷との相関は最大0.50`);
  const order = [["JJA", "夏"], ["SON", "秋"], ["MAM", "春"], ["DJF", "冬"]];
  const ch = lineChart(s, {
    x: M, y: 1.3, w: 5.6, h: 3.5, labels: di.hours.map((h) => String(h)), freq: 3, min: 5, max: 50, step: 10,
    series: order.map(([k, n]) => ({ name: n, values: di.series[k], color: k === "SON" ? "7F7F7F" : NAVY })),
    layout: { x: 0.09, y: 0.04, w: 0.82, h: 0.82 }, lineSize: 2,
  });
  endLabels(s, ch, [["JJA", "夏"], ["DJF", "冬"]].map(([k, n]) => ({ text: n, value: di.series[k][23], color: NAVY, bold: true })), { w: 0.4, size: 12 });
  const xp = xOf(ch, di.peak_hour.MAM);
  text(s, "春", xp - 0.2, yOf(ch, di.series.MAM[di.peak_hour.MAM]) - 0.32, 0.4, 0.24, { size: 12, color: NAVY, bold: true, align: "center" });
  text(s, "秋", xp - 0.2, yOf(ch, di.series.SON[di.peak_hour.SON]) + 0.08, 0.4, 0.24, { size: 12, color: "7F7F7F", bold: true, align: "center" });
  chartLabel(s, "平均油温（℃）", ch.x, ch.y - 0.2);
  text(s, "時刻", plot(ch).x + plot(ch).w - 0.5, ch.y + ch.h - 0.15, 0.5, 0.22, { size: 11, color: MUTED, align: "right" });
  const tr = Object.values(di.trough_hour);
  text(s, `どの季節も${di.peak_hour.MAM}時ごろ最高、朝${Math.min(...tr)}〜${Math.max(...tr)}時に最低`, plot(ch).x + 0.1, yOf(ch, 50) + 0.02, 4.2, 0.3, { size: 13, color: BLUE, bold: true });
  bullets(s, [
    "直前の値と強く連動\n（1時間前との相関 0.994）\n→ 油温の履歴を土台に",
    "負荷との相関は最大0.50\n→ 効くかは実験で確かめる",
  ], { x: 6.55, y: 1.55, w: 2.95, size: 16, gap: 0.4, emphasis: [] });
  note(s, "設備2の全期間（2016年7月〜2018年6月）の時刻別平均。データ：Zhou et al. (2021)");
}

// ========== 6. 評価のやり方 — 使う情報と期間の分け方
{
  const s = slide();
  title(s, "観測済みの情報で6時間先を予測し、最後の8か月で採点");
  // 上: 予測に使う情報
  const y0 = 1.72, x0 = 0.9, xs = 5.2, xT = 7.7, x1 = 9.1;
  s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y0 - 0.045, w: xs - x0, h: 0.09, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addShape(pres.shapes.RECTANGLE, { x: xs, y: y0 - 0.045, w: x1 - xs, h: 0.09, fill: { color: LIGHT }, line: { color: LIGHT, width: 0 } });
  text(s, "観測済み（油温と負荷）", x0, y0 - 0.45, xs - x0, 0.3, { size: 14, color: NAVY, bold: true, align: "center" });
  text(s, "未来（使うのは時刻と季節だけ）", xs, y0 - 0.45, x1 - xs, 0.3, { size: 14, color: MUTED, bold: true, align: "center" });
  s.addShape(pres.shapes.OVAL, { x: xs - 0.11, y: y0 - 0.11, w: 0.22, h: 0.22, fill: { color: NAVY }, line: { color: WHITE, width: 1.5 } });
  s.addShape(pres.shapes.OVAL, { x: xT - 0.11, y: y0 - 0.11, w: 0.22, h: 0.22, fill: { color: BLUE }, line: { color: WHITE, width: 1.5 } });
  text(s, "予測時点（現在）", xs - 1.0, y0 + 0.17, 2.0, 0.28, { size: 14, color: INK, bold: true, align: "center" });
  text(s, "予測先（6時間後）", xT - 1.3, y0 + 0.17, 2.6, 0.28, { size: 14, color: BLUE, bold: true, align: "center" });
  // 中: 期間の分け方
  const y = 2.72, h = 0.46, unit = CW / (12 + 4 + 7.9);
  let x = M;
  [["学習　12か月", "2016年7月〜2017年6月", 12, NAVY], ["検証　4か月", "2017年7月〜10月", 4, MUTED], ["テスト　約8か月", "2017年11月〜2018年6月", 7.9, BLUE]].forEach(([t, sub, m, c]) => {
    const w = unit * m;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    text(s, t, x, y, w, h, { size: 14, color: WHITE, bold: true, align: "center", valign: "middle" });
    text(s, sub, x, y + h + 0.05, w, 0.24, { size: 11, color: MUTED, align: "center" });
    x += w;
  });
  // 下: 3 行だけ
  bullets(s, [
    "比較相手：持続予測（今と同じ）・前日同時刻・リッジ回帰",
    "設定は検証期間だけで選び、テストは最後に一度だけ採点",
    "未来の情報が混ざらないことを、実行のたびに自動で確認",
  ], { y: 3.72, size: 14, gap: 0.16, emphasis: [] });
  note(s, "参考評価として、予測先の負荷が事前に分かると仮定した場合も比べた。1時間先と24時間先の結果は付録B");
}

// ========== 7. 結果1 精度
{
  const s = slide();
  const r6 = raw.ETTh2["6"], m6 = mae.ETTh2["6"], r6b = raw.ETTh1["6"];
  title(s, `6時間先の誤差は${f2(m6.lgbm)}℃、最良の比較手法より${cut(r6.ridge, r6.lgbm)}%小さい`);
  const ch = barChart(s, {
    x: M, y: 1.25, w: 6.2, h: 3.55, labels: ["持続予測", "前日同時刻", "リッジ回帰", "LightGBM"],
    values: [m6.persistence, m6.seasonal, m6.ridge, m6.lgbm], colors: ["BFBFBF", "BFBFBF", "7F7F7F", BLUE], max: 5.2,
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.78 },
  });
  chartLabel(s, "設備2・平均絶対誤差（℃）", ch.x, ch.y - 0.05);
  const xL = xOf(ch, 3), yTop = yOf(ch, r6.lgbm);
  text(s, "リッジ比", xL - 0.55, yTop - 1.02, 1.1, 0.26, { size: 12, color: BLUE, align: "center" });
  text(s, `−${cut(r6.ridge, r6.lgbm)}%`, xL - 0.75, yTop - 0.75, 1.5, 0.34, { size: 20, color: BLUE, bold: true, align: "center" });
  bullets(s, [`設備1の改善は${cut(r6b.persistence, r6b.lgbm)}%だけ\n1・24時間先は付録B`], { x: 6.95, y: 2.2, w: 2.55, size: 15, emphasis: [], color: MUTED });
  note(s, `設備2・6時間先・テスト期間（${D.n_test_h6.ETTh2.toLocaleString("en-US")}時点）。データ：ETT（Zhou et al., 2021）`);
}

// ========== 8. 結果2 予測と実測（週の誤差がテスト期間の中央値に最も近い 1 週間）
{
  const s = slide();
  title(s, "6時間前の予測は、1日の山と谷を追えている");
  const labels = WEEK.labels;
  const ch = lineChart(s, {
    x: M, y: 1.3, w: CW, h: 3.45, labels, min: 5, max: 40, step: 5, noXLabels: true,
    series: [{ name: "実測", values: WEEK.actual, color: INK }, { name: "6時間前の予測", values: WEEK.pred, color: BLUE }],
    layout: { x: 0.05, y: 0.05, w: 0.8, h: 0.8 }, lineSize: 1.75,
  });
  const p = plot(ch);
  labels.forEach((lb, i) => {
    if (i % 24) return;
    const [md] = lb.split(" ");
    text(s, md, xOf(ch, i) - 0.05, p.y + p.h + 0.05, 0.6, 0.22, { size: 11, color: MUTED });
  });
  endLabels(s, ch, [
    { text: "実測", value: WEEK.actual[labels.length - 1], color: INK, bold: true },
    { text: "6時間前の予測", value: WEEK.pred[labels.length - 1], color: BLUE, bold: true },
  ], { w: 1.3, size: 12 });
  chartLabel(s, "設備2・油温（℃）", ch.x, ch.y - 0.2);
  note(s, `設備2・${WEEK.start.slice(0, 10).replace(/-/g, "/")}〜${WEEK.end.slice(5, 10).replace(/-/g, "/")}。この週の誤差 ${f2(WEEK.mae)}℃ は、テスト期間の${WEEK.n_full_weeks}週の中央値（${f2(WEEK.median_weekly_mae)}℃）と同じ。都合のよい週を選ばないよう中央値の週を使った`);
}

// ========== 9. 結果3 安定性
{
  const s = slide();
  titleTwoLines(s, "設備2では、テスト期間の8か月すべてで\n持続予測とリッジ回帰より誤差が小さかった");
  const mo = D.monthly_h6.ETTh2;
  const labels = mo.months.map((m) => `${Number(m.slice(5))}月`);
  const ch = lineChart(s, {
    x: M, y: 1.5, w: 5.2, h: 3.3, labels, min: 0, max: 6, step: 1,
    series: [{ name: "持続予測", values: mo.persistence, color: GRAY }, { name: "リッジ回帰", values: mo.ridge, color: "7F7F7F" }, { name: "LightGBM", values: mo.lgbm, color: BLUE }],
    layout: { x: 0.07, y: 0.05, w: 0.74, h: 0.82 },
  });
  const last = mo.months.length - 1;
  endLabels(s, ch, [
    { text: "持続予測", value: mo.persistence[last], color: GRAY },
    { text: "リッジ回帰", value: mo.ridge[last], color: "7F7F7F" },
    { text: "LightGBM", value: mo.lgbm[last], color: BLUE, bold: true },
  ], { w: 1.1, size: 12 });
  text(s, "平均絶対誤差（℃）", ch.x, ch.y - 0.2, 2.2, 0.22, { size: 11, color: MUTED });
  const imp = mo.ridge.map((v, i) => 1 - mo.lgbm[i] / v);
  const lo = Math.round(Math.min(...imp) * 100), hi = Math.round(Math.max(...imp) * 100);
  text(s, "8か月すべてで最小", xOf({ ...ch, labels }, 0) - 0.1, yOf(ch, Math.min(...mo.lgbm)) + 0.12, 2.4, 0.28, { size: 13, color: BLUE, bold: true });
  const wk = D.weekly_win_rate_h6;
  bulletsFit(s, [
    `月ごとに見ても、\nリッジ回帰より${lo}〜${hi}%小さい`,
    `週ごとでも${wk.ETTh2.weeks}週中、\n持続予測より${Math.round(wk.ETTh2.vs_persistence * wk.ETTh2.weeks)}週、\nリッジ回帰より${Math.round(wk.ETTh2.vs_ridge * wk.ETTh2.weeks)}週で小さい`,
    `設備1が持続予測より\n小さかったのは${wk.ETTh1.weeks}週中${Math.round(wk.ETTh1.vs_persistence * wk.ETTh1.weeks)}週\n週によって優劣が分かれた`,
  ], { x: 6.4, y: 1.5, w: 3.1, h: 3.3, size: 15, gap: 12 });
  note(s, "設備2・6時間先・主な評価の条件。月と週は予測先の時刻で区切り、それぞれの平均絶対誤差を比べた。データ：ETT（Zhou et al., 2021）");
}

// ========== 10. 結果4 何が効いたか
{
  const s = slide();
  title(s, "効いたのは負荷ではなく、時刻と季節だった");
  const ab = D.ablation_h6.ETTh2, ab1 = D.ablation_h6.ETTh1;
  const keys = ["1_ot_only", "2_+calendar", "3_+load_origin (B)", "4_+load_target (A)"];
  const ch = barChart(s, {
    x: M, y: 1.25, w: 5.9, h: 3.55, labels: ["油温の履歴\nだけ", "＋時刻・季節", "＋予測時点\nまでの負荷", "＋予測先の\n実測負荷（参考）"],
    values: keys.map((k) => ab[k]), colors: [GRAY, BLUE, GRAY, GRAY], max: 3.2, catSize: 12,
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.72 },
  });
  chartLabel(s, "設備2・平均絶対誤差（℃）", ch.x, ch.y - 0.05);
  const xm = (xOf(ch, 0) + xOf(ch, 1)) / 2;
  text(s, "履歴だけ比", xm - 0.2, yOf(ch, ab[keys[0]]) - 0.2, 1.4, 0.24, { size: 12, color: BLUE });
  text(s, `−${cut(ab[keys[0]], ab[keys[1]])}%`, xm - 0.2, yOf(ch, ab[keys[0]]) + 0.05, 1.2, 0.34, { size: 20, color: BLUE, bold: true });
  bullets(s, [
    `主な結果の${f2(mae.ETTh2["6"].lgbm)}℃は、\nテスト前に固定した\n負荷ありの構成`,
    `設備1も同じ傾向\n（${f2(ab1[keys[0]])}→${f2(ab1[keys[1]])}℃）`,
  ], { x: 6.75, y: 1.6, w: 2.75, size: 15, gap: 0.35, emphasis: [0] });
  note(s, "設備2・6時間先・テスト期間。特徴量は左から順に足した（4本目は予測時点では分からない情報を使う参考）");
}

// ========== 11. 結果5 高温の事前検知
{
  const s = slide();
  title(s, `高温${lg.events}区間のうち${lg.detected}区間を、平均${Math.round(lg.lead_h)}時間前に検知した`);
  const ch = barChart(s, {
    x: M, y: 1.25, w: 6.0, h: 3.55, labels: ["持続予測", "前日同時刻", "リッジ回帰", "LightGBM"],
    values: [det.persistence.detected, det.seasonal.detected, det.ridge.detected, lg.detected], colors: [GRAY, GRAY, GRAY, BLUE], max: 15, fmt: "0",
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.78 },
  });
  chartLabel(s, "設備2・事前に検知できた高温区間の数", ch.x, ch.y - 0.05, 4);
  const p = plot(ch), yT = yOf(ch, lg.events);
  s.addShape(pres.shapes.LINE, { x: p.x, y: yT, w: p.w, h: 0, line: { color: ORANGE, width: 1, dashType: "dash" } });
  text(s, `高温区間は全部で${lg.events}`, p.x + 0.05, yT - 0.3, 2.4, 0.26, { size: 12, color: ORANGE, bold: true });
  bullets(s, [
    `通知${lg.notices}件のうち\n誤通知は${lg.false_notices}件`,
    "高温区間は5〜6月に集中",
  ], { x: 6.85, y: 1.6, w: 2.65, size: 16, gap: 0.35, emphasis: [0] });
  note(s, `設備2・6時間先・閾値${f1(thr2)}℃（学習期間の上位5%）。区間と通知の数え方は付録D`);
}

// ========== 12. 結果6 事例（注目点は図に直接書く）
function eventChart(s, ev, { x = M, y = 1.3, w = 8.2, h = 3.5, min = 26, max = 50, layoutW = 0.8 } = {}) {
  const ch = lineChart(s, {
    x, y, w, h, labels: ev.hours.map((t) => String(Number(t))), freq: 3, min, max, step: 4,
    series: [{ name: "実測", values: ev.actual, color: INK }, { name: "6時間前の予測", values: ev.pred, color: BLUE }, { name: "閾値", values: ev.hours.map(() => thr2), color: ORANGE }],
    layout: { x: 0.07, y: 0.05, w: layoutW, h: 0.82 }, lineSize: 2,
  });
  const p = plot(ch), n = ev.hours.length, slot = p.w / n;
  if (ev.exceed_hours.length) {
    const a = Math.min(...ev.exceed_hours), b = Math.max(...ev.exceed_hours);
    s.addShape(pres.shapes.RECTANGLE, { x: p.x + slot * a, y: p.y, w: slot * (b - a + 1), h: p.h, fill: { color: ORANGE, transparency: 88 }, line: { color: ORANGE, width: 0, transparency: 100 } });
  }
  ev.alarm_hours.forEach((hh) => {
    const cx = xOf(ch, hh), cy = yOf(ch, ev.pred[hh]);
    s.addShape(pres.shapes.OVAL, { x: cx - 0.06, y: cy - 0.06, w: 0.12, h: 0.12, fill: { color: ORANGE }, line: { color: WHITE, width: 0.75 } });
  });
  endLabels(s, ch, [
    { text: "実測", value: ev.actual[n - 1], color: INK, bold: true },
    { text: "6時間前の予測", value: ev.pred[n - 1], color: BLUE, bold: true },
    { text: `閾値 ${f1(thr2)}℃`, value: thr2, color: ORANGE, bold: true },
  ], { w: 1.4, size: 11 });
  chartLabel(s, "油温（℃）", ch.x, ch.y - 0.2);
  text(s, "予測先の時刻", p.x + p.w - 1.4, ch.y + ch.h - 0.14, 1.4, 0.22, { size: 11, color: MUTED, align: "right" });
  return ch;
}
{
  const s = slide();
  const ev = D.event_detected;
  const a = Math.min(...ev.exceed_hours), first = Math.min(...ev.alarm_hours);
  const issued = first - D.horizon_h, lead = a - issued;
  title(s, `${issued}時に${first}時の高温を予測できたが、上がり始めは逃した`);
  const ch = eventChart(s, ev, { min: 30, max: 50 });
  const p = plot(ch), tx = p.x + 0.12, ty = yOf(ch, 49.6), tw = 2.9;
  text(s, `${issued}時に警報（超える${lead}時間前）\n→ ${first}時の超過を予測`, tx, ty, tw, 0.46, { size: 13, color: ORANGE, bold: true });
  const x1 = tx + tw - 0.3, y1 = ty + 0.2, x2 = xOf(ch, first) - 0.07, y2 = yOf(ch, ev.pred[first]);
  s.addShape(pres.shapes.LINE, { x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1), flipV: y2 < y1, line: { color: ORANGE, width: 0.75 } });
  text(s, `${a}〜${first - 1}時の超過は\n予測が閾値に届かず`, xOf(ch, a) - 0.05, yOf(ch, ev.pred[a]) + 0.12, 1.8, 0.42, { size: 12, color: BLUE, bold: true });
  note(s, "設備2・2018年5月15日。横軸は予測先の時刻（予測はその6時間前に出したもの）。橙の帯は実測が閾値を超えた時間");
}

// ========== 13. 設計上の判断 — 変化量の予測を式の図に、残りは表
{
  const s = slide();
  title(s, "変化量を予測し、先の時間ごとに別のモデルを作った");
  const y = 1.3, bh = 0.62;
  const box = (x, w, t, c, bold = true) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: bh, fill: { color: WHITE }, line: { color: c, width: 1.25 } });
    text(s, t, x, y, w, bh, { size: 16, color: c, bold, align: "center", valign: "middle" });
  };
  box(M, 2.6, "6時間後の油温", NAVY);
  text(s, "＝", M + 2.6, y, 0.6, bh, { size: 22, color: MUTED, bold: true, align: "center", valign: "middle" });
  box(M + 3.2, 2.3, "今の油温", INK, false);
  text(s, "＋", M + 5.5, y, 0.6, bh, { size: 22, color: MUTED, bold: true, align: "center", valign: "middle" });
  box(M + 6.1, 2.9, "予測した変化量", BLUE);
  text(s, "油温そのものを予測するより、検証期間の誤差が小さかった", M, y + bh + 0.1, CW, 0.3, { size: 13, color: MUTED });
  table(s, [
    ["ほかの判断", "理由"],
    ["1・6・24時間先で別々のモデル", "予測を入力に戻して繰り返すと、誤差がたまる"],
    ["絶対誤差で学習", "評価指標（MAE）と揃える"],
    ["特徴量と閾値はテスト前に固定", "結果を見た選び直しを防ぐ"],
  ], { y: 2.65, colW: [4.2, 4.8], size: 15, rowH: 0.52, bold: [0] });
}

// ========== 14. 限界 — まだ分からないことと、次にどう確かめるか
{
  const s = slide();
  title(s, "保全効果と他の設備での有効性は、まだ確認できていない");
  table(s, [
    ["まだ分からないこと", "次にどう確かめるか"],
    ["保全の効果（故障・点検の記録がない）", "試行で、通知と現場の対応を記録する"],
    ["他の設備での有効性（評価は2台）", "設備ごとに学習し直して評価する"],
    ["危険温度・通年での検知\n（高温は上位5%、5〜6月に集中）", "閾値の意味を現場と決め、通年で評価する"],
    ["24時間先の精度", "外気温や気象予報を加える"],
  ], { y: 1.3, colW: [4.5, 4.5], size: 15, rowH: 0.62, bold: [0] });
}

// ========== 15. 結論と次の一手 — 3 段階の流れ
{
  const s = slide({ dark: true });
  const r6 = raw.ETTh2["6"];
  title(s, "まず通知せずに予測を試し、通知の数と外れを確かめる", { color: WHITE });
  text(s, `設備2：6時間先の誤差 −${cut(r6.ridge, r6.lgbm)}%、高温${lg.events}区間中${lg.detected}区間を事前に検知`, M, 1.2, CW, 0.4, { size: 18, color: WHITE, bold: true });
  text(s, "保全効果と、評価していない設備での有効性は未検証", M, 1.65, CW, 0.32, { size: 14, color: TINT });
  const steps = ["通知せずに\n予測だけ動かす", "通知候補の数・先行時間・\n見逃しを記録する", "許容できる誤通知の数を\n現場と決めて通知を始める"];
  const bw = 2.65, gap = (CW - 3 * bw) / 2, by = 2.45, bh = 1.05;
  steps.forEach((t, i) => {
    const x = M + i * (bw + gap);
    s.addShape(pres.shapes.RECTANGLE, { x, y: by, w: bw, h: bh, fill: { color: NAVY }, line: { color: TINT, width: 1 } });
    text(s, String(i + 1), x + 0.14, by + 0.1, 0.4, 0.3, { size: 13, color: TINT, bold: true });
    text(s, t, x + 0.1, by + 0.3, bw - 0.2, bh - 0.35, { size: 14, color: WHITE, bold: true, align: "center", valign: "middle" });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + bw + 0.06, y: by + bh / 2, w: gap - 0.12, h: 0, line: { color: TINT, width: 1.25, endArrowType: "triangle" } });
  });
  text(s, "github.com/HJRKTNG/ett-oil-temperature-poc　｜　沓脱 聖", M, 4.75, 7.5, 0.3, { size: 12, color: TINT });
}

// ========== 16. 参考文献
{
  const s = slide();
  title(s, "参考文献");
  table(s, [
    ["データの出典論文", "Zhou, H. et al. (2021). Informer: Beyond Efficient Transformer for Long Sequence Time-Series Forecasting. AAAI 2021."],
    ["データの配布元", "ETDataset. https://github.com/zhouhaoyi/ETDataset"],
    ["利用条件", "CC BY-ND 4.0。配布 CSV は改変せず保存し、特徴量はコードで作成"],
    ["予測手法", "Ke, G. et al. (2017). LightGBM: A Highly Efficient Gradient Boosting Decision Tree. NeurIPS 2017."],
  ], { y: 1.2, colW: [1.9, 7.1], size: 14, rowH: 0.62, head: false, bold: [0] });
}

// ========== 付録A 特徴量とモデル設定
{
  const s = slide();
  appendixLabel(s, "付録A　特徴量とモデル設定");
  title(s, "特徴量は、予測時点までの値と予測先の時刻・季節だけ", { size: 22, y: 0.45 });
  table(s, [
    ["油温の履歴", "現在値、1・2・3・6・12・24・48・168時間前の値、\n現在値と1時間前・24時間前との差、\n6・24・168時間の移動平均・標準偏差・最大・最小"],
    ["時刻・季節", "予測先の時刻と年内の日（sin・cos）、曜日"],
    ["負荷6種", "現在値、現在値と1時間前との差、24時間平均\n参考評価のみ：予測先の実測負荷も追加"],
    ["LightGBM", "絶対誤差（L1）で学習。学習率0.03、最大葉数31、\n葉ごとの最小データ数50、特徴量と行の抽出率0.8。\n検証期間の誤差が100回続けて改善しなければ学習を終了"],
  ], { y: 1.3, colW: [1.8, 7.2], size: 14, rowH: 0.7, head: false, bold: [0] });
}

// ========== 付録B 予測する時間ごとの結果（持続予測と LightGBM）
{
  const s = slide();
  appendixLabel(s, "付録B　予測する時間ごとの結果");
  title(s, "24時間先では、持続予測に対する改善がなくなる", { size: 22, y: 0.45 });
  const hs = ["1", "6", "24"], labels = ["1時間先", "6時間先", "24時間先"];
  [["ETTh1", "設備1", M, 2.0, 0.5, "0.0"], ["ETTh2", "設備2", M + 4.6, 5.0, 1, "0"]].forEach(([k, name, x, max, step, fmt]) => {
    const series = [["persistence", "持続予測", GRAY], ["lgbm", "LightGBM", BLUE]];
    const ch = lineChart(s, {
      x, y: 1.35, w: 4.3, h: 3.35, labels, min: 0, max, step, fmt, symbol: "circle",
      series: series.map(([sk, sn, c]) => ({ name: sn, values: hs.map((h) => mae[k][h][sk]), color: c })),
      layout: { x: 0.1, y: 0.12, w: 0.64, h: 0.74 }, lineSize: 2,
    });
    text(s, name, x, 1.15, 2, 0.28, { size: 14, color: NAVY, bold: true });
    endLabels(s, ch, series.map(([sk, sn, c]) => ({ text: sn, value: mae[k]["24"][sk], color: c, bold: sk === "lgbm" })), { w: 1.2, size: 11, minGap: 0.24 });
    hs.forEach((h, i) => {
      const v = mae[k][h].lgbm, pv = mae[k][h].persistence, below = v <= pv;
      const pb = plot(ch), low = yOf(ch, v) + 0.32 > pb.y + pb.h;   // 下に置くと横軸の文字と重なるときは右に置く
      if (below && low) text(s, f2(v), xOf(ch, i) + 0.1, yOf(ch, v) - 0.11, 0.7, 0.22, { size: 11, color: BLUE, bold: true });
      else text(s, f2(v), xOf(ch, i) - 0.35, yOf(ch, v) + (below ? 0.08 : -0.3), 0.7, 0.22, { size: 11, color: BLUE, bold: true, align: "center" });
      text(s, f2(pv), xOf(ch, i) - 0.35, yOf(ch, pv) + (below ? -0.3 : 0.08), 0.7, 0.22, { size: 11, color: MUTED, align: "center" });
    });
  });
  note(s, "平均絶対誤差（℃）・テスト期間・主な評価の条件。リッジ回帰と前日同時刻は省略。データ：ETT（Zhou et al., 2021）");
}

// ========== 付録C 見逃した例
{
  const s = slide();
  const ev = D.event_missed;
  appendixLabel(s, "付録C　見逃した例");
  const a = Math.min(...ev.exceed_hours), b = Math.max(...ev.exceed_hours);
  title(s, `${a}時と${b}時だけの短い高温は、6時間前に捉えられなかった`, { size: 22, y: 0.45 });
  const ch = eventChart(s, ev, { min: 26, max: 50 });
  [a, b].forEach((hh) => {
    s.addShape(pres.shapes.OVAL, { x: xOf(ch, hh) - 0.07, y: yOf(ch, ev.pred[hh]) - 0.07, w: 0.14, h: 0.14, fill: { color: WHITE }, line: { color: BLUE, width: 1.5 } });
  });
  const tx = xOf(ch, 8.6), ty = yOf(ch, 29.4);
  text(s, `白丸：${a - D.horizon_h}時・${b - D.horizon_h}時に出した予測\n（${f1(ev.pred_by_hour[a])}℃・${f1(ev.pred_by_hour[b])}℃）はどちらも閾値未満`, tx, ty, 3.2, 0.46, { size: 12, color: BLUE, bold: true });
  s.addShape(pres.shapes.LINE, { x: xOf(ch, a) - 0.02, y: yOf(ch, ev.pred[a]) + 0.08, w: 0, h: ty - yOf(ch, ev.pred[a]) - 0.12, line: { color: BLUE, width: 0.75 } });
  note(s, "設備2・2018年6月4日。横軸は予測先の時刻。次の検証で、通知の閾値を下げたときの見逃しと誤通知を確かめる");
}

// ========== 付録D 高温検知の数え方（5月15日の実例を 1 本の時間軸で）
{
  const s = slide();
  appendixLabel(s, "付録D　高温検知の数え方");
  const ev = D.event_detected, hz = D.horizon_h;
  const a = Math.min(...ev.exceed_hours), b = Math.max(...ev.exceed_hours);
  const firstIssue = Math.min(...ev.alarm_hours) - hz;
  title(s, "区間・通知・先行時間を、5月15日の例で数える", { size: 22, y: 0.45 });
  const h0 = 5, h1 = 19, gx0 = 1.3, gx1 = 8.05, top = 1.45, rowH = 0.34;
  const X = (hh) => gx0 + (gx1 - gx0) * (hh - h0) / (h1 - h0);
  const issues = [];
  for (let t = a - hz; t <= b - hz; t++) if (t < a) issues.push(t);   // 予測時点で実測が閾値以下の起点（9〜11時）＋その前（6〜8時）
  const bottom = top + 0.5 + issues.length * rowH;
  // 実測が閾値を超えた時間（予測先の時刻の帯）
  s.addShape(pres.shapes.RECTANGLE, { x: X(a), y: top, w: X(b + 1) - X(a), h: bottom - top, fill: { color: ORANGE, transparency: 88 }, line: { color: ORANGE, width: 0, transparency: 100 } });
  text(s, `高温区間：実測が閾値を連続して超えた${a}〜${b}時 ＝ 1区間`, X(a), top + 0.04, X(b + 1) - X(a) + 1.2, 0.26, { size: 12, color: ORANGE, bold: true });
  issues.forEach((t, i) => {
    const y = top + 0.5 + i * rowH + rowH / 2;
    const alarm = ev.alarm_hours.includes(t + hz);
    const c = alarm ? ORANGE : GRAY;
    s.addShape(pres.shapes.OVAL, { x: X(t) + 0.02, y: y - 0.06, w: 0.12, h: 0.12, fill: { color: c }, line: { color: c, width: 0 } });
    s.addShape(pres.shapes.LINE, { x: X(t) + 0.14, y, w: X(t + hz) - X(t) - 0.08, h: 0, line: { color: c, width: alarm ? 1.75 : 1, endArrowType: "triangle", dashType: alarm ? "solid" : "dash" } });
    text(s, `${t}時の予測`, 0.35, y - 0.12, 0.85, 0.24, { size: 11, color: alarm ? ORANGE : MUTED, bold: alarm, align: "right" });
  });
  const alarmRows = issues.map((t, i) => [t, i]).filter(([t]) => ev.alarm_hours.includes(t + hz));
  const yA = top + 0.5 + alarmRows[0][1] * rowH, yB = top + 0.5 + (alarmRows[alarmRows.length - 1][1] + 1) * rowH;
  // 通知（連続した警報を 1 件）: 右端の括弧
  const bx = X(b - 0) + 0.25;
  s.addShape(pres.shapes.LINE, { x: bx, y: yA + 0.04, w: 0, h: yB - yA - 0.08, line: { color: ORANGE, width: 1.5 } });
  text(s, "通知 1件\n連続した警報を\nまとめて数える", bx + 0.1, yA + 0.02, 1.45, 0.66, { size: 12, color: ORANGE, bold: true });
  // 時刻の目盛り
  hline(s, gx0, bottom + 0.05, gx1 - gx0, GRAY, 0.75);
  for (let hh = h0 + 1; hh <= h1; hh += 1) if (hh % 3 === 0) text(s, `${hh}時`, X(hh) - 0.3, bottom + 0.1, 0.6, 0.22, { size: 11, color: MUTED, align: "center" });
  // 先行時間
  const ly = bottom + 0.45;
  s.addShape(pres.shapes.LINE, { x: X(firstIssue) + 0.08, y: ly, w: X(a) - X(firstIssue) - 0.08, h: 0, line: { color: BLUE, width: 1.75, beginArrowType: "oval", endArrowType: "triangle" } });
  text(s, `先行時間 ${a - firstIssue}時間：最初に当たった予測を出した${firstIssue}時 → 実際に超えた${a}時`, X(firstIssue) - 0.1, ly + 0.08, 6.5, 0.26, { size: 12, color: BLUE, bold: true });
  note(s, "灰色の破線＝予測が閾値未満（警報なし）、橙＝警報。誤通知は、通知が指した時刻に実測が一度も閾値を超えなかった通知（この日は0件）");
}

// ========== 付録E 高温の基準を上位10%にした場合
{
  const s = slide();
  appendixLabel(s, "付録E　高温の基準を上位10%にした場合");
  const q = lg.q90, pc = (x) => Math.round(x * 100);
  title(s, "上位10%に広げても検知率はほぼ同じで、通知は倍に", { size: 22, y: 0.45 });
  const changed = [4, 5, 6];
  table(s, [
    ["設備2・6時間先", `上位5%（${f1(thr2)}℃）`, `上位10%（${f1(D.thresholds.ETTh2.q90)}℃）`],
    ["高温区間", `${lg.events}区間`, `${q.events}区間`],
    ["事前に検知", `${lg.detected}区間（${pc(lg.detected / lg.events)}%）`, `${q.detected}区間（${pc(q.detected / q.events)}%）`],
    ["先行時間（平均）", `${f1(lg.lead_h)}時間`, `${f1(q.lead_h)}時間`],
    ["通知", `${lg.notices}件`, `${q.notices}件`],
    ["1日あたりの通知", `${f2(lg.notices_per_day)}件`, `${f2(q.notices_per_day)}件`],
    ["誤通知", `${lg.false_notices}件`, `${q.false_notices}件`],
  ], { y: 1.3, w: 7.2, colW: [2.6, 2.3, 2.3], size: 15, rowH: 0.46, bold: [0], cellStyle: (ri, ci) => (changed.includes(ri) && ci === 2 ? { color: BLUE, bold: true } : null) });
  note(s, "実測の高温区間と予測による通知の両方に、各列の閾値を当てた（通知の条件だけを緩めた比較ではない）");
}

const out = path.join(OUT_DIR, "ett_oil_temperature_poc.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written", out, "slides", page));

