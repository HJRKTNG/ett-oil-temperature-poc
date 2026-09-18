/* 報告スライド（v3）を生成する。文章の控え docs/slides_text.md は、生成後に scripts/dump_slide_text.py で PPTX から書き出す。
 * 数値は reports/slide_data.json（scripts/export_slide_data.py が作る）からだけ読む。
 * デザインの決まり: 白背景（表紙と結論だけ紺）・Arial・色は紺/青/橙の 3 色まで・
 * タイトルは言いたいことを 1 文で・結果のページは図 1 つで左に図、右に解釈・注目点は図に直接書く。
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

// 色は 3 色まで: 紺 = 見出しと主張、青 = 注目する系列、橙 = 閾値と警報。ほかは無彩色。
const NAVY = "1F4E79", BLUE = "2E75B6", ORANGE = "C55A11";
const INK = "262626", BODY = "3A3A3A", MUTED = "6E6E6E", GRAY = "A6A6A6", LIGHT = "D9D9D9", WHITE = "FFFFFF", TINT = "C9D8EA";
const FONT = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.author = "Kutsunugi Hijiri";
pres.title = "変圧器油温の先読み PoC";
const W = 10, H = 5.625, M = 0.5, CW = W - 2 * M;
let page = 0;

const f2 = (v) => Number(v).toFixed(2);
const f1 = (v) => Number(v).toFixed(1);
const cut = (a, b) => Math.round((1 - b / a) * 100); // a から b への減少率（%）
const mae = D.mae, raw = D.mae_raw, det = D.detection_ETTh2_h6, thr2 = D.thresholds.ETTh2.q95, thr1 = D.thresholds.ETTh1.q95;
const lg = det.lgbm;

// ---------- 共通部品
function slide({ dark = false } = {}) {
  const s = pres.addSlide();
  s.background = { color: dark ? NAVY : WHITE };
  page += 1;
  s.addText(String(page), { x: W - M - 0.6, y: H - 0.4, w: 0.6, h: 0.24, fontFace: FONT, fontSize: 10, color: dark ? TINT : MUTED, align: "right", valign: "bottom", margin: 0, isTextBox: true });
  return s;
}
function title(s, text, { color = NAVY, size = 24, y = 0.32, h = 0.95 } = {}) {
  s.addText(text, { x: M, y, w: CW, h, fontFace: FONT, fontSize: size, bold: true, color, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.08 });
}
function appendixLabel(s, text) {
  s.addText(text, { x: M, y: 0.14, w: CW, h: 0.26, fontFace: FONT, fontSize: 12, italic: true, color: MUTED, margin: 0, isTextBox: true });
}
// 項目の中の "\n" は、同じ箇条書きの中での改行（語の途中で折り返さないよう、改行位置を自分で決める）
function bullets(s, items, { x = M, y = 1.5, w = CW, h = 3.3, size = 20, gap = 12, color = BODY } = {}) {
  const runs = [];
  items.forEach((t, i) => {
    const parts = String(t).split("\n");
    parts.forEach((part, j) => {
      // 続きの行は、見えない行頭記号（ノーブレークスペース）の段落にして字下げをそろえる
      const bullet = j === 0 ? { indent: Math.round(size * 0.9) } : { indent: Math.round(size * 0.9), characterCode: "00A0" };
      const o = { bullet, paraSpaceAfter: j === parts.length - 1 ? gap : 0 };
      if (!(i === items.length - 1 && j === parts.length - 1)) o.breakLine = true;
      runs.push({ text: part, options: o });
    });
  });
  s.addText(runs, { x, y, w, h, fontFace: FONT, fontSize: size, color, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.08 });
}
function text(s, t, x, y, w, h, { size = 14, color = BODY, bold = false, align = "left", valign = "top", italic = false } = {}) {
  s.addText(t, { x, y, w, h, fontFace: FONT, fontSize: size, color, bold, italic, align, valign, margin: 0, isTextBox: true, lineSpacingMultiple: 1.08 });
}
function note(s, t, { dark = false } = {}) {
  s.addText(t, { x: M, y: H - 0.72, w: CW - 0.8, h: 0.5, fontFace: FONT, fontSize: 11, color: dark ? TINT : MUTED, valign: "bottom", margin: 0, isTextBox: true, lineSpacingMultiple: 1.05 });
}
function hline(s, x, y, w, color = LIGHT, pt = 0.75) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: pt } });
}
// 罫線だけの表（縦線なし・塗りなし）。rows[0] を見出し行にするかは head で指定。
function table(s, rows, { x = M, y = 1.5, w = CW, colW, size = 16, head = true, rowH = 0.5, bold = [] } = {}) {
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const isHead = head && ri === 0;
    return {
      text: String(c),
      options: {
        fontFace: FONT, fontSize: isHead ? size - 3 : size, color: isHead ? MUTED : (bold.includes(ci) ? NAVY : BODY), bold: !isHead && bold.includes(ci),
        align: "left", valign: "middle", margin: [4, 6, 4, 0],
        border: [{ type: "none" }, { type: "none" }, { pt: 0.75, color: LIGHT }, { type: "none" }],
      },
    };
  }));
  s.addTable(data, { x, y, w, colW, rowH, autoPage: false });
}

// ---------- グラフ（描画領域を固定し、注記の位置を値から計算する）
function plot(ch) { return { x: ch.x + ch.layout.x * ch.w, y: ch.y + ch.layout.y * ch.h, w: ch.layout.w * ch.w, h: ch.layout.h * ch.h }; }
function yOf(ch, v) { const p = plot(ch); return p.y + p.h * (1 - (v - ch.min) / (ch.max - ch.min)); }
function xOf(ch, i) { const p = plot(ch); return p.x + p.w * (i + 0.5) / ch.labels.length; }
function barChart(s, ch) {
  ch.min = ch.min || 0;
  s.addChart(pres.charts.BAR, [{ name: ch.name || "value", labels: ch.labels, values: ch.values }], {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, barDir: "col", chartColors: ch.colors, barGapWidthPct: ch.gap || 80,
    valAxisMinVal: ch.min, valAxisMaxVal: ch.max, valAxisHidden: true, valGridLine: { style: "none" },
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: ch.catSize || 13, catAxisLabelColor: BODY, catAxisLineShow: true, catAxisLineColor: GRAY,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FONT, dataLabelFontSize: ch.valSize || 16, dataLabelFontBold: true, dataLabelColor: INK, dataLabelFormatCode: ch.fmt || "0.00",
    showLegend: false, layout: ch.layout,
  });
  return ch;
}
function lineChart(s, ch) {
  s.addChart(pres.charts.LINE, ch.series.map((se) => ({ name: se.name, labels: ch.labels, values: se.values })), {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, chartColors: ch.series.map((se) => se.color), lineSize: ch.lineSize || 2.25, lineDataSymbol: "none",
    valAxisMinVal: ch.min, valAxisMaxVal: ch.max, valAxisMajorUnit: ch.step, valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12, valAxisLabelColor: MUTED, valAxisLabelFormatCode: "0", valAxisLineShow: false,
    valGridLine: { color: "EBEBEB", size: 0.75 }, catGridLine: { style: "none" },
    catAxisLabelPos: "none", catAxisMajorTickMark: "none", catAxisLineColor: GRAY,
    showLegend: false, layout: ch.layout,
  });
  // 横軸の目盛りの数字は自前で描く（LibreOffice は間引いた目盛りの数字を半マスずらして描くため）
  const p = plot(ch), step = ch.freq || 1;
  ch.labels.forEach((lb, i) => {
    if (i % step) return;
    text(s, lb, xOf(ch, i) - 0.3, p.y + p.h + 0.05, 0.6, 0.22, { size: 12, color: MUTED, align: "center" });
  });
  return ch;
}
// 系列の右端に名前を書く（凡例の代わり）。重なるときは縦にずらす。
function endLabels(s, ch, items, { w = 1.5, size = 13, minGap = 0.26 } = {}) {
  const p = plot(ch);
  const pos = items.map((it) => ({ ...it, y: yOf(ch, it.value) })).sort((a, b) => a.y - b.y);
  for (let i = 1; i < pos.length; i++) if (pos[i].y - pos[i - 1].y < minGap) pos[i].y = pos[i - 1].y + minGap;
  pos.forEach((it) => text(s, it.text, p.x + p.w + 0.08, it.y - 0.13, w, 0.26, { size, color: it.color, bold: !!it.bold, valign: "middle" }));
}

// ========== 表紙
{
  const s = slide({ dark: true });
  text(s, "変圧器の油温を6時間先まで予測し、\n高温になる前に知らせられるか", M, 1.3, 8.8, 1.5, { size: 30, color: WHITE, bold: true });
  text(s, "ETT公開データを使った油温予測の検証報告", M, 2.95, 8.8, 0.4, { size: 16, color: TINT });
  text(s, "沓脱 聖　｜　2026年9月", M, 3.5, 8.8, 0.35, { size: 15, color: TINT });
}

// ========== 1. 要旨
{
  const s = slide();
  const d = det.lgbm, r6 = raw.ETTh2["6"];
  title(s, `設備2では6時間先の誤差を${cut(r6.ridge, r6.lgbm)}%減らし、\n高温${d.events}区間中${d.detected}区間を事前に検知した`);
  table(s, [
    ["精度", `6時間先のMAEは${f2(mae.ETTh2["6"].lgbm)}℃（最良の比較手法は${f2(mae.ETTh2["6"].ridge)}℃）`],
    ["事前検知", `平均${f1(d.lead_h)}時間前に検知。通知${d.notices}件のうち誤通知は${d.false_notices}件`],
    ["限界", "24時間先と設備1は改善が小さく、短い高温は見逃す"],
  ], { y: 1.55, colW: [1.7, 7.3], size: 17, rowH: 0.6, head: false, bold: [0] });
  text(s, "提案：現場へ通知せずに試行し、通知候補の数と外れた数を確かめる", M, 3.65, CW, 0.4, { size: 18, color: NAVY, bold: true });
  note(s, "設備2（ETTh2）・テスト期間 2017年11月〜2018年6月・主な評価（予測時点までの観測だけを使う）の条件。区間と通知は別の数え方（定義は付録D）");
}

// ========== 2. 背景
{
  const s = slide();
  title(s, "超えてから気づく閾値監視から、\n超える前に手を打てる予防保全へ移りたい");
  bullets(s, [
    "油温は、設備の劣化や事故のリスクと密接に関係する",
    "これまでは経験則にもとづく閾値監視が中心だった",
    "負荷の変動や外部要因で、温度の動きが複雑になっている",
  ], { y: 1.75, w: 8.4, size: 22, gap: 22 });
  note(s, "ご相談時に伺った背景より");
}

// ========== 3. 問い
{
  const s = slide();
  title(s, "問いは、予測の精度・高温の事前検知・\n外れる理由の3つに絞った");
  table(s, [
    ["問い", "測るもの", "事前に決めた目標"],
    ["予測の精度", "平均絶対誤差（MAE）", "最良の比較手法より\nMAEで10%以上改善"],
    ["高温の事前検知", "検知率・通知数・先行時間", "通知が平均1日1件以下で\n検知率50%以上"],
    ["外れる理由と効く情報", "誤差の傾向と特徴量の効果", "目標は置かず、改善点を探す"],
  ], { y: 1.55, colW: [2.5, 3.0, 3.5], size: 16, rowH: 0.62, bold: [0] });
  note(s, "高温は、学習期間の油温の上位5%で定義した。故障や危険を示す温度ではない");
}

// ========== 4. データ
{
  const s = slide();
  title(s, "油温は時刻と季節で大きく動き、\n負荷との相関は最大でも0.50だった");
  const di = D.diurnal_ETTh2;
  const order = [["JJA", "夏"], ["SON", "秋"], ["MAM", "春"], ["DJF", "冬"]];
  const ch = lineChart(s, {
    x: M, y: 1.6, w: 5.3, h: 3.2, labels: di.hours.map((h) => String(h)), freq: 3, min: 5, max: 50, step: 10,
    series: order.map(([k, n]) => ({ name: n, values: di.series[k], color: k === "SON" ? "7F7F7F" : NAVY })),
    layout: { x: 0.09, y: 0.04, w: 0.82, h: 0.82 }, lineSize: 2,
  });
  endLabels(s, ch, [["JJA", "夏"], ["DJF", "冬"]].map(([k, n]) => ({ text: n, value: di.series[k][23], color: NAVY, bold: true })), { w: 0.4, size: 12 });
  // 春と秋は右端で重なるので、線が離れている山（${pk}時）の上下に書く
  const xp = xOf(ch, di.peak_hour.MAM);
  text(s, "春", xp - 0.2, yOf(ch, di.series.MAM[di.peak_hour.MAM]) - 0.32, 0.4, 0.24, { size: 12, color: NAVY, bold: true, align: "center" });
  text(s, "秋", xp - 0.2, yOf(ch, di.series.SON[di.peak_hour.SON]) + 0.08, 0.4, 0.24, { size: 12, color: "7F7F7F", bold: true, align: "center" });
  text(s, "平均油温（℃）", ch.x, ch.y - 0.2, 2, 0.22, { size: 11, color: MUTED });
  text(s, "時刻", plot(ch).x + plot(ch).w - 0.5, ch.y + ch.h - 0.02, 0.5, 0.22, { size: 11, color: MUTED, align: "right" });
  const pk = di.peak_hour.MAM;
  const tr = Object.values(di.trough_hour);
  text(s, `どの季節も${pk}時ごろが最高、朝${Math.min(...tr)}〜${Math.max(...tr)}時が最低`, plot(ch).x + 0.1, yOf(ch, 50) + 0.02, 4.0, 0.3, { size: 13, color: BLUE, bold: true });
  bullets(s, [
    `1日の平均的な振れ幅は、\n季節ごとに${f1(Math.min(...Object.values(di.amplitude)))}〜${f1(Math.max(...Object.values(di.amplitude)))}℃`,
    "油温の1時間ラグ相関は0.994\n履歴を予測の土台にする",
    "負荷6種との同時刻の相関は\n最大でも0.50",
    "時刻と季節を特徴量に入れ、\n負荷の効果は誤差で確かめる",
  ], { x: 6.35, y: 1.6, w: 3.15, h: 3.2, size: 16, gap: 14 });
  note(s, "設備2（ETTh2）の全期間 2016年7月〜2018年6月。時刻別の平均を季節ごとに計算。負荷との相関は同時刻の値。データ：Zhou et al. (2021)");
}

// ========== 5. 予測の設定
{
  const s = slide();
  title(s, "主な評価では、予測時点までの観測だけから\n6時間先の油温を予測した");
  const y0 = 2.3, x0 = 0.9, xs = 5.2, xT = 7.7, x1 = 9.1;
  s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y0 - 0.045, w: xs - x0, h: 0.09, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addShape(pres.shapes.RECTANGLE, { x: xs, y: y0 - 0.045, w: x1 - xs, h: 0.09, fill: { color: LIGHT }, line: { color: LIGHT, width: 0 } });
  text(s, "観測済み（油温と負荷）", x0, y0 - 0.5, xs - x0, 0.32, { size: 15, color: NAVY, bold: true, align: "center" });
  text(s, "未来（使うのは時刻と季節だけ）", xs, y0 - 0.5, x1 - xs, 0.32, { size: 15, color: MUTED, bold: true, align: "center" });
  s.addShape(pres.shapes.OVAL, { x: xs - 0.11, y: y0 - 0.11, w: 0.22, h: 0.22, fill: { color: NAVY }, line: { color: WHITE, width: 1.5 } });
  s.addShape(pres.shapes.OVAL, { x: xT - 0.11, y: y0 - 0.11, w: 0.22, h: 0.22, fill: { color: BLUE }, line: { color: WHITE, width: 1.5 } });
  text(s, "予測時点（現在）", xs - 1.0, y0 + 0.2, 2.0, 0.3, { size: 15, color: INK, bold: true, align: "center" });
  text(s, "予測先（6時間後）", xT - 1.3, y0 + 0.2, 2.6, 0.3, { size: 15, color: BLUE, bold: true, align: "center" });
  bullets(s, [
    "主な評価：予測時点までの油温・負荷と、予測先の時刻・季節を使う",
    "参考評価：予測先の負荷が事前に分かると仮定した場合も比べる",
    "1時間先と24時間先の結果は付録Bに載せる",
  ], { y: 3.15, h: 1.8, size: 18, gap: 10 });
}

// ========== 6. 評価の設計
{
  const s = slide();
  title(s, "期間を暦で区切り、モデルの選択は検証期間だけで行った");
  const y = 1.5, h = 0.52, x0 = M, total = 12 + 4 + 7.9, unit = CW / total;
  const segs = [["学習　12か月", "2016年7月〜2017年6月", 12, NAVY], ["検証　4か月", "2017年7月〜10月", 4, "5B8DC0"], ["テスト　約8か月", "2017年11月〜2018年6月", 7.9, BLUE]];
  let x = x0;
  segs.forEach(([t, sub, m, c]) => {
    const w = unit * m;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    text(s, t, x, y, w, h, { size: 14, color: WHITE, bold: true, align: "center", valign: "middle" });
    text(s, sub, x, y + h + 0.06, w, 0.26, { size: 11, color: MUTED, align: "center" });
    x += w;
  });
  bullets(s, [
    "LightGBMを、持続予測・前日同時刻の値・リッジ回帰と比べた",
    "予測時点より後の観測を除いても、\n特徴量が変わらないことを実行のたびに確かめている",
    "テストの結果を見て、モデルや特徴量を選び直していない",
  ], { y: 2.6, h: 2.3, size: 18, gap: 12 });
}

// ========== 7. 結果1 精度
{
  const s = slide();
  const r6 = raw.ETTh2["6"], m6 = mae.ETTh2["6"];
  title(s, `6時間先の誤差は設備2で${f2(m6.lgbm)}℃、\n最良の比較手法より${cut(r6.ridge, r6.lgbm)}%小さい`);
  const ch = barChart(s, {
    x: M, y: 1.4, w: 5.4, h: 3.4, labels: ["持続予測", "前日同時刻", "リッジ回帰", "LightGBM"],
    values: [m6.persistence, m6.seasonal, m6.ridge, m6.lgbm], colors: [GRAY, GRAY, GRAY, BLUE], max: 5.2,
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.78 },
  });
  text(s, "平均絶対誤差（℃）", ch.x, ch.y - 0.05, 2.2, 0.25, { size: 11, color: MUTED });
  const xL = xOf(ch, 3), yTop = yOf(ch, r6.lgbm);
  text(s, "リッジ比", xL - 0.55, yTop + 0.1, 1.1, 0.24, { size: 11, color: WHITE, align: "center" });
  text(s, `−${cut(r6.ridge, r6.lgbm)}%`, xL - 0.55, yTop + 0.32, 1.1, 0.34, { size: 18, color: WHITE, bold: true, align: "center" });
  const r6b = raw.ETTh1["6"], r1 = mae.ETTh2["1"];
  bullets(s, [
    `設備1は、最良の比較手法より\n${cut(r6b.persistence, r6b.lgbm)}%小さいだけ`,
    `設備2の1時間先は、\n持続予測の${f2(r1.persistence)}℃に対して${f2(r1.lgbm)}℃`,
    "24時間先は明確な改善なし（付録B）",
  ], { x: 6.25, y: 1.5, w: 3.25, h: 3.3, size: 16, gap: 14 });
  note(s, `設備2（ETTh2）・6時間先・テスト期間 2017年11月〜2018年6月（${D.n_test_h6.ETTh2.toLocaleString("en-US")}時点）・主な評価の条件。データ：ETT（Zhou et al., 2021）`);
}

// ========== 8. 結果2 安定性
{
  const s = slide();
  title(s, "設備2では、テスト期間の8か月すべてで\n持続予測とリッジ回帰より誤差が小さかった");
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
  bullets(s, [
    `月ごとに見ても、\nリッジ回帰より${lo}〜${hi}%小さい`,
    `週ごとでも${wk.ETTh2.weeks}週中、\n持続予測より${Math.round(wk.ETTh2.vs_persistence * wk.ETTh2.weeks)}週、\nリッジ回帰より${Math.round(wk.ETTh2.vs_ridge * wk.ETTh2.weeks)}週で小さい`,
    `設備1が持続予測より\n小さかったのは${wk.ETTh1.weeks}週中${Math.round(wk.ETTh1.vs_persistence * wk.ETTh1.weeks)}週\n週によって優劣が分かれた`,
  ], { x: 6.4, y: 1.5, w: 3.1, h: 3.3, size: 15, gap: 12 });
  note(s, "設備2・6時間先・主な評価の条件。月と週は予測先の時刻で区切り、それぞれの平均絶対誤差を比べた。データ：ETT（Zhou et al., 2021）");
}

// ========== 9. 結果3 何が効いているか
{
  const s = slide();
  title(s, "今回の比較では、負荷より時刻と季節を足したほうが\n誤差が大きく下がった");
  const ab = D.ablation_h6.ETTh2, ab1 = D.ablation_h6.ETTh1;
  const keys = ["1_ot_only", "2_+calendar", "3_+load_origin (B)", "4_+load_target (A)"];
  const ch = barChart(s, {
    x: M, y: 1.4, w: 5.5, h: 3.4, labels: ["油温の履歴\nだけ", "＋時刻・季節", "＋予測時点\nまでの負荷", "＋予測先の\n実測負荷（参考）"],
    values: keys.map((k) => ab[k]), colors: [GRAY, BLUE, GRAY, GRAY], max: 3.2, catSize: 12,
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.72 },
  });
  text(s, "平均絶対誤差（℃）", ch.x, ch.y - 0.05, 2.2, 0.25, { size: 11, color: MUTED });
  const xm = (xOf(ch, 0) + xOf(ch, 1)) / 2;
  text(s, `−${cut(ab[keys[0]], ab[keys[1]])}%`, xm - 0.2, yOf(ch, ab[keys[0]]) + 0.05, 1.2, 0.34, { size: 18, color: BLUE, bold: true, align: "left" });
  bullets(s, [
    `設備1でも同じ傾向\n（${f2(ab1[keys[0]])}→${f2(ab1[keys[1]])}℃）`,
    `予測時点までの負荷を足すと、\n設備1は約${cut(ab1[keys[1]], ab1[keys[2]])}%改善、\n設備2は悪化`,
    `主な結果の${f2(mae.ETTh2["6"].lgbm)}℃は、\n事前に固定した\n負荷ありの構成`,
    "負荷なしの構成は、\n別の期間で確かめる",
  ], { x: 6.3, y: 1.5, w: 3.2, h: 3.4, size: 16, gap: 12 });
  note(s, "設備2・6時間先・テスト期間。特徴量は左から順に足した。4本目は予測時点では分からない情報を使う参考比較。データ：ETT（Zhou et al., 2021）");
}

// ========== 10. 結果4 高温の事前検知
{
  const s = slide();
  title(s, `設備2では高温${lg.events}区間のうち${lg.detected}区間を事前に検知し、\n誤通知は${lg.false_notices}件だった`);
  const ch = barChart(s, {
    x: M, y: 1.4, w: 5.4, h: 3.4, labels: ["持続予測", "前日同時刻", "リッジ回帰", "LightGBM"],
    values: [det.persistence.detected, det.seasonal.detected, det.ridge.detected, lg.detected], colors: [GRAY, GRAY, GRAY, BLUE], max: 15, fmt: "0",
    layout: { x: 0.02, y: 0.1, w: 0.96, h: 0.78 },
  });
  text(s, "事前に検知できた高温区間の数", ch.x, ch.y - 0.05, 3.2, 0.25, { size: 11, color: MUTED });
  const p = plot(ch), yT = yOf(ch, lg.events);
  s.addShape(pres.shapes.LINE, { x: p.x, y: yT, w: p.w, h: 0, line: { color: ORANGE, width: 1, dashType: "dash" } });
  text(s, `高温区間は全部で${lg.events}`, p.x + 0.05, yT - 0.3, 2.4, 0.26, { size: 12, color: ORANGE, bold: true });
  bullets(s, [
    `通知${lg.notices}件のうち、誤通知は${lg.false_notices}件\n通知と区間は1対1ではない`,
    `検知した${lg.detected}区間では、\n閾値を超える平均${f1(lg.lead_h)}時間前に\n予測が的中`,
    "高温区間は5〜6月に集中",
  ], { x: 6.25, y: 1.5, w: 3.25, h: 3.3, size: 16, gap: 14 });
  note(s, `設備2・6時間先・閾値${f1(thr2)}℃（学習期間の上位5%）。区間は連続した超過、通知は連続した警報をそれぞれ1件と数える。定義は付録D。データ：ETT（Zhou et al., 2021）`);
}

// ========== 11. 結果5 事例
function eventChart(s, ev, { x = M, y = 1.55, w = 6.0, h = 3.25, min = 26, max = 50 } = {}) {
  const ch = lineChart(s, {
    x, y, w, h, labels: ev.hours.map((t) => String(Number(t))), freq: 3, min, max, step: 4,
    series: [{ name: "実測", values: ev.actual, color: INK }, { name: "6時間前の予測", values: ev.pred, color: BLUE }, { name: "閾値", values: ev.hours.map(() => thr2), color: ORANGE }],
    layout: { x: 0.08, y: 0.05, w: 0.7, h: 0.82 }, lineSize: 2,
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
  text(s, "油温（℃）", ch.x, ch.y - 0.2, 1.5, 0.22, { size: 11, color: MUTED });
  text(s, "予測先の時刻", p.x + p.w - 1.4, ch.y + ch.h - 0.02, 1.4, 0.22, { size: 11, color: MUTED, align: "right" });
  return ch;
}
{
  const s = slide();
  const ev = D.event_detected;
  const a = Math.min(...ev.exceed_hours), b = Math.max(...ev.exceed_hours), first = Math.min(...ev.alarm_hours);
  const issued = first - D.horizon_h, lead = a - issued; // 最初の警報を出した時刻と、超過開始までの時間
  title(s, `閾値を超える${lead}時間前に警報を出せたが、\n${a}〜${first - 1}時の超過は6時間前に予測できなかった`);
  const ch = eventChart(s, ev, { min: 30, max: 50 });
  const p = plot(ch), tx = p.x + 0.12, ty = yOf(ch, 49.6), tw = 2.1;
  text(s, `警報を出した時刻：${issued}時\n予測先：${first}時`, tx, ty, tw, 0.46, { size: 12, color: ORANGE, bold: true });
  const x1 = tx + tw - 0.05, y1 = ty + 0.2, x2 = xOf(ch, first) - 0.07, y2 = yOf(ch, ev.pred[first]);
  s.addShape(pres.shapes.LINE, { x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1), flipV: y2 < y1, line: { color: ORANGE, width: 0.75 } });
  text(s, `${a}〜${first - 1}時は\n予測が閾値に届かず`, xOf(ch, a) - 0.05, yOf(ch, ev.pred[a]) + 0.12, 1.4, 0.42, { size: 11, color: BLUE, bold: true });
  bullets(s, [
    `超過は${a}〜${b}時の${b - a + 1}時間`,
    `${issued}時に警報を出した`,
    `${a}〜${first - 1}時の超過は\n予測できなかった`,
  ], { x: 6.9, y: 1.5, w: 2.6, h: 3.3, size: 15, gap: 12 });
  note(s, "設備2・2018年5月15日。横軸は予測先の時刻で、予測はその6時間前に出したもの。橙の点は警報が指す予測先の時刻、橙の帯は実測が閾値を超えた時間。データ：ETT（Zhou et al., 2021）");
}

// ========== 12. 設計上の判断
{
  const s = slide();
  title(s, "1・6・24時間先ごとに油温の変化量を予測し、\n設計は検証期間で決めた");
  table(s, [
    ["判断", "理由"],
    ["現在から予測先までの油温差を予測", "油温の直接予測より検証誤差が小さい"],
    ["1・6・24時間先で別々のモデル", "予測値の再入力による誤差の蓄積を避ける"],
    ["絶対誤差で学習", "評価指標のMAEに合わせる"],
    ["特徴量と閾値はテスト前に固定", "テスト結果を見た選び直しを防ぐ"],
  ], { y: 1.5, colW: [4.1, 4.9], size: 17, rowH: 0.62, bold: [0] });
}

// ========== 13. 限界
{
  const s = slide();
  title(s, "保全効果と、評価していない設備での有効性は、\n今回の検証では確認できていない");
  bullets(s, [
    "故障・点検記録がなく、保全効果は未検証",
    "評価は2台のみ。設備1の改善は小さい",
    "高温は油温の上位5%で定義。危険温度ではない",
    "高温区間は5〜6月に集中。通年の検知性能は未検証",
    "24時間先は、持続予測に対する明確な改善なし",
  ], { y: 1.55, w: 8.4, size: 20, gap: 12 });
}

// ========== 14. 結論と次の一手
{
  const s = slide({ dark: true });
  title(s, "まず現場へ通知せずに予測を試し、\n通知候補の数と、そのうち外れた数を確かめる", { color: WHITE });
  const r6 = raw.ETTh2["6"];
  s.addText([
    { text: `設備2は6時間先の誤差を${cut(r6.ridge, r6.lgbm)}%減らし、高温${lg.events}区間中${lg.detected}区間を事前に検知`, options: { bullet: { indent: 16 }, breakLine: true } },
    { text: "保全効果と、評価していない設備での有効性は未検証", options: { bullet: { indent: 16 } } },
  ], { x: M, y: 1.5, w: CW, h: 1.3, fontFace: FONT, fontSize: 18, color: WHITE, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 12, lineSpacingMultiple: 1.08 });
  text(s, "試行では、通知候補の数・先行時間・見逃しを記録する\n許容できる誤通知の数は、現場と決める", M, 3.0, CW, 0.7, { size: 16, color: TINT });
  text(s, "github.com/HJRKTNG/ett-oil-temperature-poc　｜　沓脱 聖", M, 4.75, 7.5, 0.3, { size: 12, color: TINT });
}

// ========== 15. 参考文献
{
  const s = slide();
  title(s, "参考文献", { size: 24 });
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
  title(s, "主な評価の特徴量は、予測時点までの値と\n予測先の時刻・季節に限った", { size: 22, y: 0.45 });
  table(s, [
    ["油温の履歴", "現在値、1・2・3・6・12・24・48・168時間前の値、現在値と1時間前・24時間前との差、6・24・168時間の移動平均・標準偏差・最大・最小"],
    ["時刻・季節", "予測先の時刻と年内の日（sin・cos）、曜日"],
    ["負荷6種", "現在値、現在値と1時間前との差、24時間平均\n参考評価のみ：予測先の実測負荷も追加"],
    ["LightGBM", "絶対誤差（L1）で学習。学習率0.03、最大葉数31、葉ごとの最小データ数50、特徴量と行の抽出率0.8。検証期間の誤差が100回続けて改善しなければ学習を終了"],
  ], { y: 1.55, colW: [1.8, 7.2], size: 13, rowH: 0.7, head: false, bold: [0] });
}

// ========== 付録B 全結果
{
  const s = slide();
  appendixLabel(s, "付録B　予測する時間ごとの結果");
  title(s, "1時間先は設備2で大きく改善し、\n24時間先は持続予測をほぼ上回らなかった", { size: 22, y: 0.45 });
  const names = [["persistence", "持続予測"], ["seasonal", "前日同時刻"], ["ridge", "リッジ回帰"], ["lgbm", "LightGBM"]];
  const hs = ["1", "6", "24"];
  const line = { pt: 0.75, color: LIGHT }, none = { type: "none" };
  const cell = (t, o = {}) => ({ text: t, options: { fontFace: FONT, fontSize: 13, color: BODY, align: "center", valign: "middle", margin: [3, 2, 3, 2], border: [none, none, line, none], ...o } });
  const head = (t, o = {}) => cell(t, { fontSize: 11, color: MUTED, ...o });
  const rows = [
    [head("平均絶対誤差（℃）\n小さいほど良い", { rowspan: 2, align: "left" }), head("設備1", { colspan: 3, bold: true, color: NAVY }), head("設備2", { colspan: 3, bold: true, color: NAVY })],
    [...hs.map((h) => head(`${h}時間先`)), ...hs.map((h) => head(`${h}時間先`))],
    ...names.map(([k, n]) => [
      cell(n, { align: "left", bold: true, color: NAVY }),
      ...hs.map((h) => cell(f2(mae.ETTh1[h][k]))),
      ...hs.map((h) => {
        const focus = h === "1" && (k === "lgbm" || k === "persistence");
        return cell(f2(mae.ETTh2[h][k]), focus ? { bold: true, color: k === "lgbm" ? BLUE : INK } : {});
      }),
    ]),
  ];
  s.addTable(rows, { x: M, y: 1.45, w: 6.3, colW: [1.62, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78], rowH: 0.42, autoPage: false });
  const r1 = raw.ETTh2["1"];
  bullets(s, [
    `設備2の1時間先は、\n${f2(mae.ETTh2["1"].persistence)}℃から${f2(mae.ETTh2["1"].lgbm)}℃へ（−${cut(r1.persistence, r1.lgbm)}%）`,
    "24時間先のLightGBMは\n持続予測に比べて、\n設備1で悪化、\n設備2でほぼ同じ",
  ], { x: 7.1, y: 1.55, w: 2.4, h: 3.0, size: 14, gap: 12 });
  note(s, "テスト期間・主な評価の条件。データ：ETT（Zhou et al., 2021）");
}

// ========== 付録C 見逃した例
{
  const s = slide();
  const ev = D.event_missed;
  appendixLabel(s, "付録C　見逃した例");
  const a = Math.min(...ev.exceed_hours), b = Math.max(...ev.exceed_hours);
  title(s, `${a}時と${b}時の高温を、6時間前の予測では捉えられなかった`, { size: 22, y: 0.45 });
  const ch = eventChart(s, ev, { min: 26, max: 50 });
  [a, b].forEach((hh) => {
    s.addShape(pres.shapes.OVAL, { x: xOf(ch, hh) - 0.07, y: yOf(ch, ev.pred[hh]) - 0.07, w: 0.14, h: 0.14, fill: { color: WHITE }, line: { color: BLUE, width: 1.5 } });
  });
  const p = plot(ch);
  text(s, `${a}〜${b}時の超過`, xOf(ch, a) - 0.6, p.y + 0.03, 1.4, 0.22, { size: 11, color: ORANGE, bold: true, align: "center" });
  bullets(s, [
    `超過したのは\n${a}時と${b}時の${ev.exceed_hours.length}時点だけ`,
    `白丸は、${a - D.horizon_h}時と${b - D.horizon_h}時に出した\n${a}時・${b}時の予測\n（${f1(ev.pred_by_hour[a])}℃・${f1(ev.pred_by_hour[b])}℃）\nどちらも閾値未満`,
    "次の検証で、\n通知の閾値を下げたときの\n見逃しと誤通知を確かめる",
  ], { x: 6.9, y: 1.55, w: 2.6, h: 3.3, size: 14, gap: 10 });
  note(s, "設備2・2018年6月4日。横軸は予測先の時刻。データ：ETT（Zhou et al., 2021）、予測は本PoCのLightGBM");
}

// ========== 付録D 高温検知の数え方
{
  const s = slide();
  appendixLabel(s, "付録D　高温検知の数え方");
  title(s, "高温の検知は、区間・通知・先行時間を次のように数えた", { size: 22, y: 0.45 });
  table(s, [
    ["高温区間", "実測が閾値を連続して超えた期間を1区間と数える"],
    ["検知", "高温区間の中の時刻に対する予測が閾値を超え、\nその予測を出した時点の実測は閾値以下だった"],
    ["通知", "予測時点で閾値以下のときに出た警報のうち、連続するものを1件と数える"],
    ["誤通知", "通知にまとめた各警報の予測先の時刻で、\n実測が一度も閾値を超えなかった通知"],
    ["先行時間", "検知した区間で、最初に的中した予測を出した時点から、\n実際に超えた時刻まで"],
  ], { y: 1.4, colW: [1.8, 7.2], size: 15, rowH: 0.6, head: false, bold: [0] });
  note(s, `閾値は学習期間の油温の上位5%で固定（設備2は${f1(thr2)}℃）`);
}

// ========== 付録E 高温の基準を変えた場合
{
  const s = slide();
  appendixLabel(s, "付録E　高温の基準を上位10%にした場合");
  const q = lg.q90, pc = (x) => Math.round(x * 100);
  title(s, "高温の基準を上位10%に広げても、\n検知率はほぼ変わらなかった", { size: 22, y: 0.45 });
  table(s, [
    ["設備2・6時間先", `上位5%（${f1(thr2)}℃）`, `上位10%（${f1(D.thresholds.ETTh2.q90)}℃）`],
    ["高温区間", `${lg.events}区間`, `${q.events}区間`],
    ["事前に検知", `${lg.detected}区間（${pc(lg.detected / lg.events)}%）`, `${q.detected}区間（${pc(q.detected / q.events)}%）`],
    ["通知", `${lg.notices}件`, `${q.notices}件`],
    ["1日あたりの通知", `${f2(lg.notices_per_day)}件`, `${f2(q.notices_per_day)}件`],
    ["誤通知", `${lg.false_notices}件`, `${q.false_notices}件`],
    ["先行時間（平均）", `${f1(lg.lead_h)}時間`, `${f1(q.lead_h)}時間`],
  ], { y: 1.45, w: 5.9, colW: [2.1, 1.9, 1.9], size: 14, rowH: 0.44, bold: [0] });
  bullets(s, [
    `検知率は${pc(lg.detected / lg.events)}%と${pc(q.detected / q.events)}%で\nほぼ同じだが、通知は\n${lg.notices}件から${q.notices}件に増えた`,
    `誤通知は${lg.false_notices}件から\n${q.false_notices}件に増えた`,
  ], { x: 6.7, y: 1.55, w: 2.8, h: 3.0, size: 15, gap: 12 });
  note(s, "実測の高温区間と予測による通知の両方に、各列の閾値を当てた。通知の条件だけを緩めた比較ではない。データ：ETT（Zhou et al., 2021）");
}

const out = path.join(OUT_DIR, "ett_oil_temperature_poc.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written", out, "slides", page));
