/* 報告スライドを生成する。文章は docs/slides_text.md（v2）と同期させること。
 * 使い方: NODE_PATH=<pptxgenjs のある node_modules> node scripts/build_slides.js
 * 出力:   reports/slides/ett_oil_temperature_poc.pptx
 */
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const FIG = path.join(ROOT, "reports", "figures");
const OUT_DIR = path.join(ROOT, "reports", "slides");
fs.mkdirSync(OUT_DIR, { recursive: true });

// palette（図と同じ系統）
const INK = "1B1E26", INK2 = "4A4F5C", MUTED = "7C8291", RULE = "D9DCE3", BG = "FFFFFF";
const NAVY = "3F55A8", NAVY_S = "E3E7F5", ORANGE = "B9642A", ORANGE_S = "F5E6DA", GREEN = "2F7D4F", RED = "A33F3F";
const FONT = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.author = "Kutsunugi Hijiri";
pres.title = "変圧器油温の先読み PoC";
const W = 13.33, H = 7.5, M = 0.6;
let pageNo = 0;

// ---------- helpers
function base(opts = {}) {
  const s = pres.addSlide();
  s.background = { color: opts.dark ? "1B2447" : BG };
  pageNo += 1;
  const fc = opts.dark ? "9AA6D6" : MUTED;
  s.addText(String(pageNo), { x: W - M - 0.6, y: H - 0.45, w: 0.6, h: 0.3, fontFace: FONT, fontSize: 10, color: fc, align: "right", margin: 0, isTextBox: true });
  s.addText("ETT 油温予測 PoC", { x: M, y: H - 0.45, w: 4, h: 0.3, fontFace: FONT, fontSize: 9, color: fc, margin: 0, isTextBox: true });
  return s;
}
function kicker(s, text) {
  s.addText(text, { x: M, y: 0.18, w: W - 2 * M, h: 0.28, fontFace: FONT, fontSize: 10, color: NAVY, bold: true, charSpacing: 2, margin: 0, isTextBox: true });
}
function message(s, text, opts = {}) {
  s.addText(text, { x: M, y: 0.45, w: W - 2 * M, h: 1.15, fontFace: FONT, fontSize: opts.size || 19, bold: true, color: INK, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.15 });
}
function bullets(s, items, x, y, w, h, opts = {}) {
  const arr = items.map((t, i) => {
    const last = i === items.length - 1;
    if (typeof t === "string") return { text: t, options: { bullet: { indent: 12 }, breakLine: !last, paraSpaceAfter: opts.gap || 6 } };
    return { text: t.text, options: Object.assign({ bullet: t.head ? false : { indent: 12 }, bold: !!t.head, color: t.head ? INK : INK2, breakLine: !last, paraSpaceAfter: t.head ? 3 : (opts.gap || 6) }, t.options || {}) };
  });
  s.addText(arr, { x, y, w, h, fontFace: FONT, fontSize: opts.size || 13, color: INK2, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.12 });
}
function label(s, text, x, y, w, opts = {}) {
  s.addText(text, { x, y, w, h: 0.3, fontFace: FONT, fontSize: opts.size || 11, bold: true, color: opts.color || MUTED, charSpacing: 1.5, margin: 0, isTextBox: true });
}
function para(s, text, x, y, w, h, opts = {}) {
  s.addText(text, { x, y, w, h, fontFace: FONT, fontSize: opts.size || 12.5, color: opts.color || INK2, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.15 });
}
function caption(s, text, x, y, w, h = 0.3) {
  s.addText(text, { x, y, w, h, fontFace: FONT, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true, valign: "top" });
}
function img(s, file, x, y, w, h) {
  s.addImage({ path: path.join(FIG, file), x, y, w, h, sizing: { type: "contain", w, h } });
}
function table(s, rows, x, y, w, colW, opts = {}) {
  const size = opts.size || 11;
  const data = rows.map((r, ri) => r.map((c) => {
    const isHead = ri === 0;
    const obj = typeof c === "object" && c !== null;
    const text = obj ? c.text : String(c);
    const o = { fontFace: FONT, fontSize: size, color: isHead ? INK : INK2, bold: isHead || (obj && !!c.bold), fill: { color: isHead ? "F0F1F4" : BG }, align: "left", valign: "middle", margin: [3, 5, 3, 5], border: [{ type: "none" }, { type: "none" }, { pt: 0.75, color: RULE }, { type: "none" }] };
    return { text, options: o };
  }));
  s.addTable(data, { x, y, w, colW, rowH: opts.rowH || 0.34, autoPage: false });
}
function box(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: 0.06 });
}
function circleNum(s, n, x, y, d, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 } });
  s.addText(String(n), { x, y, w: d, h: d, fontFace: FONT, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
}

// ============================================================ 表紙
{
  const s = base({ dark: true });
  s.addText("変圧器油温の先読み PoC", { x: M, y: 2.2, w: W - 2 * M, h: 1.0, fontFace: FONT, fontSize: 36, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText("予防保全に向けた技術検証 — ETT 公開データによる検証結果", { x: M, y: 3.2, w: W - 2 * M, h: 0.6, fontFace: FONT, fontSize: 18, color: "CFD6EE", margin: 0, isTextBox: true });
  s.addText("2026 年 9 月　沓脱 聖", { x: M, y: 5.6, w: W - 2 * M, h: 0.4, fontFace: FONT, fontSize: 13, color: "CFD6EE", margin: 0, isTextBox: true });
  s.addText("対象: 変圧器 2 台（設備 1 = ETTh1、設備 2 = ETTh2）・2016-07〜2018-06・1 時間ごとの負荷 6 種と油温 OT ／ 主結果はテスト期間 2017-11-01〜2018-06-26（約 8 か月）、予測起点までの観測だけを使う条件で算出", { x: M, y: 5.95, w: W - 2 * M, h: 0.7, fontFace: FONT, fontSize: 10.5, color: "9AA6D6", margin: 0, isTextBox: true });
}

// ============================================================ 1 結論
{
  const s = base();
  kicker(s, "結論");
  message(s, "設備 2 の 6 時間先予測は、最良ベースラインより平均絶対誤差（MAE）を 18% 改善した。現場実証の候補とするが、設備 1 への有効性と保全効果は未確認である。");
  table(s, [
    ["設備", "6 時間先の MAE（条件 B）", "最良ベースライン比", "高温の事前検知（閾値 = 学習期間の 95 パーセンタイル）"],
    [{ text: "設備 2（ETTh2）", bold: true }, { text: "1.95 ℃", bold: true }, "−18%（リッジ回帰 2.37 ℃）／持続予測比 −56%（4.44 ℃）", "高温 13 区間中 10 区間に事前の的中予測（77%）。通知 12 件・誤通知 3 件（約 8 か月）。検知した区間の先行時間は平均 4.0 時間"],
    [{ text: "設備 1（ETTh1）", bold: true }, "1.11 ℃", "−4%（持続予測 1.15 ℃）", "テスト期間に閾値超えなし → 評価保留"],
  ], M, 1.85, W - 2 * M, [1.9, 2.0, 3.5, 4.73], { size: 11.5, rowH: 0.55 });
  const yb = 3.95;
  box(s, M, yb, 5.85, 2.6, "EEF5F0");
  box(s, M + 6.28, yb, 5.85, 2.6, "F5EEEC");
  label(s, "示せたこと", M + 0.25, yb + 0.15, 5.3, { color: GREEN });
  bullets(s, ["6 時間先の精度改善（設備 2 で最良ベースライン比 −18%、持続予測比 −56%）", "少ない通知数（約 8 か月で 12 件）での高温の事前検知（設備 2）", "効いている情報の特定: 油温の履歴と、対象時刻の時刻・季節"], M + 0.25, yb + 0.55, 5.35, 2.0, { size: 12.5 });
  label(s, "まだ示せないこと", M + 6.53, yb + 0.15, 5.3, { color: RED });
  bullets(s, ["保全効果そのもの（故障・点検・介入の記録がない）", "未学習の設備への汎化（検証は 2 台）と、設備 1 での有効性", "24 時間先の明確な改善、危険温度での検知（閾値は相対的高温の代理）"], M + 6.53, yb + 0.55, 5.35, 2.0, { size: 12.5 });
}

// ============================================================ 2 依頼と問い
{
  const s = base();
  kicker(s, "依頼と問い");
  message(s, "今回判断するのは、故障削減ではなく「油温の先読み精度」と「高温前の通知可能性」である。依頼を、数字で答えられる 3 つの問いに置き換えて検証した。");
  box(s, M, 1.85, 3.7, 4.7, "F0F1F4");
  label(s, "想定する利用場面", M + 0.25, 2.0, 3.2);
  bullets(s, ["実測値の閾値監視に先読みを加え、高温になる前の確認・対応を支援する", "負荷変動・外部要因で温度変動が複雑な設備を対象にする"], M + 0.25, 2.4, 3.2, 1.7, { size: 12 });
  label(s, "データ", M + 0.25, 4.15, 3.2);
  bullets(s, ["変圧器 2 台 × 2 年（2016-07〜2018-06）", "1 時間ごとの負荷 6 種＋油温 OT", "欠損・重複なし（ETT 公開ベンチマーク）"], M + 0.25, 4.55, 3.2, 1.9, { size: 12 });
  const qs = [
    ["問い 1", "何時間先を、どの精度で当てられるか", "指標: MAE。持続予測・季節ナイーブ・リッジ回帰の中で最良のものとの差"],
    ["問い 2", "高温になる前に知らせられるか", "指標: 高温区間の検知率、通知数／日（期間平均）、先行時間"],
    ["問い 3", "どこで外れ、何が効いているか", "指標: 対象月・時刻別の誤差、特徴量群ごとの追加効果"],
  ];
  qs.forEach((q, i) => {
    const y = 1.85 + i * 1.6;
    box(s, M + 4.0, y, W - 2 * M - 4.0, 1.4, NAVY_S);
    s.addText(q[0], { x: M + 4.25, y: y + 0.2, w: 1.2, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(q[1], { x: M + 5.5, y: y + 0.18, w: 6.5, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(q[2], { x: M + 5.5, y: y + 0.7, w: 6.5, h: 0.55, fontFace: FONT, fontSize: 11.5, color: INK2, margin: 0, isTextBox: true });
  });
}

// ============================================================ 3 前提と成功条件
{
  const s = base();
  kicker(s, "前提");
  message(s, "主な予測先は「数時間前に確認・運転調整する」想定で 6 時間先とし、高温の閾値は学習期間の 95 パーセンタイル（上位 5% 点）を代理指標として固定した。");
  label(s, "先に置いた仮定", M, 1.9, 6);
  bullets(s, ["設備ごとに別モデル（2 台を混ぜない）", "運用時も油温は毎時観測できる（欠測時の代替推定は対象外）", "データの到着遅延は 0、再学習なし（固定モデルで約 8 か月を評価）", "閾値は設備の危険温度ではなく「相対的高温」の代理。学習期間の 95 パーセンタイル = 設備 1: 35.2 ℃、設備 2: 46.4 ℃", "予測先は 6 時間先を主評価、1 時間先・24 時間先を補助"], M, 2.3, 5.9, 4.2, { size: 13 });
  box(s, M + 6.4, 1.85, W - 2 * M - 6.4, 4.7, NAVY_S);
  label(s, "事前に置いた成功条件（6 時間先・条件 B）", M + 6.65, 2.0, 5.5, { color: NAVY });
  s.addText("最良ベースラインより MAE を 10% 以上改善", { x: M + 6.65, y: 2.5, w: 5.4, h: 0.45, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, isTextBox: true });
  s.addText("通知 1 件/日以下（期間平均）で、高温区間の検知率 50% 以上", { x: M + 6.65, y: 3.05, w: 5.4, h: 0.45, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, isTextBox: true });
  label(s, "結果", M + 6.65, 3.85, 5.4, { color: NAVY });
  bullets(s, ["設備 2 は今回のテスト期間で両方を満たした（−18%、通知 0.05 件/日で検知率 77%）", "ただし高温 13 区間は 2018 年 5〜6 月に集中しており、通年運用への適合は未確認", "設備 1 は精度改善が小さく（−4%）、テスト期間に閾値超えが無いため警報は評価不能"], M + 6.65, 4.25, 5.4, 2.2, { size: 12.5 });
}

// ============================================================ 4 予測の枠組み
{
  const s = base();
  kicker(s, "予測の枠組み");
  message(s, "予測は起点 s で行い、対象時刻 T = s + h の油温を当てる。「対象時刻の負荷が既知」という条件 A と、起点までの観測だけの条件 B を分けて評価した。");
  const y0 = 2.95, x0 = M + 0.4, x1 = W - M - 0.4, xs = x0 + (x1 - x0) * 0.58, xT = x0 + (x1 - x0) * 0.86;
  s.addShape(pres.ShapeType.rect, { x: x0, y: y0 - 0.02, w: xs - x0, h: 0.12, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addShape(pres.ShapeType.rect, { x: xs, y: y0 - 0.02, w: x1 - xs, h: 0.12, fill: { color: RULE }, line: { color: RULE, width: 0 } });
  s.addText("観測済み（油温・負荷）", { x: x0, y: y0 - 0.55, w: xs - x0, h: 0.4, fontFace: FONT, fontSize: 12, bold: true, color: NAVY, align: "center", margin: 0, isTextBox: true });
  s.addText("未来（未観測。カレンダーだけは既知）", { x: xs, y: y0 - 0.55, w: x1 - xs, h: 0.4, fontFace: FONT, fontSize: 12, bold: true, color: MUTED, align: "center", margin: 0, isTextBox: true });
  s.addShape(pres.ShapeType.ellipse, { x: xs - 0.12, y: y0 - 0.08, w: 0.24, h: 0.24, fill: { color: INK }, line: { color: INK, width: 0 } });
  s.addShape(pres.ShapeType.ellipse, { x: xT - 0.12, y: y0 - 0.08, w: 0.24, h: 0.24, fill: { color: ORANGE }, line: { color: ORANGE, width: 0 } });
  s.addText("予測起点 s", { x: xs - 1.6, y: y0 + 0.3, w: 3.2, h: 0.35, fontFace: FONT, fontSize: 12, bold: true, color: INK, align: "center", margin: 0, isTextBox: true });
  s.addText("予測対象時刻 T = s + h（当てる油温 OT[T]）", { x: xT - 2.1, y: y0 + 0.3, w: 4.2, h: 0.35, fontFace: FONT, fontSize: 12, bold: true, color: ORANGE, align: "center", margin: 0, isTextBox: true });
  s.addText("予測ホライズン h = 1 / 6 / 24 時間", { x: xs, y: y0 + 0.65, w: xT - xs, h: 0.3, fontFace: FONT, fontSize: 10.5, color: MUTED, align: "center", margin: 0, isTextBox: true });
  const yb = 4.3;
  box(s, M, yb, 5.95, 2.35, NAVY_S);
  label(s, "条件 B（運用条件）— 本資料の主結果", M + 0.25, yb + 0.15, 5.5, { color: NAVY });
  bullets(s, ["起点 s までに観測した油温・負荷と、起点で既知の対象時刻 T のカレンダー（時刻・季節・曜日）を使う", "実運用で毎時できる先読みそのもの", "持続予測 = OT[s]、季節ナイーブ = OT[T − 24h] と比較"], M + 0.25, yb + 0.55, 5.5, 1.75, { size: 12 });
  box(s, M + 6.18, yb, 5.95, 2.35, ORANGE_S);
  label(s, "条件 A（課題条件）— 条件付きの参考値", M + 6.43, yb + 0.15, 5.5, { color: ORANGE });
  bullets(s, ["条件 B ＋ 対象時刻 T の負荷 6 種", "課題ルール「t=T の油温を予測する際には t=T の特徴量を使ってよい」の解釈", "将来の負荷が運転計画で既知、という前提付き。実運用ではその保証がない"], M + 6.43, yb + 0.55, 5.5, 1.75, { size: 12 });
}

// ============================================================ 5 データ
{
  const s = base();
  kicker(s, "データ");
  message(s, "油温は滑らかに動き（1 時間自己相関 0.994）、設備 2 は平均的な日内変動が大きい。負荷との相関は弱く、負荷の追加価値は相関ではなく予測誤差の比較で確かめる。");
  img(s, "eda_ETTh2_diurnal.png", M, 1.85, 6.2, 3.25);
  caption(s, "設備 2: 季節別の時刻別平均油温（全期間）", M, 5.1, 6.2);
  table(s, [
    ["", "設備 1", "設備 2"],
    ["油温の平均 / 標準偏差", "13.3 / 8.6 ℃", "26.6 / 11.9 ℃"],
    ["日内振幅（季節別の時刻別平均の最大−最小）", "1.9〜3.1 ℃", "8.1〜11.0 ℃"],
    ["1 時間で 2 ℃を超えて動く割合", "4.4%", "10.3%"],
    ["油温の 1h / 24h 自己相関", "0.994 / 0.941", "0.994 / 0.934"],
    ["負荷 6 種と油温の同時刻相関の最大", "0.22（HULL）", "0.50（MULL）"],
    ["ラグ相関の最大絶対値（HUFL・MUFL・LUFL、0〜48h）", "0.15", "0.19（ラグ 0）"],
  ], M + 6.5, 1.9, W - 2 * M - 6.5, [3.4, 1.15, 1.18], { size: 10.5, rowH: 0.4 });
  bullets(s, ["季節（外気温の代理）と日内周期が支配的 → 対象時刻の時刻・季節を特徴量に入れる価値を検証する", "設備 2 は 6 時間先が最も難しい（持続予測の MAE 4.4 ℃。24 時間先の 3.2 ℃より大きい）", "相関だけでは負荷の追加価値を判断できない → 10 ページのアブレーションで測る（相関図は付録 B）"], M + 6.5, 4.85, W - 2 * M - 6.5, 1.7, { size: 11.5, gap: 4 });
  caption(s, "注記: EDA は全期間で行った。以後のモデル選択は検証期間だけで行ったが、テスト期間を「完全に未見」とは呼ばない。", M, 5.55, 6.2, 0.5);
}

// ============================================================ 6 評価の設計
{
  const s = base();
  kicker(s, "評価の設計");
  message(s, "学習 12 か月・検証 4 か月・テスト約 8 か月を暦で固定し、対象時刻と起点の両方で区切ってリークと境界の問題を防いだ。モデル選択には検証期間だけを使った。");
  const y0 = 1.95, wtot = W - 2 * M;
  const segs = [["学習 2016-07-01 〜 2017-06-30（12 か月）", 12, NAVY], ["検証 2017-07 〜 10（4 か月）", 4, GREEN], ["テスト 2017-11-01 〜 2018-06-26（約 8 か月。6 時間先の評価件数 5,702 時点）", 8, ORANGE]];
  let x = M;
  segs.forEach(([t, m, c]) => { const w = wtot * m / 24; s.addShape(pres.ShapeType.rect, { x, y: y0, w, h: 0.6, fill: { color: c }, line: { color: BG, width: 1 } }); s.addText(t, { x, y: y0, w, h: 0.6, fontFace: FONT, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true }); x += w; });
  bullets(s, [{ text: "行の切り方", head: true }, "学習に使う行: 対象時刻 T が学習終了前（ラベルが検証期間に食い込まない）", "検証・テストの評価行: 対象時刻 T と起点 s の両方がその期間内（各期間の先頭 h 時間は除外）", { text: "リーク防止", head: true }, "起点までの観測から作る特徴量は、未来のデータを切り落としても値が変わらないことをテストで確認", "標準化の平均・標準偏差は学習期間だけで推定。早期終了と、油温差／油温そのものの予測方式の選択には検証期間だけを使う。テスト期間の結果でモデルを再選択しない（EDA は全期間）"], M, 2.85, 6.4, 3.7, { size: 13 });
  bullets(s, [{ text: "比較する手法（簡単な順）", head: true }, "B0 持続予測 OT[s]", "B1 季節ナイーブ OT[T − 24h]", "B2 リッジ回帰（条件 B の特徴量）", "M1 LightGBM（L1 目的関数。油温差 OT[T] − OT[s] を予測し OT[s] に足し戻す。予測先ごとに別モデル）", { text: "指標", head: true }, "MAE（主）、RMSE、高温区間の検知率、通知数／日、先行時間"], M + 6.8, 2.85, W - 2 * M - 6.8, 3.7, { size: 13 });
}

// ============================================================ 7 精度
{
  const s = base();
  kicker(s, "結果 1 — 精度");
  message(s, "6 時間先の MAE は設備 2 で 4.44（持続）→ 2.37（リッジ）→ 1.95 ℃（LightGBM）。設備 1 は 1.15 → 1.11 ℃で、改善は小さい。");
  img(s, "res_mae_by_horizon.png", M, 1.8, W - 2 * M, 3.0);
  table(s, [
    ["MAE ℃: 持続 / リッジ / LightGBM", "1 時間先", "6 時間先", "24 時間先"],
    ["設備 1（ETTh1）", "0.44 / 0.46 / 0.42", { text: "1.15 / 1.16 / 1.11", bold: true }, "1.63 / 1.93 / 1.80"],
    ["設備 2（ETTh2）", "0.91 / 0.50 / 0.37", { text: "4.44 / 2.37 / 1.95", bold: true }, "3.19 / 3.26 / 3.17"],
  ], M, 4.95, 7.4, [3.2, 1.4, 1.4, 1.4], { size: 10.5, rowH: 0.34 });
  caption(s, "テスト期間・条件 B。季節ナイーブの 6 時間先 MAE は設備 1 で 1.64 ℃、設備 2 で 3.21 ℃。最良ベースラインは設備 1 が持続予測、設備 2 がリッジ回帰。評価件数は各 5,684〜5,707 時点。", M, 6.12, 7.4, 0.5);
  bullets(s, ["1 時間先: 設備 2 は持続予測比 −59%（0.91 → 0.37）。業務上の価値は 1 時間の猶予で可能な対応に依存する。設備 1 は直前値でほぼ足りる", "24 時間先: 今回比較したモデル・特徴量では、持続予測を明確に上回る改善は確認できなかった（設備 1 は悪化、設備 2 は 3.19 → 3.17）"], M + 7.7, 4.95, W - 2 * M - 7.7, 1.7, { size: 11.5, gap: 4 });
}

// ============================================================ 8 安定性
{
  const s = base();
  kicker(s, "結果 2 — 安定性");
  message(s, "設備 2 の 6 時間先の改善は 8 か月すべてで成立した（持続比 28〜64%、リッジ比 3〜28%）。設備 1 は週単位で見ると持続に勝つ週が 54% しかない。");
  img(s, "res_error_by_month_hour.png", M, 1.8, W - 2 * M, 3.3);
  bullets(s, ["設備 2: 週単位で 97% の週で持続に、91% の週でリッジに勝つ", "設備 1: 月によって負ける（2018-01 は持続 0.94 ℃に対し 1.26 ℃）。3.7% の改善を安定した効果とは言わない", "対象時刻別: 昼〜午後で持続予測の誤差が大きく 14 時で最大（9.5 ℃）。LightGBM との差も 14 時で最大（3.8 ℃）"], M, 5.25, W - 2 * M, 1.3, { size: 12.5, gap: 4 });
}

// ============================================================ 9 何が効いているか
{
  const s = base();
  kicker(s, "結果 3 — 何が効いているか");
  message(s, "油温の履歴に時刻・季節を足すと誤差が大きく下がる（設備 2: 2.68 → 1.83 ℃）。起点までの負荷を足しても設備 1 はわずかな改善、設備 2 は悪化した。");
  img(s, "res_ablation_h6.png", M, 1.8, 6.4, 2.7);
  img(s, "res_importance_ETTh2_h6.png", M + 6.7, 1.8, 5.5, 3.35);
  table(s, [
    ["6 時間先 MAE ℃（テスト）", "油温の履歴のみ", "＋時刻・季節", "＋起点までの負荷（B）", "＋対象時刻の負荷（A）"],
    ["設備 1", "1.523", "1.123", "1.108", "1.084"],
    ["設備 2", "2.676", { text: "1.827", bold: true }, "1.951", "1.968"],
  ], M, 4.65, 6.4, [1.5, 1.2, 1.1, 1.3, 1.3], { size: 10.5, rowH: 0.34 });
  bullets(s, ["主結果の 1.95 ℃は事前に固定した全特徴量（条件 B）の値で、最良の構成を検証で選んだものではない", "設備 2 では検証期間でも負荷なし構成が良かった（1.74 vs 1.78 ℃）。次回はこの構成を候補にし、新しい評価期間で警報性能も確認する", "「対象時刻の負荷が既知」という条件 A でも改善は小さい（設備 1 で −2%、設備 2 で悪化）"], M + 6.7, 5.25, 5.5, 1.35, { size: 10.5, gap: 3 });
}

// ============================================================ 10 事前検知
{
  const s = base();
  kicker(s, "結果 4 — 高温の事前検知");
  message(s, "設備 2・6 時間先で、高温 13 区間のうち 10 区間に事前の的中予測があり、通知は約 8 か月で 12 件（誤通知 3 件）。検知した区間の先行時間は平均 4.0 時間だった。");
  table(s, [
    ["指標（設備 2・6 時間先・条件 B・閾値 46.4 ℃）", "値"],
    ["閾値を超えた時点数 / 高温区間（連続した超過 = 1 区間）", "65 時点 / 13 区間"],
    ["事前の的中予測があった区間", { text: "10 / 13（77%）", bold: true }],
    ["通知数（連続した事前警報 = 1 通知）／うち誤通知", "12 件（0.05 件/日）／ 3 件"],
    ["先行時間（検知した区間の平均）", { text: "4.0 時間", bold: true }],
    ["高温状態を時点単位で当てた割合（参考）", "42%（27 / 65）"],
  ], M, 1.85, 6.6, [4.5, 2.1], { size: 11, rowH: 0.42 });
  label(s, "警報のベースライン比較（同じ区間・閾値）", M, 4.5, 6.6);
  table(s, [
    ["手法", "検知した区間", "通知数／日", "誤通知／日"],
    ["持続予測", "0 / 13", "0.00", "0.00"],
    ["季節ナイーブ", "6 / 13", "0.05", "0.02"],
    ["リッジ回帰", "5 / 13", "0.03", "0.00"],
    [{ text: "LightGBM", bold: true }, { text: "10 / 13", bold: true }, "0.05", "0.01"],
  ], M, 4.85, 6.6, [1.9, 1.5, 1.6, 1.6], { size: 10.5, rowH: 0.32 });
  box(s, M + 7.0, 1.85, W - 2 * M - 7.0, 4.7, "F0F1F4");
  label(s, "定義と注意", M + 7.25, 2.0, 4.8);
  bullets(s, ["対象は「起点では閾値以下、対象時刻で超える」区間。連続した実測超過は 1 区間、連続した事前警報は 1 通知に集約", "先行時間は、検知した各区間について「最初の的中予測を出した起点」から「実際の超過開始」までを測る。集約した通知の開始時刻からではない", "的中した通知 9 件と検知した区間 10 件は別の集計単位（1 通知が複数の区間に対応しうる）。誤通知は、発報した起点の 6 時間後の対象時刻で実測が一度も閾値を超えなかった通知", "区間は 2018 年 5〜6 月に集中し、同日の再超過を含む", "設備 1 はテスト期間に超過 0 件のため評価保留。上位 10% 点での補助評価は付録 B"], M + 7.25, 2.4, 4.85, 4.1, { size: 10.5, gap: 4 });
}

// ============================================================ 11 成功例と失敗例
{
  const s = base();
  kicker(s, "結果 5 — 成功例と失敗例");
  message(s, "超過開始の 3 時間前に警報を出せた例がある一方、2 時点だけの短い高温区間は予測が閾値に届かず見逃した。");
  img(s, "res_alarm_timeline_detected.png", M, 1.75, 7.6, 2.35);
  img(s, "res_alarm_timeline_missed.png", M, 4.2, 7.6, 2.35);
  bullets(s, [{ text: "成功例（5 月 15 日）", head: true }, "12〜18 時に閾値を超過。最初の的中予測は 9 時起点の 15 時予測で、超過開始の 3 時間前に警報を出せた。ただし 12〜14 時の高温は予測できなかった（6〜8 時起点の予測は 41.5・41.5・46.0 ℃で閾値未満）", { text: "失敗例（6 月 4 日）", head: true }, "14・15 時の 2 時点で超過（実測 46.5・46.9 ℃）。6 時間前の予測は 40.0・42.5 ℃で、どちらも警報を出せなかった", { text: "含意", head: true }, "この例では短い高温区間を過小予測した。通知閾値の引き下げで拾える可能性があるが、検知率と誤通知への影響は別途評価する"], M + 7.9, 1.85, W - 2 * M - 7.9, 4.7, { size: 11, gap: 4 });
}

// ============================================================ 12 判断と限界
{
  const s = base();
  kicker(s, "設計上の判断と限界");
  message(s, "予測先ごとの油温差の予測、L1 目的関数、検証期間だけでの選択を採った。24 時間先、未学習設備への汎化、保全効果は今回の範囲では示せない。");
  box(s, M, 1.85, 5.95, 4.7, NAVY_S);
  label(s, "判断（理由）", M + 0.25, 2.0, 5.5, { color: NAVY });
  bullets(s, ["各予測先のモデルで油温差 OT[T] − OT[s] を予測し、現在値 OT[s] に足し戻す — 検証で油温そのものの予測より良かった", "将来予測を次の入力に使う再帰方式は採用しない — 誤差の伝播を避ける", "L1 目的関数 — MAE を主指標にするため。木モデルの外挿の弱さは明記", "特徴量は事前に固定。モデル選択には検証期間を使い、テスト期間の結果で再選択しない", "閾値は学習期間で固定し、区間数を必ず併記"], M + 0.25, 2.4, 5.5, 4.0, { size: 13, gap: 9 });
  box(s, M + 6.18, 1.85, 5.95, 4.7, "F5EEEC");
  label(s, "限界", M + 6.43, 2.0, 5.5, { color: RED });
  bullets(s, ["24 時間先は、今回比較したモデル・特徴量では持続予測を明確に上回る改善を確認できなかった。外気温・気象予報の追加は次に検証する候補", "設備 2 台のみ。未学習設備への汎化は未検証", "高温区間が 5〜6 月に集中し、閾値は相対値。危険温度での検知は未評価", "故障・点検・介入の記録が無く、保全効果は評価できない", "短い高温区間を過小予測した例がある（12 ページ）"], M + 6.43, 2.4, 5.5, 4.0, { size: 13, gap: 9 });
}

// ============================================================ 13 次の一手
{
  const s = base();
  kicker(s, "現場実証に向けて");
  message(s, "まず対象設備で 6 時間先予測をシャドー運用し、通知負担と対応可能な先行時間を確認する。24 時間先の改善は別の技術検証として進める。");
  label(s, "現場実証に必要なもの", M, 1.9, 6);
  bullets(s, ["対象設備の毎時の油温・負荷（欠測・遅延の実態込み）", "閾値の定義 — 危険温度・運転上限・保全基準のどれか", "警報 → 担当者の確認 → 点検・運転調整、の流れと責任者", "評価に使える記録 — 点検・介入・異常の履歴"], M, 2.3, 5.9, 2.4, { size: 13 });
  box(s, M, 4.8, 5.9, 1.7, NAVY_S);
  label(s, "実証前に現場と合意すること", M + 0.25, 4.95, 5.4, { color: NAVY });
  para(s, "必要な最低先行時間、許容できる誤通知数、見逃し時の扱い、評価期間。", M + 0.25, 5.35, 5.4, 1.0, { size: 13 });
  const steps = [["1", "6 時間先予測のシャドー運用", "対象設備で通知を出さずに予測を走らせ、通知数・先行時間・誤通知を実データで確認"], ["2", "設備別の再学習と監視", "再学習の頻度、ドリフトの検知、誤差の時系列監視。通知閾値は許容する通知数から調整"], ["3", "外気温・気象予報の追加（別の技術検証）", "24 時間先の改善候補。データ取得の目処と合わせて評価"]];
  label(s, "技術側の次の一手（優先順）", M + 6.4, 1.9, 6, { color: NAVY });
  steps.forEach(([n, t, d], i) => {
    const y = 2.35 + i * 1.4;
    circleNum(s, n, M + 6.4, y + 0.05, 0.5, NAVY);
    s.addText(t, { x: M + 7.05, y, w: 5.1, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(d, { x: M + 7.05, y: y + 0.42, w: 5.1, h: 0.8, fontFace: FONT, fontSize: 11.5, color: INK2, margin: 0, isTextBox: true, valign: "top" });
  });
}

// ============================================================ 付録 A
{
  const s = base();
  kicker(s, "付録 A");
  message(s, "特徴量・モデル設定・数表・出典", { size: 18 });
  bullets(s, [{ text: "A1 特徴量（条件 B。油温・負荷の履歴は予測起点 s 基準、カレンダーは対象時刻 T 基準）", head: true }, "油温: 現在値、ラグ 1/2/3/6/12/24/48/168h、差分 1h/24h、移動平均・標準偏差・最大・最小（窓 6/24/168h）", "カレンダー: 対象時刻 T の時刻・年内日（sin/cos）、曜日", "負荷 6 種: 現在値、1h 差分、24h 平均（条件 A はさらに対象時刻 T の負荷 6 種）", { text: "A2 LightGBM", head: true }, "L1 目的関数、学習率 0.03、葉 31、min_data_in_leaf 50、feature/bagging fraction 0.8、早期終了 100（検証 MAE）、乱数固定", { text: "用語", head: true }, "予測ホライズン h = 起点から対象時刻までの時間。高温区間 = 連続した実測超過。通知 = 連続した事前警報。時点 = 毎時の観測"], M, 1.7, 6.2, 4.8, { size: 11 });
  bullets(s, [{ text: "A3 数表・コード", head: true }, "reports/results_baselines.md、results_lgbm.md、results_stability.md、results_ablation.csv", "GitHub: github.com/HJRKTNG/ett-oil-temperature-poc（再現手順は README）", { text: "A4 出典", head: true }, "Zhou, H. et al. Informer: Beyond Efficient Transformer for Long Sequence Time-Series Forecasting. AAAI 2021. データ: github.com/zhouhaoyi/ETDataset", "Ke, G. et al. LightGBM: A Highly Efficient Gradient Boosting Decision Tree. NeurIPS 2017"], M + 6.6, 1.7, W - 2 * M - 6.6, 4.8, { size: 11 });
}

// ============================================================ 付録 B
{
  const s = base();
  kicker(s, "付録 B");
  message(s, "補助評価: 上位 10% 点での警報性能と、負荷と油温のラグ相関", { size: 18 });
  label(s, "B1 設備 2・6 時間先・条件 B。閾値 = 学習期間の 90 パーセンタイル（43.1 ℃）vs 95（46.4 ℃）", M, 1.7, 6.4, { size: 10 });
  table(s, [
    ["指標", "95 パーセンタイル", "90 パーセンタイル"],
    ["超過時点 / 高温区間", "65 / 13", "147 / 25"],
    ["事前の的中予測があった区間", "10 / 13（77%）", "19 / 25（76%）"],
    ["通知数（／日）", "12（0.05）", "24（0.10）"],
    ["誤通知", "3", "5"],
    ["先行時間（平均）", "4.0 時間", "4.6 時間"],
  ], M, 2.1, 6.2, [2.6, 1.8, 1.8], { size: 10.5, rowH: 0.36 });
  caption(s, "同じ高温定義で通知閾値だけを下げた比較ではない（高温の定義そのものを変えている）。", M, 4.35, 6.2);
  img(s, "eda_ETTh2_xcorr.png", M + 6.6, 1.7, 5.5, 2.9);
  caption(s, "B2 設備 2: 負荷を先行させた油温との相関（HUFL・MUFL・LUFL の 3 種、ラグ 0〜48h、全期間）。最大絶対値は MUFL のラグ 0 で 0.19。同時刻相関が最大の MULL（0.50）はこの図に含まない。", M + 6.6, 4.65, 5.5, 0.8);
}

const out = path.join(OUT_DIR, "ett_oil_temperature_poc.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written", out, "slides", pageNo));
