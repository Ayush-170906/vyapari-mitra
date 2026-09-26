const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "0B2447";
const NAVY2 = "13315C";
const CYAN = "00BAF2";
const CYAN_D = "0090C4";
const SAFFRON = "F2A93B";
const INK = "1C2733";
const MUTED = "6B7686";
const WHITE = "FFFFFF";
const LIGHT = "F4F7FB";
const IVORY = "FAF6EF";
const GREEN = "1FAA6B";

function freshShadow() {
  return { type: "outer", color: "0B2447", opacity: 0.18, blur: 8, offset: 3, angle: 90 };
}

function circleIcon(slide, x, y, d, glyph, bg, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  slide.addText(glyph, {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontSize: d * 30, color: fg, bold: true, fontFace: "Arial", isTextBox: true, margin: 0
  });
}

// A small rotated "sticky note" annotation — reads as a hand-added call-out
// rather than a templated bullet, which is the whole point of it.
function stickyTag(slide, x, y, w, h, text, angle, bg, fg) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.06, fill: { color: bg }, line: { type: "none" },
    rotate: angle, shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 4, offset: 2, angle: 90 }
  });
  slide.addText(text, {
    x, y, w, h, align: "center", valign: "middle", fontSize: 11, bold: true, italic: true,
    color: fg, fontFace: "Bookman Old Style", rotate: angle, isTextBox: true, margin: 0
  });
}

// A straight connector with an arrowhead, for architecture/flow diagrams.
function arrow(slide, x1, y1, x2, y2, color, dashed) {
  slide.addShape("line", {
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(x2 - x1) || 0.01, h: Math.abs(y2 - y1) || 0.01,
    flipV: y2 < y1, flipH: x2 < x1,
    line: { color, width: 1.75, endArrowType: "triangle", dashType: dashed ? "dash" : "solid" }
  });
}

// A labelled box for architecture diagrams.
function archBox(slide, x, y, w, h, label, sub, bg, fg, subColor) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.08, fill: { color: bg }, line: { type: "none" }, shadow: freshShadow() });
  slide.addText(label, {
    x: x + 0.1, y: y, w: w - 0.2, h: sub ? h * 0.58 : h, align: "center", valign: sub ? "bottom" : "middle",
    fontSize: 11.5, bold: true, color: fg, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.1, y: y + h * 0.56, w: w - 0.2, h: h * 0.42, align: "center", valign: "top",
      fontSize: 8.5, color: subColor || fg, fontFace: "Calibri", isTextBox: true, margin: 0
    });
  }
}

const LOGO_PATH = path.join(__dirname, "paytm_logo_clean.png");
const LOGO_RATIO = 505 / 1603; // h/w of the source logo

// Places the Paytm wordmark. On a dark background it sits on a small white
// card so both the navy and cyan strokes stay legible.
function addLogo(slide, x, y, w, onDark) {
  const h = w * LOGO_RATIO;
  if (onDark) {
    const padX = w * 0.16, padY = h * 0.55;
    slide.addShape("roundRect", {
      x: x - padX, y: y - padY, w: w + padX * 2, h: h + padY * 2,
      rectRadius: 0.06, fill: { color: WHITE }, line: { type: "none" }
    });
  }
  slide.addImage({ path: LOGO_PATH, x, y, w, h });
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
const W = 13.33, H = 7.5;

pres.defineSlideMaster({ title: "BLANK", background: { color: WHITE }, objects: [] });

// ---------------- Slide 1: Title ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  // Pure Paytm-brand accents only — pale cyan washes, no invented navy tint
  s.addShape("ellipse", { x: 9.4, y: -2.8, w: 7, h: 7, fill: { color: "E8F8FE" }, line: { type: "none" } });
  s.addShape("ellipse", { x: -2.6, y: 4.7, w: 5.6, h: 5.6, fill: { color: "E8F8FE" }, line: { type: "none" } });

  addLogo(s, 11.0, 0.7, 1.55, false);

  s.addText("PAYTM BUILD FOR INDIA AI HACKATHON  ·  MUMBAI EDITION", {
    x: 0.7, y: 0.6, w: 9.5, h: 0.4, fontSize: 12, color: CYAN_D, bold: true, charSpacing: 2,
    fontFace: "Calibri", isTextBox: true, margin: 0
  });

  s.addText("VYAPARI MITRA", {
    x: 0.7, y: 2.5, w: 11.9, h: 1.4, fontSize: 58, color: NAVY, bold: true,
    fontFace: "Times New Roman", isTextBox: true, margin: 0
  });
  s.addText("India's Vernacular AI Copilot for Every Paytm Merchant", {
    x: 0.7, y: 3.85, w: 11.2, h: 0.7, fontSize: 22, color: MUTED, italic: true,
    fontFace: "Calibri", isTextBox: true, margin: 0
  });

  s.addShape("roundRect", { x: 0.7, y: 4.85, w: 3.0, h: 0.55, rectRadius: 0.1, fill: { color: CYAN }, line: { type: "none" } });
  s.addText("Track: Merchant Growth AI", {
    x: 0.7, y: 4.85, w: 3.0, h: 0.55, align: "center", valign: "middle", fontSize: 13, bold: true,
    color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  s.addShape("roundRect", { x: 3.9, y: 4.85, w: 2.75, h: 0.55, rectRadius: 0.1, fill: { color: SAFFRON }, line: { type: "none" } });
  s.addText("Powered by Sarvam AI", {
    x: 3.9, y: 4.85, w: 2.75, h: 0.55, align: "center", valign: "middle", fontSize: 12.5, bold: true,
    color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  s.addShape("roundRect", { x: 6.85, y: 4.85, w: 2.55, h: 0.55, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("+ Fraud Shield", {
    x: 6.85, y: 4.85, w: 2.55, h: 0.55, align: "center", valign: "middle", fontSize: 12.5, bold: true,
    color: WHITE, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  s.addText("Team: The Vision   |   Ayush Korde  ·  Harshita Girase", {
    x: 0.7, y: 6.55, w: 8, h: 0.4, fontSize: 13, color: MUTED, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  s.addText("Oct 3, 2026", {
    x: 10.5, y: 6.55, w: 2.1, h: 0.4, align: "right", fontSize: 13, color: MUTED, fontFace: "Calibri", isTextBox: true, margin: 0
  });
}

// ---------------- Slide 2: Problem Statement ----------------
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  s.addText("Problem Statement", { x: 0.7, y: 0.5, w: 11.9, h: 0.7, fontSize: 31, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);
  s.addText("India's Paytm merchants get payments infrastructure — but not a growth partner", {
    x: 0.7, y: 1.15, w: 11.5, h: 0.5, fontSize: 15, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  const items = [
    ["H", "Language barrier", "Most kirana/local merchants think and decide in Hindi or a regional language — not English dashboards or app menus."],
    ["?", "No business insight", "Sales, top products, and complaints sit in raw transaction data merchants never actually see or understand."],
    ["T", "No time to analyze", "A shopkeeper running the counter all day has no time to read analytics — they need a 3-second spoken answer."],
    ["!", "Fake-payment fraud", "RBI FY25: 13,516 digital-payment fraud cases worth ₹520 crore; NPCI projects UPI fraud incidents to cross 1.1M — QR-swap and fake-screenshot scams hit small merchants daily."]
  ];
  const colors = [CYAN, SAFFRON, GREEN, "C0392B"];
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.7 + col * 6.0, y = 2.05 + row * 2.35;
    s.addShape("roundRect", { x, y, w: 5.7, h: 2.1, rectRadius: 0.12, fill: { color: WHITE }, line: { type: "none" }, shadow: freshShadow() });
    circleIcon(s, x + 0.35, y + 0.35, 0.75, it[0], colors[i], WHITE);
    s.addText(it[1], { x: x + 1.3, y: y + 0.3, w: 4.0, h: 0.5, fontSize: 16, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(it[2], { x: x + 0.35, y: y + 1.15, w: 5.05, h: 0.85, fontSize: 11, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top" });
  });
  stickyTag(s, 10.9, 4.15, 1.7, 0.55, "the real gap", 5, "C0392B", WHITE);
}

// ---------------- Slide 3: Proposed Solution ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Proposed Solution", { x: 0.7, y: 0.5, w: 11.9, h: 0.7, fontSize: 34, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });

  s.addShape("roundRect", { x: 0.7, y: 1.4, w: 7.0, h: 5.4, rectRadius: 0.14, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("An AI business partner that talks to merchants\nin their own language", {
    x: 1.1, y: 1.75, w: 6.2, h: 1.0, fontSize: 21, bold: true, color: WHITE, fontFace: "Times New Roman", isTextBox: true, margin: 0
  });

  const flow = [
    ["1", "Merchant asks", "\"Is hafte ka sales kaisa raha?\" — in Hindi/Hinglish, by voice or text"],
    ["2", "Sarvam AI understands", "Native Indian-language LLM reads intent + merchant's own sales/product/rating data"],
    ["3", "Vyapari Mitra replies", "Plain-language insight + ONE concrete action — never just a chart or number"],
    ["4", "Merchant acts", "One-tap suggestions: launch cashback, bundle a product, fix a complaint driver"]
  ];
  flow.forEach((f, i) => {
    const y = 2.95 + i * 0.9;
    s.addShape("ellipse", { x: 1.1, y: y, w: 0.55, h: 0.55, fill: { color: CYAN }, line: { type: "none" } });
    s.addText(f[0], { x: 1.1, y: y, w: 0.55, h: 0.55, align: "center", valign: "middle", fontSize: 16, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(f[1], { x: 1.85, y: y - 0.05, w: 5.55, h: 0.4, fontSize: 14.5, bold: true, color: WHITE, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(f[2], { x: 1.85, y: y + 0.32, w: 5.6, h: 0.55, fontSize: 11, color: "BFD3F0", fontFace: "Calibri", isTextBox: true, margin: 0 });
  });

  // Right: phone mockup screenshot
  s.addShape("roundRect", { x: 8.15, y: 1.15, w: 4.55, h: 6.0, rectRadius: 0.28, fill: { color: NAVY2 }, line: { type: "none" }, shadow: freshShadow() });
  s.addImage({ path: path.join(__dirname, "demo-shot.png"), x: 8.35, y: 1.35, w: 4.15, h: 5.6, sizing: { type: "crop", w: 4.15, h: 5.6 } });
}

// ---------------- Slide 3b: Our Moat — Fraud Shield ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Our Moat: Real-Time Fraud Shield", { x: 0.7, y: 0.5, w: 9.6, h: 0.7, fontSize: 28, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);
  stickyTag(s, 8.55, 1.2, 2.0, 0.55, "nobody else builds this", -4, "C0392B", WHITE);
  s.addText("A feature every other 'merchant insights' bot skips — because it protects money, not just gives advice", {
    x: 0.7, y: 1.15, w: 7.5, h: 0.45, fontSize: 13, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  // Left: the real-world scenario, backed by cited data
  s.addShape("roundRect", { x: 0.7, y: 1.95, w: 5.9, h: 4.9, rectRadius: 0.14, fill: { color: "FBEAEA" }, line: { type: "none" } });
  s.addText("The scenario Paytm already fights every day", {
    x: 1.0, y: 2.2, w: 5.35, h: 0.4, fontSize: 15, bold: true, color: "C0392B", fontFace: "Calibri", isTextBox: true, margin: 0
  });
  const facts = [
    "Fraudsters paste a fake QR sticker over a merchant's real QR — every scan pays the attacker, not the shop.",
    "Or a customer shows a fake “payment successful” screenshot from a modified app and walks off with the goods.",
    "RBI FY25: 13,516 digital-payment fraud cases worth ₹520 crore. NPCI projects UPI fraud incidents to cross 1.1M.",
    "Sep 2026, Basti (UP): two arrested for exactly this — fake UPI screenshots, cash taken, no real payment made."
  ];
  facts.forEach((f, i) => {
    const y = 2.75 + i * 0.98;
    circleIcon(s, 1.0, y, 0.4, "!", "C0392B", WHITE);
    s.addText(f, { x: 1.55, y: y - 0.08, w: 4.85, h: 0.95, fontSize: 10.5, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top" });
  });

  // Right: how Vyapari Mitra closes the gap, with a real screenshot as proof
  s.addShape("roundRect", { x: 6.85, y: 1.95, w: 3.55, h: 4.9, rectRadius: 0.14, fill: { color: NAVY }, line: { type: "none" }, shadow: freshShadow() });
  s.addText("How Vyapari Mitra closes it", {
    x: 7.1, y: 2.15, w: 3.1, h: 0.65, fontSize: 13.5, bold: true, color: SAFFRON, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  const closes = [
    ["Voice-verify any screenshot", "Cross-checks the real Paytm settlement ledger in real time, in Hindi."],
    ["Catch silent QR-swaps", "Flags it when Soundbox stays quiet while the merchant says sales are happening."],
    ["Zero new hardware", "Runs on the Soundbox + settlement data Paytm already has."]
  ];
  closes.forEach((c, i) => {
    const y = 2.85 + i * 1.28;
    circleIcon(s, 7.1, y, 0.36, "✓", CYAN, NAVY);
    s.addText(c[0], { x: 7.55, y: y - 0.06, w: 2.7, h: 0.32, fontSize: 11, bold: true, color: WHITE, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(c[1], { x: 7.55, y: y + 0.26, w: 2.75, h: 0.85, fontSize: 9, color: "C9D6EC", fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top" });
  });

  // Real screenshot: this exact scenario, actually working
  s.addShape("roundRect", { x: 10.6, y: 1.95, w: 2.15, h: 4.9, rectRadius: 0.2, fill: { color: NAVY2 }, line: { type: "none" }, shadow: freshShadow() });
  s.addImage({ path: path.join(__dirname, "demo-fraud-proof.png"), x: 10.73, y: 2.08, w: 1.89, h: 1.94 });
  s.addText("Actual output from the working prototype", {
    x: 10.73, y: 4.1, w: 1.89, h: 0.65, fontSize: 8.5, italic: true, color: "9FB6D9", fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top"
  });
}

// ---------------- Slide 4: How It Works / Live Demo ----------------
{
  const s = pres.addSlide();
  s.background = { color: IVORY };
  s.addText("Live Demo — Vyapari Mitra in Action", { x: 0.7, y: 0.5, w: 11.9, h: 0.7, fontSize: 27, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  s.addText("Working prototype · web app · Sarvam-105B model", {
    x: 0.7, y: 1.15, w: 8, h: 0.4, fontSize: 13, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  s.addShape("roundRect", { x: 8.55, y: 1.1, w: 4.1, h: 5.85, rectRadius: 0.28, fill: { color: NAVY }, line: { type: "none" }, shadow: freshShadow() });
  s.addImage({ path: path.join(__dirname, "demo-shot.png"), x: 8.75, y: 1.3, w: 3.7, h: 5.5, sizing: { type: "crop", w: 3.7, h: 5.5 } });

  const steps = [
    ["Merchant ka sawaal", "“Is hafte ka sales kaisa raha, detail mein batao” — typed in Hinglish, exactly like a shopkeeper would ask a friend."],
    ["Merchant ka apna context", "Sarvam AI ko un's sales (₹1.82L, +14%), top products, aur ratings pehle se pata hain — koi generic jawab nahi."],
    ["Action-first jawab", "Sirf number nahi — ek specific suggestion: 'Weekend pe Extra Cashback banner Paytm QR ke paas laga dein.'"],
    ["Ek-tap follow-up", "Quick-suggestion chips se merchant turant agla sawaal pooch sakta hai — koi typing friction nahi."]
  ];
  steps.forEach((st, i) => {
    const y = 1.95 + i * 1.15;
    s.addShape("roundRect", { x: 0.7, y, w: 7.5, h: 0.98, rectRadius: 0.1, fill: { color: WHITE }, line: { type: "none" }, shadow: freshShadow() });
    circleIcon(s, x_ic(0.7), y + 0.19, 0.6, String(i + 1), i % 2 === 0 ? CYAN : SAFFRON, WHITE);
    s.addText(st[0], { x: 1.6, y: y + 0.08, w: 6.4, h: 0.35, fontSize: 13.5, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(st[1], { x: 1.6, y: y + 0.42, w: 6.45, h: 0.5, fontSize: 10.5, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0 });
  });
  function x_ic(x) { return x + 0.28; }
}

// ---------------- Slide 5: Tech Stack ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Technology / Tech Stack", { x: 0.7, y: 0.5, w: 9.5, h: 0.7, fontSize: 30, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);
  s.addText("Lean, API-first stack — built to demo in hours, scale to millions of merchants", {
    x: 0.7, y: 1.15, w: 10.5, h: 0.4, fontSize: 14, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  const stack = [
    ["H", "Sarvam-105B", "Indian-language LLM — understands Hindi/Hinglish merchant queries natively, no translation layer", CYAN],
    ["W", "Web App (HTML/JS)", "Lightweight, installs nowhere — works on any merchant's phone browser, low-bandwidth friendly", SAFFRON],
    ["₹", "Paytm Merchant APIs", "Sales, settlement, and QR transaction data feed the copilot's real-time business context", GREEN],
    ["N", "n8n Workflows", "Automates the suggested action itself — e.g. auto-triggers a cashback campaign, not just advice", CYAN_D],
    ["M", "Cognee Memory", "Persistent merchant memory layer — copilot remembers past patterns across sessions", SAFFRON],
    ["D", "PostgreSQL", "Stores merchant profiles, transaction history, and conversation logs", NAVY2]
  ];
  stack.forEach((it, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.7 + col * 4.05, y = 2.05 + row * 2.35;
    s.addShape("roundRect", { x, y, w: 3.8, h: 2.1, rectRadius: 0.12, fill: { color: LIGHT }, line: { type: "none" }, shadow: freshShadow() });
    circleIcon(s, x + 0.3, y + 0.3, 0.65, it[0], it[3], WHITE);
    s.addText(it[1], { x: x + 0.25, y: y + 1.05, w: 3.3, h: 0.4, fontSize: 14.5, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(it[2], { x: x + 0.25, y: y + 1.42, w: 3.35, h: 0.6, fontSize: 10, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top" });
  });
  stickyTag(s, 3.35, 1.82, 1.55, 0.5, "sponsor tech", -6, GREEN, WHITE);
}

// ---------------- Slide 5b: System Architecture ----------------
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  s.addText("System Architecture", { x: 0.7, y: 0.45, w: 9.5, h: 0.65, fontSize: 32, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.58, 1.0, false);
  s.addText("One request, five hops — from a merchant's voice to a verified action", {
    x: 0.7, y: 1.08, w: 10.5, h: 0.4, fontSize: 13, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  // ---- Main pipeline (top row) ----
  const bw = 2.2, bh = 1.15, gap = 0.25, topY = 1.85;
  const bx = [0.65, 0.65 + (bw + gap), 0.65 + 2 * (bw + gap), 0.65 + 3 * (bw + gap), 0.65 + 4 * (bw + gap)];
  archBox(s, bx[0], topY, bw, bh, "Merchant", "Voice or text,\nHindi/Hinglish", NAVY2, WHITE, "9FB6D9");
  archBox(s, bx[1], topY, bw, bh, "Vyapari Mitra App", "Web app, no\ninstall needed", NAVY2, WHITE, "9FB6D9");
  archBox(s, bx[2], topY, bw, bh, "Sarvam-105B", "Hindi/Hinglish\nintent + NLU", NAVY2, WHITE, "9FB6D9");
  archBox(s, bx[3], topY, bw, bh, "Decision Layer", "Insight engine +\nFraud Shield", "C0392B", WHITE, "F5C6C6");
  archBox(s, bx[4], topY, bw, bh, "Response", "Spoken + written\nreply to merchant", CYAN_D, WHITE, "D9F3FF");

  for (let i = 0; i < 4; i++) {
    arrow(s, bx[i] + bw, topY + bh / 2, bx[i + 1], topY + bh / 2, NAVY2);
  }

  // ---- Data & sponsor-tech layer (bottom row) ----
  const cw = 1.95, ch = 0.95, cgap = 0.25, botY = 4.55;
  const cx = [4.05, 4.05 + (cw + cgap), 4.05 + 2 * (cw + cgap), 4.05 + 3 * (cw + cgap)];
  s.addText("DATA & SPONSOR-TECH LAYER", {
    x: cx[0], y: botY - 0.38, w: cx[3] + cw - cx[0], h: 0.3, fontSize: 10, bold: true, color: MUTED, charSpacing: 1.5,
    fontFace: "Calibri", isTextBox: true, margin: 0
  });
  archBox(s, cx[0], botY, cw, ch, "Cognee Memory", null, WHITE, NAVY);
  archBox(s, cx[1], botY, cw, ch, "Paytm Settlement API", null, WHITE, NAVY);
  archBox(s, cx[2], botY, cw, ch, "Soundbox Signal", null, WHITE, NAVY);
  archBox(s, cx[3], botY, cw, ch, "n8n Workflows", null, WHITE, NAVY);

  // Cognee's memory feeds the NLU layer with merchant context (leftmost source, leftmost target — no crossing)
  arrow(s, cx[0] + cw / 2, botY, bx[2] + bw * 0.75, topY + bh, NAVY2);
  // Settlement + Soundbox feed the fraud/insight decision layer
  arrow(s, cx[1] + cw / 2, botY, bx[3] + bw * 0.3, topY + bh, NAVY2);
  arrow(s, cx[2] + cw / 2, botY, bx[3] + bw * 0.7, topY + bh, NAVY2);
  // Decision layer triggers n8n to execute the action (output, not input)
  arrow(s, bx[3] + bw * 0.6, topY + bh, cx[3] + cw / 2, botY, "C0392B", true);

  stickyTag(s, 9.15, 3.55, 1.9, 0.55, "action, not just advice", 4, SAFFRON, NAVY);
}

// ---------------- Slide 6: Why Vyapari Mitra Wins (comparative analysis) ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Why Vyapari Mitra Wins", { x: 0.7, y: 0.45, w: 9.5, h: 0.65, fontSize: 33, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.58, 1.0, false);
  s.addText("A straight, feature-by-feature comparison against every category of alternative on the market", {
    x: 0.7, y: 1.12, w: 11.4, h: 0.4, fontSize: 13, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  const CHECK = { text: "✓", options: { color: "1FAA6B", bold: true, fontSize: 16, align: "center", valign: "middle" } };
  const CROSS = { text: "✗", options: { color: "C7CDD6", bold: true, fontSize: 16, align: "center", valign: "middle" } };
  const cell = (v) => v ? CHECK : CROSS;
  const headStyle = (bg, fg) => ({ fill: { color: bg }, color: fg, bold: true, fontSize: 11.5, align: "center", valign: "middle", fontFace: "Calibri" });

  const header = [
    { text: "Capability", options: { ...headStyle(NAVY, WHITE), align: "left" } },
    { text: "Generic\nChatbots", options: headStyle(NAVY2, "C9D6EC") },
    { text: "Analytics\nDashboards", options: headStyle(NAVY2, "C9D6EC") },
    { text: "Fraud-Detection\nApps", options: headStyle(NAVY2, "C9D6EC") },
    { text: "Vyapari\nMitra", options: headStyle(CYAN_D, WHITE) }
  ];

  const features = [
    ["Understands Hindi/Hinglish natively", false, false, false, true],
    ["Grounded in THIS merchant's own data", false, true, false, true],
    ["Gives ONE concrete action, not raw numbers", false, false, false, true],
    ["Real-time fraud & QR-swap protection", false, false, true, true],
    ["Can trigger the action itself (n8n)", false, false, false, true],
    ["Zero new hardware or app install", true, true, false, true]
  ];

  const bodyRows = features.map((f, i) => {
    const rowBg = i % 2 === 0 ? WHITE : LIGHT;
    return [
      { text: f[0], options: { fill: { color: rowBg }, color: INK, fontSize: 12, align: "left", valign: "middle", fontFace: "Calibri", bold: false } },
      { ...cell(f[1]), options: { ...cell(f[1]).options, fill: { color: rowBg } } },
      { ...cell(f[2]), options: { ...cell(f[2]).options, fill: { color: rowBg } } },
      { ...cell(f[3]), options: { ...cell(f[3]).options, fill: { color: rowBg } } },
      { ...cell(f[4]), options: { ...cell(f[4]).options, fill: { color: "EAF9FF" } } }
    ];
  });

  s.addTable([header, ...bodyRows], {
    x: 0.7, y: 1.7, w: 11.95,
    colW: [4.75, 1.85, 1.85, 1.85, 1.65],
    rowH: [0.6, 0.72, 0.72, 0.72, 0.72, 0.72, 0.72],
    border: { type: "solid", color: "E3E8EE", pt: 0.75 },
    autoPage: false
  });

  stickyTag(s, 9.9, 1.15, 2.4, 0.5, "the full picture", 4, SAFFRON, NAVY);
}

// ---------------- Slide 7: Impact & Benefits ----------------
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  s.addText("Impact & Benefits", { x: 0.7, y: 0.5, w: 11.9, h: 0.7, fontSize: 35, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);

  const stats = [
    ["4M+", "Paytm merchants who could get a growth copilot instantly, at near-zero marginal cost"],
    ["10+", "Indian languages Sarvam AI can extend Vyapari Mitra to beyond this Mumbai demo"],
    ["0", "New apps to install — works inside the merchant's existing Paytm touchpoints"]
  ];
  stats.forEach((st, i) => {
    const x = 0.7 + i * 4.05;
    s.addShape("roundRect", { x, y: 1.9, w: 3.8, h: 2.0, rectRadius: 0.12, fill: { color: WHITE }, line: { type: "none" }, shadow: freshShadow() });
    s.addText(st[0], { x, y: 2.05, w: 3.8, h: 0.9, align: "center", fontSize: 46, bold: true, color: CYAN_D, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
    s.addText(st[1], { x: x + 0.25, y: 2.95, w: 3.3, h: 0.85, align: "center", fontSize: 11, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0 });
  });

  const benefits = [
    ["For Merchants", "Higher sales through personalised nudges, plus real money saved from fraud caught before it happens"],
    ["For Paytm", "Fewer fraud disputes and chargeback complaints to resolve — and a daily reason to open the app beyond payments"],
    ["For India's Economy", "Millions of small businesses get enterprise-grade growth intelligence and fraud protection in their own language"]
  ];
  benefits.forEach((b, i) => {
    const y = 4.25 + i * 0.92;
    circleIcon(s, 0.7, y, 0.55, "✓", GREEN, WHITE);
    s.addText(b[0] + ":", { x: 1.45, y: y - 0.03, w: 2.3, h: 0.6, valign: "middle", fontSize: 13, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(b[1], { x: 3.8, y: y - 0.03, w: 8.6, h: 0.6, valign: "middle", fontSize: 12, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0 });
  });
}

// ---------------- Slide 7b: Why This Wins (Venn + chart) ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Why This Wins", { x: 0.7, y: 0.5, w: 9.5, h: 0.7, fontSize: 30, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);
  s.addText("Three capabilities nobody else combines, on top of the data merchants already generate", {
    x: 0.7, y: 1.15, w: 11.3, h: 0.4, fontSize: 13, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  // ---- Venn diagram: the three capabilities that only overlap in Vyapari Mitra ----
  const vT = { x: 1.95, y: 1.85, d: 3.0 };
  const vL = { x: 0.7, y: 3.4, d: 3.0 };
  const vR = { x: 3.2, y: 3.4, d: 3.0 };
  s.addShape("ellipse", { x: vT.x, y: vT.y, w: vT.d, h: vT.d, fill: { color: CYAN, transparency: 45 }, line: { color: WHITE, width: 2 } });
  s.addShape("ellipse", { x: vL.x, y: vL.y, w: vL.d, h: vL.d, fill: { color: "C0392B", transparency: 45 }, line: { color: WHITE, width: 2 } });
  s.addShape("ellipse", { x: vR.x, y: vR.y, w: vR.d, h: vR.d, fill: { color: SAFFRON, transparency: 45 }, line: { color: WHITE, width: 2 } });

  s.addShape("roundRect", { x: 2.75, y: 4.0, w: 1.55, h: 0.55, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" }, shadow: freshShadow() });
  s.addText("Vyapari\nMitra", { x: 2.75, y: 4.0, w: 1.55, h: 0.55, align: "center", valign: "middle", fontSize: 10.5, bold: true, color: WHITE, fontFace: "Calibri", isTextBox: true, margin: 0 });

  stickyTag(s, 2.15, 1.62, 2.2, 0.48, "Vernacular AI (Sarvam)", -2, CYAN, NAVY);
  stickyTag(s, 0.35, 6.05, 1.9, 0.5, "Fraud Shield", -4, "C0392B", WHITE);
  stickyTag(s, 4.5, 6.05, 2.35, 0.5, "Autonomous Action (n8n)", 3, SAFFRON, NAVY);

  // ---- Donut chart: real payment-channel mix behind the fraud story ----
  s.addText("Payment Channel Mix", {
    x: 7.15, y: 1.75, w: 5.4, h: 0.4, fontSize: 15, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  s.addChart(pres.ChartType.doughnut, [
    { name: "Channel", labels: ["Paytm QR", "Cash"], values: [68, 32] }
  ], {
    x: 7.15, y: 2.2, w: 5.3, h: 3.15,
    chartColors: [CYAN, "E3E8EE"],
    showLegend: true, legendPos: "r", legendColor: INK, legendFontSize: 11,
    showValue: true, dataLabelColor: WHITE, dataLabelFontSize: 12, dataLabelFontBold: true,
    dataLabelFormatCode: "0\"%\"",
    showTitle: false, holeSize: 55
  });
  s.addText("68% of transactions already run through Paytm QR — exactly the ledger Fraud Shield cross-checks against.", {
    x: 7.15, y: 5.45, w: 5.35, h: 0.75, fontSize: 11, italic: true, color: MUTED, fontFace: "Calibri", isTextBox: true, margin: 0
  });
}

// ---------------- Slide 8: Business Model ----------------
{
  const s = pres.addSlide();
  s.background = { color: IVORY };
  s.addText("Business Model", { x: 0.7, y: 0.5, w: 11.9, h: 0.7, fontSize: 29, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0 });
  addLogo(s, 11.65, 0.62, 1.0, false);
  s.addText("Built to plug into Paytm's existing merchant monetisation — not a separate product to sell", {
    x: 0.7, y: 1.15, w: 11, h: 0.4, fontSize: 13.5, color: MUTED, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });

  const models = [
    ["₹", "Freemium Copilot", "Free basic insights for every merchant; premium tier for automated campaign execution & deeper analytics", CYAN],
    ["+", "Merchant Subscription Add-on", "Bundled inside existing Paytm for Business plans as a value-add tier — low CAC, existing distribution", SAFFRON],
    ["$", "Sponsored Growth Actions", "Brands pay to surface relevant offers (e.g. supplier discounts) through the copilot's suggestions", GREEN],
    ["%", "Aggregated Insights (B2B)", "Anonymised, aggregated local-commerce trend data licensed to FMCG/brand partners", CYAN_D]
  ];
  models.forEach((m, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.7 + col * 6.0, y = 2.0 + row * 2.4;
    s.addShape("roundRect", { x, y, w: 5.7, h: 2.15, rectRadius: 0.12, fill: { color: WHITE }, line: { type: "none" }, shadow: freshShadow() });
    circleIcon(s, x + 0.35, y + 0.35, 0.75, m[0], m[3], WHITE);
    s.addText(m[1], { x: x + 1.3, y: y + 0.32, w: 4.05, h: 0.5, fontSize: 15.5, bold: true, color: NAVY, fontFace: "Calibri", isTextBox: true, margin: 0 });
    s.addText(m[2], { x: x + 0.35, y: y + 1.2, w: 5.05, h: 0.85, fontSize: 11.5, color: INK, fontFace: "Calibri", isTextBox: true, margin: 0, valign: "top" });
  });

  circleIcon(s, 0.7, 6.85, 0.5, "✓", NAVY, WHITE);
  s.addText("Built compliance-first: no fund movement by the copilot itself — every payment/cashback action stays inside Paytm's existing NPCI-compliant rails; Vyapari Mitra only recommends and triggers, it never touches money directly.", {
    x: 1.35, y: 6.83, w: 10.9, h: 0.55, valign: "middle", fontSize: 11, italic: true, color: MUTED, fontFace: "Calibri", isTextBox: true, margin: 0
  });
}

// ---------------- Slide 9: Closing ----------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape("ellipse", { x: -2.8, y: -3.0, w: 6, h: 6, fill: { color: "E8F8FE" }, line: { type: "none" } });
  s.addShape("ellipse", { x: 9.8, y: 4.5, w: 5.5, h: 5.5, fill: { color: "E8F8FE" }, line: { type: "none" } });
  addLogo(s, 11.0, 0.7, 1.55, false);
  s.addText("Build with AI. Solve for India.", {
    x: 0.9, y: 2.6, w: 11.5, h: 1.0, fontSize: 43, bold: true, color: NAVY, fontFace: "Times New Roman", isTextBox: true, margin: 0
  });
  s.addText("Vyapari Mitra — giving every Paytm merchant an AI business partner who speaks their language.", {
    x: 0.9, y: 3.55, w: 10.8, h: 0.6, fontSize: 16, color: CYAN_D, italic: true, fontFace: "Calibri", isTextBox: true, margin: 0
  });
  s.addText("Team: The Vision   ·   Ayush Korde  &  Harshita Girase   ·   Powered by Sarvam AI", {
    x: 0.9, y: 6.6, w: 11.5, h: 0.4, fontSize: 13, color: MUTED, fontFace: "Calibri", isTextBox: true, margin: 0
  });
}

pres.writeFile({ fileName: path.join(__dirname, "vyapari-mitra-deck.pptx") }).then(() => {
  console.log("Deck written: vyapari-mitra-deck.pptx");
});
