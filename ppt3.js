const pptxgen = require("pptxgenjs");

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Pasuquin Wind Farm - Windmill Project Proposal";

  const DARK_BG = "0A1628";
  const MID_BG  = "112240";
  const CARD_BG = "0E1E35";
  const ACCENT  = "00B4D8";
  const ACCENT2 = "90E0EF";
  const GOLD    = "F4A261";
  const WHITE   = "FFFFFF";
  const LIGHT   = "CAF0F8";
  const GRAY    = "8EACC4";
  const GREEN   = "34D399";
  const RED_AC  = "FB7185";
  const PURPLE  = "A78BFA";



  // ─── SLIDE 1: TITLE ───────────────────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    // Top accent bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });
    // Bottom accent bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.55, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    // Windmill blade deco (right side)
    for (let i = 0; i < 6; i++) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 7.2 + i * 0.38, y: 0.1, w: 0.18, h: 5.5,
        fill: { color: ACCENT, transparency: 88 - i * 8 }, line: { color: ACCENT, transparency: 95 }
      });
    }

    // University badge
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.3, w: 4.2, h: 0.38, fill: { color: MID_BG }, line: { color: ACCENT, transparency: 50 } });
    slide.addText("PANGASINAN STATE UNIVERSITY  |  BSME-4A", {
      x: 0.5, y: 0.3, w: 4.2, h: 0.38, fontSize: 9, color: ACCENT2, bold: true, align: "center", valign: "middle", charSpacing: 1
    });

    // Main title
    slide.addText("WINDMILL PROJECT", { x: 0.5, y: 0.95, w: 6.5, h: 0.85, fontSize: 44, bold: true, color: WHITE, fontFace: "Calibri" });
    slide.addText("PROPOSAL", { x: 0.5, y: 1.75, w: 6.5, h: 0.85, fontSize: 44, bold: true, color: ACCENT, fontFace: "Calibri" });

    // Subtitle
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.72, w: 0.06, h: 0.9, fill: { color: GOLD }, line: { color: GOLD } });
    slide.addText("A Feasibility and Siting Study\nPasuquin, Ilocos Norte, Philippines", {
      x: 0.72, y: 2.72, w: 5.8, h: 0.9, fontSize: 15, color: LIGHT, lineSpacingMultiple: 1.5
    });

    // Info row
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 6.3, h: 0.75, fill: { color: MID_BG }, line: { color: ACCENT, transparency: 60 } });
    slide.addText("Submitted by: Dan Mark C. Pastoral   |   April 29, 2026   |   Engr. Marfel D. Rosario, PME", {
      x: 0.6, y: 3.9, w: 6.1, h: 0.65, fontSize: 11, color: GRAY, align: "center", valign: "middle"
    });

    // Tag
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.82, w: 2.8, h: 0.38, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText("50–100 MW  |  20 Turbines  |  Luzon Grid", {
      x: 0.5, y: 4.82, w: 2.8, h: 0.38, fontSize: 10, bold: true, color: DARK_BG, align: "center", valign: "middle"
    });
  }



  // ─── SLIDE 2: EXECUTIVE SUMMARY ───────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.07, w: 0.07, h: 5.55, fill: { color: GOLD }, line: { color: GOLD } });

    slide.addText("EXECUTIVE SUMMARY", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Project Overview", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const params = [
      { label: "Project Location",      value: "Pasuquin, Ilocos Norte, Philippines", color: ACCENT },
      { label: "Target Capacity",        value: "~50–100 MW (expandable)",             color: GREEN },
      { label: "Turbine Type",           value: "Horizontal Axis Wind Turbine (HAWT)", color: ACCENT2 },
      { label: "Estimated Turbines",     value: "15–30 units (3–5 MW each)",           color: GOLD },
      { label: "Grid Connection",        value: "Luzon Grid via NGCP",                 color: PURPLE },
      { label: "Project Lifespan",       value: "25–30 years",                         color: RED_AC },
      { label: "Key Benefit",            value: "Clean, renewable baseload power for Ilocos Norte", color: GREEN },
    ];

    params.forEach((p, i) => {
      const y = 1.22 + i * 0.6;
      const even = i % 2 === 0;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.52, fill: { color: even ? CARD_BG : MID_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.06, h: 0.52, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(p.label, { x: 0.5, y: y + 0.06, w: 2.8, h: 0.4, fontSize: 11, bold: true, color: p.color, valign: "middle" });
      slide.addText(p.value, { x: 3.4, y: y + 0.06, w: 6.1, h: 0.4, fontSize: 11, color: LIGHT, valign: "middle" });
    });
  }



  // ─── SLIDE 3: SITE SELECTION ──────────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION II — SITE SELECTION", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Why Pasuquin, Ilocos Norte?", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const reasons = [
      { num: "01", title: "High Wind Power Density",          body: "7–9 m/s average (Class 4–5). Located along Northwest Luzon Wind Corridor exposed to Northeast Monsoon (Amihan). Capacity factor: 35–45%.", color: ACCENT },
      { num: "02", title: "Strategic Coastal & Hilly Terrain", body: "Coastal plains for sea wind capture + ridge areas for high-velocity winds at altitude. Maximizes energy output per turbine.", color: GREEN },
      { num: "03", title: "Proximity to Infrastructure",       body: "Near Bangui & Burgos wind farms. NGCP transmission grid and heavy-haul routes already established in Ilocos Norte.", color: GOLD },
      { num: "04", title: "Eco-Tourism Potential",            body: "Dual purpose: power plant + tourist landmark (similar to Bangui Wind Farm). Boosts local economy and small businesses.", color: PURPLE },
      { num: "05", title: "Sustainability & Energy Security", body: "Reduces fossil fuel reliance. Wind is free and inexhaustible — hedge against volatile coal/oil prices. Supports PH climate goals.", color: RED_AC },
    ];

    reasons.forEach((r, i) => {
      const y = 1.22 + i * 0.82;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.72, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.65, h: 0.72, fill: { color: r.color }, line: { color: r.color } });
      slide.addText(r.num, { x: 0.3, y, w: 0.65, h: 0.72, fontSize: 13, bold: true, color: DARK_BG, align: "center", valign: "middle" });
      slide.addText(r.title, { x: 1.1, y: y + 0.06, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: r.color });
      slide.addText(r.body,  { x: 1.1, y: y + 0.36, w: 8.4, h: 0.3, fontSize: 10.5, color: LIGHT });
    });
  }



  // ─── SLIDE 4: PROCESS FLOW DIAGRAM ────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION II-B — PROCESS FLOW", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Wind-to-Grid Energy Flow", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const stages = [
      { label: "WIND\nRESOURCE",   sub: "Amihan\n7–9 m/s",          color: ACCENT },
      { label: "ROTOR\nBLADES",    sub: "3 Blades\n63.5 m each",    color: GREEN },
      { label: "GEARBOX",          sub: "1:90 ratio\n1,500 RPM",    color: GOLD },
      { label: "DFIG\nGENERATOR",  sub: "690 V AC\n60 Hz",          color: PURPLE },
      { label: "TRANSFORMER\n1",   sub: "690V →\n33 kV",            color: RED_AC },
      { label: "SUBSTATION\n33kV", sub: "33kV →\n115 kV",           color: ACCENT2 },
      { label: "LUZON\nGRID",      sub: "115 kV\n~5.58 MW net",     color: GREEN },
    ];

    const boxW = 1.18;
    const boxH = 1.25;
    const startX = 0.25;
    const y = 1.6;

    stages.forEach((s, i) => {
      const x = startX + i * 1.38;
      // Box
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: boxW, h: boxH, fill: { color: CARD_BG }, line: { color: s.color, transparency: 30 } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: boxW, h: 0.07, fill: { color: s.color }, line: { color: s.color } });
      slide.addText(s.label, { x, y: y + 0.1, w: boxW, h: 0.55, fontSize: 10, bold: true, color: s.color, align: "center", valign: "middle" });
      slide.addText(s.sub,   { x, y: y + 0.65, w: boxW, h: 0.55, fontSize: 9, color: LIGHT, align: "center", valign: "middle" });
      // Arrow
      if (i < stages.length - 1) {
        slide.addShape(pres.shapes.RECTANGLE, { x: x + boxW, y: y + 0.55, w: 0.18, h: 0.1, fill: { color: GRAY }, line: { color: GRAY } });
        slide.addText("▶", { x: x + boxW + 0.05, y: y + 0.48, w: 0.15, h: 0.22, fontSize: 9, color: GRAY, align: "center" });
      }
    });

    // Energy balance note
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.05, w: 9.4, h: 0.55, fill: { color: MID_BG }, line: { color: ACCENT, transparency: 60 } });
    slide.addText("Energy Balance: Wind 13.96 MW  →  Aero Loss 7.68 MW  →  Gearbox 6.03 MW  →  Generator 5.85 MW  →  Net Grid Delivery: 5.58 MW/turbine", {
      x: 0.4, y: 3.07, w: 9.2, h: 0.5, fontSize: 10, color: ACCENT2, align: "center", valign: "middle"
    });

    // SCADA box
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.75, w: 4.5, h: 0.9, fill: { color: CARD_BG }, line: { color: GOLD, transparency: 40 } });
    slide.addText("SCADA System", { x: 0.4, y: 3.8, w: 4.3, h: 0.3, fontSize: 11, bold: true, color: GOLD });
    slide.addText("Real-time monitoring: power output, vibration, temperature, fault alarms — all 20 turbines via OPC-UA protocol", {
      x: 0.4, y: 4.1, w: 4.3, h: 0.5, fontSize: 10, color: LIGHT, lineSpacingMultiple: 1.3
    });

    // Support systems
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.75, w: 4.6, h: 0.9, fill: { color: CARD_BG }, line: { color: PURPLE, transparency: 40 } });
    slide.addText("Support Systems", { x: 5.2, y: 3.8, w: 4.4, h: 0.3, fontSize: 11, bold: true, color: PURPLE });
    slide.addText("Pitch/yaw control  •  Emergency braking  •  Power converter  •  Auxiliary diesel genset (500 kVA)", {
      x: 5.2, y: 4.1, w: 4.4, h: 0.5, fontSize: 10, color: LIGHT, lineSpacingMultiple: 1.3
    });
  }



  // ─── SLIDE 5: TECHNICAL SPECIFICATIONS ────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION III — TECHNICAL SPECS", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Turbine Specifications", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const specs = [
      { param: "Turbine Type",         value: "3-Bladed Horizontal Axis Wind Turbine (HAWT)",  color: ACCENT },
      { param: "Rated Capacity",       value: "3.0 – 5.0 MW per unit",                         color: GREEN },
      { param: "Rotor Diameter",       value: "120 – 145 meters",                               color: GOLD },
      { param: "Hub Height",           value: "90 – 110 meters",                                color: PURPLE },
      { param: "Cut-in Wind Speed",    value: "3.0 m/s",                                        color: ACCENT2 },
      { param: "Rated Wind Speed",     value: "12 – 14 m/s",                                    color: GREEN },
      { param: "Cut-out Wind Speed",   value: "25 m/s",                                         color: RED_AC },
      { param: "Blade Material",       value: "Fiberglass-reinforced epoxy composite (GFRP)",   color: ACCENT },
      { param: "Generator Type",       value: "Doubly Fed Induction Generator (DFIG)",          color: GOLD },
      { param: "Design Lifespan",      value: "25 years (with scheduled maintenance)",          color: ACCENT2 },
    ];

    const half = Math.ceil(specs.length / 2);
    specs.forEach((s, i) => {
      const col = Math.floor(i / half);
      const row = i % half;
      const x = 0.3 + col * 4.85;
      const y = 1.22 + row * 0.46;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 0.4, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 0.4, fill: { color: s.color }, line: { color: s.color } });
      slide.addText(s.param, { x: x + 0.12, y: y + 0.04, w: 2.0, h: 0.32, fontSize: 10, bold: true, color: s.color, valign: "middle" });
      slide.addText(s.value, { x: x + 2.15, y: y + 0.04, w: 2.35, h: 0.32, fontSize: 10, color: LIGHT, valign: "middle" });
    });
  }



  // ─── SLIDE 6: DESIGN CONSTRAINTS ──────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("DESIGN CONSTRAINTS", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Engineering Basis & Limitations", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const constraints = [
      { num: "C1", title: "Betz Limit",            body: "Max 59.3% wind energy extractable. At V=8.5 m/s, 130m rotor → 27.7 MW available → 5.0 MW practical rated output after losses.", color: ACCENT },
      { num: "C2", title: "Wake Effect Spacing",   body: "7D inline (910 m) × 4D crosswind (520 m) spacing to minimize turbulence losses between turbines.", color: GREEN },
      { num: "C3", title: "Wind Shear (Hub Ht.)",  body: "Power Law: V(100m)=8.19 m/s (+26% vs ground). Since P∝V³, this gives ~2× more power justifying 100m towers.", color: GOLD },
      { num: "C4", title: "Safety Setbacks",       body: "≥500m from residences (noise <45dB). ≥200m from coastline (storm surge buffer). Non-negotiable regulatory requirement.", color: RED_AC },
      { num: "C5", title: "Grid Voltage (33 kV)",  body: "Collection at 33 kV — industry standard up to 200 MW. Step-up 3.48:1 transformer to 115 kV for NGCP Ilocos Norte grid.", color: PURPLE },
      { num: "C6", title: "Foundation Loads",      body: "Overturning moment = 115,000 kN·m. Octagonal RC mat 21m wide. Safety factor = 2.30 ≥ 1.5 required. 144 anchor bolts at 65% utilization.", color: ACCENT2 },
    ];

    constraints.forEach((c, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.3 + col * 4.85;
      const y = 1.22 + row * 1.42;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 1.28, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 0.06, fill: { color: c.color }, line: { color: c.color } });
      slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.12, y: y + 0.12, w: 0.52, h: 0.52, fill: { color: c.color, transparency: 75 }, line: { color: c.color } });
      slide.addText(c.num, { x: x + 0.12, y: y + 0.12, w: 0.52, h: 0.52, fontSize: 10, bold: true, color: c.color, align: "center", valign: "middle" });
      slide.addText(c.title, { x: x + 0.75, y: y + 0.14, w: 3.7, h: 0.32, fontSize: 12, bold: true, color: c.color });
      slide.addText(c.body,  { x: x + 0.12, y: y + 0.7, w: 4.35, h: 0.52, fontSize: 10, color: LIGHT, lineSpacingMultiple: 1.3 });
    });
  }



  // ─── SLIDE 7: ENGINEERING CALCULATIONS ────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION III-A — CALCULATIONS", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Key Engineering Calculations", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // Betz Calc card
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.55, h: 3.45, fill: { color: CARD_BG }, line: { color: ACCENT, transparency: 50 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.55, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText("Wind Power Density (Betz)", { x: 0.4, y: 1.27, w: 4.3, h: 0.38, fontSize: 13, bold: true, color: ACCENT });
    slide.addText("P = ½ × ρ × A × V³", { x: 0.4, y: 1.7, w: 4.3, h: 0.4, fontSize: 13, bold: true, color: GOLD, align: "center" });

    const betzRows = [
      ["Rotor radius (r)", "65 m"],
      ["Swept area (A)", "13,273 m²"],
      ["Air density (ρ)", "1.225 kg/m³"],
      ["Wind speed (V)", "8.5 m/s"],
      ["Available power", "~27.7 MW"],
      ["After Betz (×0.593)", "~16.4 MW"],
      ["After drivetrain losses", "~11.5 MW"],
      ["Practical rated output", "5.0 MW ✓"],
    ];
    betzRows.forEach(([param, val], i) => {
      const y = 2.18 + i * 0.36;
      const bg = i % 2 === 0 ? "0E1E35" : "122540";
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 4.45, h: 0.33, fill: { color: bg }, line: { color: "1A3550" } });
      slide.addText(param, { x: 0.45, y, w: 2.5, h: 0.33, fontSize: 10, color: LIGHT, valign: "middle" });
      slide.addText(val,   { x: 2.9, y, w: 1.8, h: 0.33, fontSize: 10, bold: true, color: i === 7 ? GREEN : ACCENT2, align: "right", valign: "middle" });
    });

    // Capacity Factor card
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.55, h: 3.45, fill: { color: CARD_BG }, line: { color: GREEN, transparency: 50 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.55, h: 0.06, fill: { color: GREEN }, line: { color: GREEN } });
    slide.addText("Capacity Factor & Annual Energy", { x: 5.25, y: 1.27, w: 4.3, h: 0.38, fontSize: 13, bold: true, color: GREEN });
    slide.addText("CF = AEP ÷ (MW × 8,760 hrs)", { x: 5.25, y: 1.7, w: 4.3, h: 0.4, fontSize: 13, bold: true, color: GOLD, align: "center" });

    const cfRows = [
      ["Rated capacity",        "5.0 MW"],
      ["Hours/year",            "8,760 hrs"],
      ["Max theoretical AEP",   "43,800 MWh/yr"],
      ["Capacity factor",       "38%"],
      ["Gross AEP/turbine",     "16,644 MWh/yr"],
      ["Electrical losses",     "7% (1,165 MWh)"],
      ["Net AEP/turbine",       "15,479 MWh/yr"],
      ["Total AEP (20 units)",  "309,580 MWh/yr ✓"],
    ];
    cfRows.forEach(([param, val], i) => {
      const y = 2.18 + i * 0.36;
      const bg = i % 2 === 0 ? "0E1E35" : "122540";
      slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y, w: 4.45, h: 0.33, fill: { color: bg }, line: { color: "1A3550" } });
      slide.addText(param, { x: 5.3, y, w: 2.5, h: 0.33, fontSize: 10, color: LIGHT, valign: "middle" });
      slide.addText(val,   { x: 7.75, y, w: 1.8, h: 0.33, fontSize: 10, bold: true, color: i === 7 ? GREEN : ACCENT2, align: "right", valign: "middle" });
    });
  }



  // ─── SLIDE 8: PLANT SITING & LAYOUT ───────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION IV — PLANT SITING", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Site Layout & Configuration", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // Layout config
    const config = [
      { label: "Turbines",         value: "20 WTGs (5 rows × 4 units)", color: ACCENT },
      { label: "Inline spacing",   value: "7D = 910 m (wake recovery)",  color: GREEN },
      { label: "Cross-wind",       value: "4D = 520 m",                  color: GOLD },
      { label: "Gross area",       value: "~9.46 km²",                   color: PURPLE },
      { label: "Setback: homes",   value: "≥500 m (noise <45dB)",        color: RED_AC },
      { label: "Setback: coast",   value: "≥200 m (storm surge buffer)", color: ACCENT2 },
    ];
    config.forEach((c, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.3 + col * 4.85;
      const y = 1.22 + row * 0.55;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 0.46, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 0.46, fill: { color: c.color }, line: { color: c.color } });
      slide.addText(c.label, { x: x + 0.15, y: y + 0.07, w: 2.1, h: 0.32, fontSize: 11, bold: true, color: c.color, valign: "middle" });
      slide.addText(c.value, { x: x + 2.3, y: y + 0.07, w: 2.2, h: 0.32, fontSize: 11, color: LIGHT, valign: "middle" });
    });

    // Ancillary facilities table
    slide.addText("ANCILLARY FACILITIES", { x: 0.3, y: 3.0, w: 9, h: 0.35, fontSize: 11, bold: true, color: GOLD, charSpacing: 2 });
    const facilities = [
      ["Maintenance Workshop", "Storage: spare parts, gearbox oil, blade tools"],
      ["Control Room / SCADA", "24/7 supervisory wind-farm control center"],
      ["On-site Substation",   "33/115 kV step-up transformer + switchgear"],
      ["Emergency Fire Station","Nacelle fire response, high-angle rescue"],
      ["Eco-Tourism Center",   "Public viewing deck, exhibits — like Bangui Wind Farm"],
    ];
    facilities.forEach(([name, desc], i) => {
      const y = 3.42 + i * 0.39;
      const bg = i % 2 === 0 ? CARD_BG : MID_BG;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.35, fill: { color: bg }, line: { color: "1A3550" } });
      slide.addText(name, { x: 0.4, y, w: 2.8, h: 0.35, fontSize: 10, bold: true, color: ACCENT2, valign: "middle" });
      slide.addText(desc, { x: 3.3, y, w: 6.3, h: 0.35, fontSize: 10, color: LIGHT, valign: "middle" });
    });
  }



  // ─── SLIDE 9: ENVIRONMENTAL IMPACT ────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION V — ENVIRONMENTAL", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Environmental Impact Assessment", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // Positive impacts
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 0.06, h: 1.95, fill: { color: GREEN }, line: { color: GREEN } });
    slide.addText("POSITIVE IMPACTS", { x: 0.45, y: 1.22, w: 9, h: 0.3, fontSize: 11, bold: true, color: GREEN, charSpacing: 2 });
    const positive = [
      "GHG Reduction — each 5 MW turbine avoids ~8,000–10,000 tonnes CO₂/year vs coal power",
      "Zero direct air pollutant emissions during operation",
      "Small physical footprint — land between turbines usable for agriculture",
      "No water discharge; no thermal pollution to Luzon Strait",
    ];
    positive.forEach((p, i) => {
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.42, y: 1.57 + i * 0.38, w: 0.1, h: 0.1, fill: { color: GREEN }, line: { color: GREEN } });
      slide.addText(p, { x: 0.62, y: 1.52 + i * 0.38, w: 9.1, h: 0.35, fontSize: 11, color: LIGHT });
    });

    // Negative impacts + mitigation
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.28, w: 0.06, h: 2.1, fill: { color: RED_AC }, line: { color: RED_AC } });
    slide.addText("POTENTIAL IMPACTS & MITIGATION", { x: 0.45, y: 3.28, w: 9, h: 0.3, fontSize: 11, bold: true, color: RED_AC, charSpacing: 2 });
    const mitigation = [
      ["Bird/bat mortality",    "Radar-based curtailment during migration season"],
      ["Noise pollution",       "500m setback; low-noise models; quarterly monitoring"],
      ["Visual/landscape",      "Community consultation; Bangui aesthetic branding"],
      ["Shadow flicker",        "Model flicker zones; adjust operating hours if needed"],
    ];
    mitigation.forEach(([impact, measure], i) => {
      const y = 3.65 + i * 0.42;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.42, y, w: 9.3, h: 0.36, fill: { color: i % 2 === 0 ? CARD_BG : MID_BG }, line: { color: "1A3550" } });
      slide.addText(impact,  { x: 0.52, y, w: 2.4, h: 0.36, fontSize: 10, bold: true, color: RED_AC, valign: "middle" });
      slide.addText(measure, { x: 3.0, y, w: 6.6, h: 0.36, fontSize: 10, color: LIGHT, valign: "middle" });
    });
  }



  // ─── SLIDE 10: FINANCIAL ANALYSIS ─────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION VI — FINANCIAL", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Financial & Economic Analysis", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // CAPEX card
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.55, h: 2.45, fill: { color: CARD_BG }, line: { color: GOLD, transparency: 40 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.55, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });
    slide.addText("CAPITAL EXPENDITURE", { x: 0.4, y: 1.27, w: 4.3, h: 0.35, fontSize: 12, bold: true, color: GOLD });
    const capex = [
      ["Wind Turbines (15–30 units)", "PHP 3.0–6.0 Billion"],
      ["Civil Works & Foundations",   "PHP 450–900 Million"],
      ["Electrical & Substation",     "PHP 300–600 Million"],
      ["Grid Connection (NGCP)",      "PHP 200–500 Million"],
      ["Development & EIS",           "PHP 50–100 Million"],
      ["TOTAL CAPEX",                 "PHP 4.4–8.9 Billion"],
    ];
    capex.forEach(([item, cost], i) => {
      const y = 1.68 + i * 0.32;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 4.45, h: 0.28, fill: { color: i === 5 ? "1A3550" : "0E1E35" }, line: { color: "1A3550" } });
      slide.addText(item, { x: 0.45, y, w: 2.5, h: 0.28, fontSize: i === 5 ? 10.5 : 10, bold: i === 5, color: i === 5 ? GOLD : LIGHT, valign: "middle" });
      slide.addText(cost, { x: 2.8, y, w: 1.9, h: 0.28, fontSize: i === 5 ? 10.5 : 10, bold: i === 5, color: i === 5 ? GOLD : ACCENT2, align: "right", valign: "middle" });
    });

    // Revenue card
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.55, h: 2.45, fill: { color: CARD_BG }, line: { color: GREEN, transparency: 40 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.55, h: 0.06, fill: { color: GREEN }, line: { color: GREEN } });
    slide.addText("REVENUE PROJECTIONS", { x: 5.25, y: 1.27, w: 4.3, h: 0.35, fontSize: 12, bold: true, color: GREEN });
    const revenue = [
      ["Installed Capacity",     "50–100 MW"],
      ["Capacity Factor",        "38% average"],
      ["Annual Energy Output",   "166–332 GWh/year"],
      ["FiT Rate (ERC)",         "PHP 8.53/kWh"],
      ["Gross Annual Revenue",   "PHP 1.42–2.83 Billion"],
      ["Simple Payback Period",  "8–12 years"],
    ];
    revenue.forEach(([item, val], i) => {
      const y = 1.68 + i * 0.32;
      slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y, w: 4.45, h: 0.28, fill: { color: i === 4 ? "1A3550" : "0E1E35" }, line: { color: "1A3550" } });
      slide.addText(item, { x: 5.3, y, w: 2.4, h: 0.28, fontSize: i === 4 ? 10.5 : 10, bold: i === 4, color: i === 4 ? GREEN : LIGHT, valign: "middle" });
      slide.addText(val,  { x: 7.55, y, w: 2.0, h: 0.28, fontSize: i === 4 ? 10.5 : 10, bold: i === 4, color: i === 4 ? GREEN : ACCENT2, align: "right", valign: "middle" });
    });

    // Economic benefits
    slide.addText("ECONOMIC & SOCIAL BENEFITS", { x: 0.3, y: 3.82, w: 9, h: 0.3, fontSize: 11, bold: true, color: ACCENT2, charSpacing: 2 });
    const benefits = [
      "50–100 permanent jobs during operations  •  300–500 construction jobs",
      "Real property tax & business tax to Pasuquin LGU  •  Eco-tourism revenue",
      "Reduces fossil fuel dependence  •  Stabilizes long-term energy prices",
    ];
    benefits.forEach((b, i) => {
      const y = 4.17 + i * 0.38;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.32, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.06, h: 0.32, fill: { color: ACCENT2 }, line: { color: ACCENT2 } });
      slide.addText(b, { x: 0.45, y, w: 9.15, h: 0.32, fontSize: 10.5, color: LIGHT, valign: "middle" });
    });
  }



  // ─── SLIDE 11: CONCLUSION ─────────────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.55, w: 10, h: 0.07, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("CONCLUSION & RECOMMENDATIONS", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT, bold: true, charSpacing: 3 });
    slide.addText("Key Takeaways", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const points = [
      { num: "01", title: "Technically Feasible",      body: "Pasuquin's Class 4–5 wind resource, 38% capacity factor, and existing NGCP infrastructure make the project viable using standard HAWT technology.", color: ACCENT },
      { num: "02", title: "Economically Viable",       body: "PHP 4.4–8.9B CAPEX, PHP 8.53/kWh FiT rate, and 8–12 year simple payback period over a 25-year project life deliver strong long-term returns.", color: GREEN },
      { num: "03", title: "Environmentally Responsible", body: "Avoids ~160,000–200,000 tonnes CO₂/year. Zero water discharge. Small footprint. Full EIS and ECC compliance with DENR and DOE.", color: GOLD },
      { num: "04", title: "Serves ~77,400 Households", body: "Total net AEP of 309,580 MWh/year from 20 turbines. Supports PH 35% RE target by 2030 under RA 9513.", color: PURPLE },
    ];

    points.forEach((p, i) => {
      const y = 1.28 + i * 0.95;
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.83, fill: { color: CARD_BG }, line: { color: "1A3550" } });
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.7, h: 0.83, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(p.num, { x: 0.3, y, w: 0.7, h: 0.83, fontSize: 14, bold: true, color: DARK_BG, align: "center", valign: "middle" });
      slide.addText(p.title, { x: 1.12, y: y + 0.07, w: 4.5, h: 0.33, fontSize: 13, bold: true, color: p.color });
      slide.addText(p.body,  { x: 1.12, y: y + 0.43, w: 8.45, h: 0.33, fontSize: 11, color: LIGHT });
    });

    // Final line
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.12, w: 9.4, h: 0.35, fill: { color: MID_BG }, line: { color: MID_BG } });
    slide.addText("Dan Mark C. Pastoral  |  BSME-4A  |  Pangasinan State University  |  April 2026", {
      x: 0.3, y: 5.14, w: 9.4, h: 0.3, fontSize: 10, color: GRAY, align: "center"
    });
  }



  await pres.writeFile({ fileName: "Pasuquin_WindFarm_Proposal.pptx" });
  console.log("Done! Pasuquin_WindFarm_Proposal.pptx generated.");
}

main().catch(console.error);
