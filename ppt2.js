const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Requirements for Each Construction Type - Section 403";

  // Same color palette as previous PPT
  const DARK_BG = "0D1B2A";
  const MID_BG  = "1B2A3B";
  const ACCENT  = "F4A261";
  const ACCENT2 = "2EC4B6";
  const WHITE   = "FFFFFF";
  const LIGHT   = "E8EEF4";
  const GRAY    = "94A3B8";
  const CARD_BG = "152232";
  const RED_AC  = "FB7185";
  const PURPLE  = "A78BFA";
  const GREEN   = "34D399";

  // ─── SLIDE 1: Title ───────────────────────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    // Decorative construction-themed shapes right side
    // Structural beams visual
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 0.5, w: 0.12, h: 4.8, fill: { color: "1E3A4F" }, line: { color: "1E3A4F" } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 7.9, y: 1.2, w: 0.12, h: 4.1, fill: { color: "1E3A4F" }, line: { color: "1E3A4F" } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 9.0, y: 0.8, w: 0.12, h: 4.5, fill: { color: "1E3A4F" }, line: { color: "1E3A4F" } });
    // Horizontal beams
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.1, w: 2.35, h: 0.10, fill: { color: ACCENT2, transparency: 40 }, line: { color: ACCENT2, transparency: 40 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 2.2, w: 2.35, h: 0.10, fill: { color: ACCENT2, transparency: 40 }, line: { color: ACCENT2, transparency: 40 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 3.3, w: 2.35, h: 0.10, fill: { color: ACCENT2, transparency: 40 }, line: { color: ACCENT2, transparency: 40 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 4.4, w: 2.35, h: 0.10, fill: { color: ACCENT2, transparency: 40 }, line: { color: ACCENT2, transparency: 40 } });
    // Roman numeral type badges
    const types = ["I", "II", "III", "IV", "V"];
    types.forEach((t, i) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 6.0 - i * 0.05, y: 0.5 + i * 0.6, w: 0.5, h: 0.5,
        fill: { color: ACCENT, transparency: 20 + i * 10 }, line: { color: ACCENT, transparency: 60 }
      });
      slide.addText(t, { x: 6.0 - i * 0.05, y: 0.5 + i * 0.6, w: 0.5, h: 0.5, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });

    // Section tag
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.3, w: 2.3, h: 0.35, fill: { color: ACCENT2 }, line: { color: ACCENT2 } });
    slide.addText("SECTION 401 – 403", { x: 0.5, y: 0.3, w: 2.3, h: 0.35, fontSize: 10, bold: true, color: DARK_BG, align: "center", valign: "middle", margin: 0 });

    // Main title
    slide.addText("Requirements for Each", { x: 0.5, y: 0.9, w: 6.0, h: 0.9, fontSize: 40, bold: true, color: WHITE, fontFace: "Calibri" });
    slide.addText("Construction Type", { x: 0.5, y: 1.75, w: 6.0, h: 0.9, fontSize: 40, bold: true, color: ACCENT, fontFace: "Calibri" });

    slide.addText("Exploring Types I–V, fire-resistive ratings, interior finish\nstandards, and structural requirements under the National\nBuilding Code of the Philippines.", {
      x: 0.5, y: 2.85, w: 6.0, h: 1.1, fontSize: 13, color: LIGHT, lineSpacingMultiple: 1.5
    });

    // Bottom strip
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.72, fill: { color: MID_BG }, line: { color: MID_BG } });
    slide.addText("National Building Code of the Philippines  |  Fire Code Provisions  |  IRR Section 403", {
      x: 0.5, y: 4.95, w: 9, h: 0.6, fontSize: 11, color: GRAY, align: "center", valign: "middle"
    });
  }

  // ─── SLIDE 2: Overview of 5 Construction Types ───────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.06, w: 0.06, h: 5.56, fill: { color: ACCENT2 }, line: { color: ACCENT2 } });

    slide.addText("SECTION 401", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Five Types of Building Construction", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const types = [
      { num: "I",   label: "Wood Construction",                   desc: "Wood as primary material; no specific restrictions on structural materials", color: GREEN,  fire: "Basic" },
      { num: "II",  label: "Wood + Protective Fire-Resistant",     desc: "Wood with 1-hr fire-resistive protection; fire-retardant treated wood allowed", color: ACCENT2, fire: "1-Hour" },
      { num: "III", label: "Masonry & Wood Construction",          desc: "Non-combustible exterior walls; 1-hr fire-resistive rating throughout", color: ACCENT,  fire: "1-Hour" },
      { num: "IV",  label: "Steel, Iron, Concrete, or Masonry",    desc: "Non-combustible materials; fire-retardant wood in non-bearing partitions", color: PURPLE, fire: "2–4 Hr" },
      { num: "V",   label: "Fire-Resistive Construction",          desc: "Highest standard — all walls, ceilings, partitions must be incombustible", color: RED_AC, fire: "3–4 Hr" },
    ];

    types.forEach((t, i) => {
      const y = 1.25 + i * 0.83;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.35, y, w: 9.3, h: 0.72, fill: { color: CARD_BG }
      });
      // Type badge
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.85, h: 0.72, fill: { color: t.color }, line: { color: t.color } });
      slide.addText(`Type ${t.num}`, { x: 0.35, y, w: 0.85, h: 0.72, fontSize: 10, bold: true, color: DARK_BG, align: "center", valign: "middle", margin: 0 });
      // Label + desc
      slide.addText(t.label, { x: 1.32, y: y + 0.08, w: 5.8, h: 0.32, fontSize: 13, bold: true, color: t.color });
      slide.addText(t.desc,  { x: 1.32, y: y + 0.4, w: 5.8, h: 0.28, fontSize: 10.5, color: LIGHT });
      // Fire badge right
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 8.3, y: y + 0.17, w: 1.1, h: 0.38,
        fill: { color: t.color, transparency: 75 }, line: { color: t.color, transparency: 40 }
      });
      slide.addText(t.fire, { x: 8.3, y: y + 0.17, w: 1.1, h: 0.38, fontSize: 11, bold: true, color: t.color, align: "center", valign: "middle", margin: 0 });
    });
  }

  // ─── SLIDE 3: Types I, II, III Details ────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("CONSTRUCTION TYPES", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Types I, II & III — Wood & Masonry", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const cards = [
      {
        num: "I", title: "Wood Construction", color: GREEN,
        material: "Wood as primary material; no restrictions on structural material mix.",
        implication: "Wood is combustible — fire alarms and sprinklers are crucial safety measures.",
        rating: "No mandated fire-resistive rating"
      },
      {
        num: "II", title: "Wood + Fire-Resistant Materials", color: ACCENT2,
        material: "Wood with 1-hour fire-resistive protection throughout; non-bearing partitions may use fire-retardant treated wood.",
        implication: "Fire-resistant materials add a protective layer, slowing fire spread despite wood's combustibility.",
        rating: "1-Hour fire-resistive rating"
      },
      {
        num: "III", title: "Masonry & Wood", color: ACCENT,
        material: "Non-combustible, fire-resistive exterior walls; wood allowed in limited interior areas. Must maintain 1-hr fire-resistive rating.",
        implication: "Masonry provides strength and fire resistance; wood use is limited, balancing cost with safety.",
        rating: "1-Hour fire-resistive rating"
      },
    ];

    cards.forEach((c, i) => {
      const x = 0.3 + i * 3.23;
      // Card background
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.28, w: 3.07, h: 3.9,
        fill: { color: CARD_BG },
        shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.3 }
      });
      // Top color bar
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.28, w: 3.07, h: 0.07, fill: { color: c.color }, line: { color: c.color } });
      // Type badge
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.15, y: 1.42, w: 0.65, h: 0.55,
        fill: { color: c.color, transparency: 75 }, line: { color: c.color }
      });
      slide.addText(`Type\n${c.num}`, { x: x + 0.15, y: 1.42, w: 0.65, h: 0.55, fontSize: 10, bold: true, color: c.color, align: "center", valign: "middle", margin: 0 });
      // Title
      slide.addText(c.title, { x: x + 0.9, y: 1.42, w: 2.0, h: 0.55, fontSize: 12, bold: true, color: c.color, lineSpacingMultiple: 1.2 });

      // Sections
      slide.addText("MATERIALS", { x: x + 0.15, y: 2.15, w: 2.7, h: 0.28, fontSize: 9, bold: true, color: c.color, charSpacing: 2 });
      slide.addText(c.material, { x: x + 0.15, y: 2.42, w: 2.8, h: 0.9, fontSize: 10.5, color: LIGHT, lineSpacingMultiple: 1.3 });

      slide.addText("IMPLICATION", { x: x + 0.15, y: 3.42, w: 2.7, h: 0.28, fontSize: 9, bold: true, color: GRAY, charSpacing: 2 });
      slide.addText(c.implication, { x: x + 0.15, y: 3.7, w: 2.8, h: 0.75, fontSize: 10.5, color: LIGHT, lineSpacingMultiple: 1.3 });

      // Rating pill at bottom
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.15, y: 4.65, w: 2.7, h: 0.35,
        fill: { color: c.color, transparency: 80 }, line: { color: c.color, transparency: 50 }
      });
      slide.addText(c.rating, { x: x + 0.15, y: 4.65, w: 2.7, h: 0.35, fontSize: 10, bold: true, color: c.color, align: "center", valign: "middle", margin: 0 });
    });
  }

  // ─── SLIDE 4: Types IV & V Details ───────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("CONSTRUCTION TYPES", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Types IV & V — Steel, Concrete & Fire-Resistive", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const cards = [
      {
        num: "IV", title: "Steel, Iron, Concrete, or Masonry", color: PURPLE,
        material: "Non-combustible materials (steel, iron, concrete, masonry). Fire-retardant treated wood allowed only in non-bearing partitions.",
        implication: "Much higher fire protection than wood-based types. Incombustible materials prevent fire spread and maintain structural integrity.",
        rating: "Exterior: 4-Hr  |  Interior: 1-Hr  |  Frame: 2-Hr"
      },
      {
        num: "V", title: "Fire-Resistive Construction", color: RED_AC,
        material: "All walls, ceilings, and partitions must be non-combustible and fire-resistant — steel, iron, concrete, or masonry only.",
        implication: "Highest fire safety standard. More expensive due to quality materials required, but provides maximum protection against fire hazards.",
        rating: "Exterior: 4-Hr  |  Bearing Walls: 3-Hr  |  Frame: 3-Hr"
      },
    ];

    cards.forEach((c, i) => {
      const x = 0.35 + i * 4.95;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.28, w: 4.7, h: 3.95,
        fill: { color: CARD_BG },
        shadow: { type: "outer", color: "000000", blur: 12, offset: 4, angle: 135, opacity: 0.35 }
      });
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.28, w: 4.7, h: 0.07, fill: { color: c.color }, line: { color: c.color } });

      // Type badge large
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.2, y: 1.45, w: 0.9, h: 0.75,
        fill: { color: c.color, transparency: 70 }, line: { color: c.color }
      });
      slide.addText(`Type\n${c.num}`, { x: x + 0.2, y: 1.45, w: 0.9, h: 0.75, fontSize: 12, bold: true, color: c.color, align: "center", valign: "middle", margin: 0 });

      slide.addText(c.title, { x: x + 1.25, y: 1.5, w: 3.2, h: 0.7, fontSize: 14, bold: true, color: c.color, lineSpacingMultiple: 1.2 });

      slide.addText("MATERIALS", { x: x + 0.2, y: 2.4, w: 4.2, h: 0.28, fontSize: 9, bold: true, color: c.color, charSpacing: 2 });
      slide.addText(c.material, { x: x + 0.2, y: 2.68, w: 4.25, h: 0.95, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.35 });

      slide.addText("IMPLICATION", { x: x + 0.2, y: 3.72, w: 4.2, h: 0.28, fontSize: 9, bold: true, color: GRAY, charSpacing: 2 });
      slide.addText(c.implication, { x: x + 0.2, y: 4.0, w: 4.25, h: 0.75, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.3 });

      // Rating bar bottom
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 4.9, w: 4.7, h: 0.33,
        fill: { color: c.color, transparency: 75 }, line: { color: c.color, transparency: 40 }
      });
      slide.addText(c.rating, { x: x + 0.1, y: 4.9, w: 4.5, h: 0.33, fontSize: 10, bold: true, color: c.color, align: "center", valign: "middle", margin: 0 });
    });
  }

  // ─── SLIDE 5: Fire Resistive Requirements ────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION 403 — IRR", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Fire-Resistive Rating Requirements", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // Table header
    const colColors = [GREEN, ACCENT2, ACCENT, PURPLE, RED_AC];
    const typeLabels = ["Type I", "Type II", "Type III", "Type IV", "Type V"];
    const headerH = 0.45;
    const rowH = 0.72;
    const colW = [2.5, 1.35, 1.35, 1.35, 1.35, 1.35];
    const startX = 0.3;
    const startY = 1.22;

    // Component label column header
    slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: startY, w: colW[0], h: headerH, fill: { color: "1A3550" }, line: { color: "243C52" } });
    slide.addText("COMPONENT", { x: startX + 0.1, y: startY, w: colW[0] - 0.1, h: headerH, fontSize: 11, bold: true, color: GRAY, valign: "middle" });

    // Type header columns
    typeLabels.forEach((label, i) => {
      const x = startX + colW[0] + i * colW[1];
      slide.addShape(pres.shapes.RECTANGLE, { x, y: startY, w: colW[1], h: headerH, fill: { color: colColors[i] }, line: { color: colColors[i] } });
      slide.addText(label, { x, y: startY, w: colW[1], h: headerH, fontSize: 11, bold: true, color: DARK_BG, align: "center", valign: "middle", margin: 0 });
    });

    const rows = [
      { component: "Exterior Walls",                  values: ["—", "1-Hr", "1-Hr", "4-Hr", "4-Hr"] },
      { component: "Interior Walls & Partitions",     values: ["—", "1-Hr", "1-Hr", "1-Hr", "3-Hr*"] },
      { component: "Floors & Roofs",                  values: ["—", "1-Hr", "1-Hr", "1-Hr", "1-Hr"] },
      { component: "Structural Frames",               values: ["—", "1-Hr", "1-Hr", "2-Hr", "3-Hr"] },
      { component: "Exterior Doors & Windows",        values: ["1-Hr", "1-Hr", "1-Hr", "1-Hr", "1-Hr"] },
    ];

    rows.forEach((row, ri) => {
      const y = startY + headerH + ri * rowH;
      const rowFill = ri % 2 === 0 ? CARD_BG : "101E2D";

      // Component cell
      slide.addShape(pres.shapes.RECTANGLE, { x: startX, y, w: colW[0], h: rowH, fill: { color: rowFill }, line: { color: "243C52" } });
      slide.addText(row.component, { x: startX + 0.1, y, w: colW[0] - 0.1, h: rowH, fontSize: 12, color: LIGHT, valign: "middle" });

      row.values.forEach((val, vi) => {
        const x = startX + colW[0] + vi * colW[1];
        slide.addShape(pres.shapes.RECTANGLE, { x, y, w: colW[1], h: rowH, fill: { color: rowFill }, line: { color: "243C52" } });
        const isHigh = val.includes("3") || val.includes("4");
        const isEmpty = val === "—";
        slide.addText(val, {
          x, y, w: colW[1], h: rowH,
          fontSize: 13, bold: !isEmpty,
          color: isEmpty ? GRAY : (isHigh ? colColors[vi] : LIGHT),
          align: "center", valign: "middle", margin: 0
        });
      });
    });

    // Footnote
    slide.addText("* 3-Hr rating applies to bearing walls in Type V. Vertical openings, floors, and roofs require 1-Hr.", {
      x: 0.35, y: 5.2, w: 9.3, h: 0.32, fontSize: 10, color: GRAY, italic: true
    });
  }

  // ─── SLIDE 6: Interior Wall & Ceiling Finishes ───────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION 403 — IRR", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Interior Wall & Ceiling Finishes", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    const items = [
      {
        num: "A", title: "Flame-Spread Characteristics",
        color: ACCENT,
        points: [
          "Must not exceed flame-spread density of untreated wood (Tunnel Test)",
          "Combustion products must not be more toxic than untreated wood",
          "Ensures minimal contribution to fire growth and toxic gas production"
        ]
      },
      {
        num: "B", title: "Exemptions from Flame-Spread Requirements",
        color: ACCENT2,
        points: [
          "Door and window frames and trims are exempted",
          "Materials less than 1.00 mm thick cemented to walls or ceilings are exempted",
          "Prevents unnecessary restrictions on non-hazardous materials"
        ]
      },
      {
        num: "C", title: "Flame-Retardant Treatment",
        color: PURPLE,
        points: [
          "Finishes requiring flame-spread proofing must have a rating of 50 or less",
          "Treated materials perform better in fire — slowing spread and minimizing risk",
          "Applies to any finish not naturally meeting the flame-spread standard"
        ]
      },
    ];

    items.forEach((item, i) => {
      const x = 0.3 + i * 3.23;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.28, w: 3.07, h: 3.95,
        fill: { color: CARD_BG },
        shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.3 }
      });
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.28, w: 3.07, h: 0.07, fill: { color: item.color }, line: { color: item.color } });

      // Badge
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.15, y: 1.45, w: 0.55, h: 0.55,
        fill: { color: item.color }, line: { color: item.color }
      });
      slide.addText(item.num, { x: x + 0.15, y: 1.45, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: DARK_BG, align: "center", valign: "middle", margin: 0 });

      slide.addText(item.title, { x: x + 0.85, y: 1.5, w: 2.05, h: 0.5, fontSize: 12, bold: true, color: item.color, lineSpacingMultiple: 1.2 });

      item.points.forEach((pt, pi) => {
        const py = 2.2 + pi * 0.9;
        slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: py + 0.12, w: 0.08, h: 0.08, fill: { color: item.color }, line: { color: item.color } });
        slide.addText(pt, { x: x + 0.35, y: py, w: 2.6, h: 0.75, fontSize: 11, color: LIGHT, lineSpacingMultiple: 1.3 });
      });
    });
  }

  // ─── SLIDE 7: Structural Framework Standards ─────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION 403 — IRR", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Structural Framework Standards", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    slide.addText("Structural elements must follow the relevant provisions of the Fire Code of the Philippines to ensure compatibility with national fire safety laws.", {
      x: 0.4, y: 1.22, w: 9.1, h: 0.65, fontSize: 13, color: LIGHT, lineSpacingMultiple: 1.4
    });

    const elements = [
      {
        icon: "🏗️", title: "Structural Framework",
        body: "Must maintain structural integrity under high-temperature conditions. Frame fire ratings increase with construction type — from 1-Hr (Type II/III) to 3-Hr (Type V).",
        color: ACCENT
      },
      {
        icon: "🚪", title: "Exits & Stairs",
        body: "Must support safe and rapid evacuation during fire emergencies. Design must comply with minimum width, clear passage, and fire-resistance requirements.",
        color: ACCENT2
      },
      {
        icon: "🪟", title: "Exterior Openings",
        body: "Doors and windows must have a minimum 1-Hr fire-resistive rating across all construction types to prevent fire spread through openings — common weak points.",
        color: PURPLE
      },
      {
        icon: "🏠", title: "Roofs",
        body: "Regulated to minimize fire entry from the exterior and prevent spread between structures. Must comply with material and fire-resistance standards per construction type.",
        color: RED_AC
      },
    ];

    elements.forEach((el, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.35 + col * 4.85;
      const y = 2.1 + row * 1.65;

      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 4.6, h: 1.45,
        fill: { color: CARD_BG },
        shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.25 }
      });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.45, fill: { color: el.color }, line: { color: el.color } });

      slide.addText(el.title, { x: x + 0.2, y: y + 0.1, w: 4.2, h: 0.38, fontSize: 14, bold: true, color: el.color });
      slide.addText(el.body,  { x: x + 0.2, y: y + 0.5, w: 4.25, h: 0.88, fontSize: 11.5, color: LIGHT, lineSpacingMultiple: 1.3 });
    });
  }

  // ─── SLIDE 8: Section 402 — Changes in Construction Types ────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SECTION 402", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Changes in Construction Types", { x: 0.4, y: 0.5, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE });

    // Main rule box
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y: 1.25, w: 9.3, h: 0.9,
      fill: { color: "1A3550" }
    });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.25, w: 0.06, h: 0.9, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText("No building may be altered to a different construction type unless it fully complies with the standards and requirements of the new type or sub-type. Any change must not compromise the fire safety and structural integrity of the building.", {
      x: 0.55, y: 1.3, w: 9.0, h: 0.8, fontSize: 13, color: LIGHT, italic: true, lineSpacingMultiple: 1.4
    });

    const rules = [
      {
        title: "Compliance Required",
        body: "Alteration to a different type demands full compliance with new type's standards — materials, fire ratings, and structural requirements must all be met.",
        color: ACCENT
      },
      {
        title: "Building Official Approval",
        body: "The Building Official may approve a change if the proposed construction is less hazardous than the current one, allowing flexibility in renovation projects.",
        color: ACCENT2
      },
      {
        title: "Less Hazardous = Approvable",
        body: "Changes that result in a lower fire and safety risk may receive approval — enabling practical renovation without compromising public safety.",
        color: PURPLE
      },
      {
        title: "Structural Integrity Preserved",
        body: "Any modification must ensure the building's ability to maintain structural integrity under fire and load conditions is not diminished.",
        color: GREEN
      },
    ];

    rules.forEach((r, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.35 + col * 4.85;
      const y = 2.35 + row * 1.55;

      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 4.6, h: 1.35,
        fill: { color: CARD_BG },
        shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.25 }
      });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 0.05, fill: { color: r.color }, line: { color: r.color } });
      slide.addText(r.title, { x: x + 0.2, y: y + 0.1, w: 4.2, h: 0.38, fontSize: 13, bold: true, color: r.color });
      slide.addText(r.body,  { x: x + 0.2, y: y + 0.5, w: 4.2, h: 0.78, fontSize: 11.5, color: LIGHT, lineSpacingMultiple: 1.3 });
    });
  }

  // ─── SLIDE 9: Summary / Key Takeaways ────────────────────────────────────
  {
    const slide = pres.addSlide();
    slide.background = { color: DARK_BG };

    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: ACCENT }, line: { color: ACCENT } });

    slide.addText("SUMMARY", { x: 0.4, y: 0.2, w: 9, h: 0.35, fontSize: 11, color: ACCENT2, bold: true, charSpacing: 3 });
    slide.addText("Key Takeaways", { x: 0.4, y: 0.5, w: 8, h: 0.6, fontSize: 28, bold: true, color: WHITE });

    const points = [
      { num: "01", title: "Five Construction Types (I–V)",     body: "Each type differs in primary materials — from basic wood (Type I) to fully non-combustible fire-resistive construction (Type V).", color: ACCENT },
      { num: "02", title: "Fire Ratings Increase with Type",   body: "From 1-Hr basic ratings in Types II/III up to 4-Hr exterior walls and 3-Hr structural frames in Type V.", color: ACCENT2 },
      { num: "03", title: "Finishes Must Meet Flame Standards", body: "Interior finishes must not exceed untreated wood's flame-spread. Proofed materials must have a rating of 50 or less.", color: PURPLE },
      { num: "04", title: "Changes Require Full Compliance",    body: "Altering a building's construction type requires meeting all standards of the new type — no shortcuts allowed.", color: RED_AC },
    ];

    points.forEach((p, i) => {
      const y = 1.3 + i * 1.0;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.35, y, w: 9.3, h: 0.88, fill: { color: CARD_BG }
      });
      slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.7, h: 0.88, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(p.num, { x: 0.35, y, w: 0.7, h: 0.88, fontSize: 15, bold: true, color: DARK_BG, align: "center", valign: "middle", margin: 0 });
      slide.addText(p.title, { x: 1.2, y: y + 0.08, w: 4.5, h: 0.35, fontSize: 14, bold: true, color: p.color });
      slide.addText(p.body,  { x: 1.2, y: y + 0.47, w: 8.2, h: 0.35, fontSize: 11.5, color: LIGHT });
    });

    // Bottom bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.37, fill: { color: MID_BG }, line: { color: MID_BG } });
    slide.addText("National Building Code of the Philippines — Sections 401, 402 & 403", {
      x: 0.5, y: 5.27, w: 9, h: 0.33, fontSize: 10, color: GRAY, align: "center"
    });
  }

  await pres.writeFile({ fileName: "Construction_Types_Section403.pptx" });
  console.log("Done!");
}

main().catch(console.error);
