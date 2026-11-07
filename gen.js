function buildReport2(D, r) {
  var NAVY = "1A1A1A", ACCENT = "FBAA11", GRAY = "5F5F5F", BAND = "FFF3D6", INK = "1F1F1F";
  var FONT = "Calibri";
  var CENTER = D.AlignmentType.CENTER;
  var NONE = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };

  function fit(w, h, maxW, maxH) {
    var k = Math.min(maxW / w, maxH / h, 1);
    return { width: Math.max(1, Math.round(w * k)), height: Math.max(1, Math.round(h * k)) };
  }

  function fill(text) {
    var v = r.vars || {};
    return String(text || "")
      .replace(/\{\{SEGURO\}\}/g, v.seguro || "")
      .replace(/\{\{POLIZA\}\}/g, v.poliza || "")
      .replace(/\{\{SINIESTRO\}\}/g, v.siniestro || "")
      .replace(/\{\{VALOR\}\}/g, v.valor || "")
      .replace(/\{\{APARTAMENTO\}\}/g, (r.cover && r.cover.apartamento) || "")
      .replace(/\{\{CONJUNTO\}\}/g, (r.cover && r.cover.conjunto) || "");
  }

  function textPara(b, keepNext) {
    var t = fill(b.x).replace(/\s+/g, " ").trim();
    if (!t) return null;
    if (b.t === "h") {
      return new D.Paragraph({
        keepNext: true,
        spacing: { before: 240, after: 120 },
        children: [new D.TextRun({ text: t, bold: true, size: 24, color: NAVY, font: FONT })]
      });
    }
    if (b.t === "ol") {
      return new D.Paragraph({
        numbering: { reference: "numeros", level: 0 },
        alignment: D.AlignmentType.JUSTIFIED,
        keepNext: !!keepNext,
        spacing: { after: 160, line: 276, lineRule: D.LineRuleType.AUTO },
        children: [new D.TextRun({ text: t, font: FONT, size: 22, color: INK })]
      });
    }
    if (b.t === "li") {
      return new D.Paragraph({
        numbering: { reference: "viñetas", level: 0 },
        keepNext: !!keepNext,
        spacing: { after: 100, line: 276, lineRule: D.LineRuleType.AUTO },
        children: [new D.TextRun({ text: t, font: FONT, size: 22, color: INK })]
      });
    }
    return new D.Paragraph({
      alignment: D.AlignmentType.JUSTIFIED,
      keepNext: !!keepNext,
      spacing: { after: 200, line: 276, lineRule: D.LineRuleType.AUTO },
      children: [new D.TextRun({ text: t, font: FONT, size: 22, color: INK })]
    });
  }

  function sectionTitle(text, first) {
    return new D.Paragraph({
      heading: D.HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      keepNext: true,
      spacing: { before: 0, after: 280 },
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 18, color: ACCENT, space: 4 } },
      children: [new D.TextRun({ text: text, bold: true, size: 32, color: NAVY, font: FONT })]
    });
  }

  function imagePara(im, maxW, maxH, keepNext) {
    var d = fit(im.w, im.h, maxW, maxH);
    return new D.Paragraph({
      alignment: CENTER,
      keepNext: !!keepNext,
      spacing: { before: 60, after: 200 },
      children: [new D.ImageRun({ data: im.data, type: im.type || "jpg", transformation: d })]
    });
  }

  function photoTable(photos) {
    var n = photos.length;
    var maxH = n <= 2 ? 320 : n <= 4 ? 240 : 190;
    var maxW = n === 1 ? 330 : 280;
    var rows = [];
    for (var i = 0; i < n; i += 2) {
      var pair = n === 1 ? [photos[i]] : [photos[i], photos[i + 1]];
      rows.push(new D.TableRow({
        cantSplit: true,
        children: pair.map(function (p) {
          var kids = [];
          if (p) {
            var d = fit(p.w, p.h, maxW, maxH);
            kids.push(new D.Paragraph({ alignment: CENTER, keepNext: !!p.caption, spacing: { after: 60 }, children: [new D.ImageRun({ data: p.data, type: p.type || "jpg", transformation: d })] }));
            if (p.caption) kids.push(new D.Paragraph({ alignment: CENTER, spacing: { after: 120 }, children: [new D.TextRun({ text: p.caption, italics: true, size: 18, color: GRAY, font: FONT })] }));
            else kids.push(new D.Paragraph({ spacing: { after: 80 }, children: [] }));
          } else {
            kids.push(new D.Paragraph(""));
          }
          return new D.TableCell({
            width: { size: n === 1 ? 9360 : 4680, type: D.WidthType.DXA },
            borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
            margins: { top: 60, bottom: 60, left: 80, right: 80 },
            children: kids
          });
        })
      }));
    }
    var cols = n === 1 ? [9360] : [4680, 4680];
    return new D.Table({
      width: { size: 9360, type: D.WidthType.DXA },
      columnWidths: cols,
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
      rows: rows
    });
  }

  var c = r.cover || {};
  var children = [];

  children.push(new D.Paragraph({
    spacing: { before: 2400, after: 320 },
    border: { bottom: { style: D.BorderStyle.SINGLE, size: 36, color: ACCENT, space: 8 } },
    children: [new D.TextRun({ text: "INFORME", bold: true, size: 96, color: NAVY, font: FONT })]
  }));
  children.push(new D.Table({
    width: { size: 9360, type: D.WidthType.DXA },
    columnWidths: [9360],
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    rows: [new D.TableRow({ children: [new D.TableCell({
      width: { size: 9360, type: D.WidthType.DXA },
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
      shading: { type: D.ShadingType.CLEAR, color: "auto", fill: ACCENT },
      margins: { top: 130, bottom: 130, left: 240, right: 240 },
      children: [new D.Paragraph({ children: [new D.TextRun({ text: (c.subtitulo || "INSPECCIÓN Y VALORACIÓN DE DAÑOS").toUpperCase(), bold: true, size: 32, color: NAVY, font: FONT })] })]
    })] })]
  }));
  children.push(new D.Paragraph({
    spacing: { before: 700, after: 80 },
    children: [new D.TextRun({ text: c.apartamento || "", bold: true, size: 40, color: NAVY, font: FONT })]
  }));
  children.push(new D.Paragraph({
    children: [new D.TextRun({ text: c.conjunto || "", bold: true, size: 28, color: NAVY, font: FONT })]
  }));
  if (c.fecha) children.push(new D.Paragraph({
    spacing: { before: 400 },
    children: [new D.TextRun({ text: c.fecha, size: 24, color: GRAY, font: FONT })]
  }));
  var elab = String(c.elaboradores || "").split(/\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  if (elab.length) children.push(new D.Paragraph({
    spacing: { before: 1000, after: 80 },
    border: { left: { style: D.BorderStyle.SINGLE, size: 24, color: ACCENT, space: 8 } },
    children: [new D.TextRun({ text: "Elaborado por:", bold: true, size: 20, color: NAVY, font: FONT })]
  }));
  elab.forEach(function (l) {
    children.push(new D.Paragraph({
      spacing: { after: 40 },
      border: { left: { style: D.BorderStyle.SINGLE, size: 24, color: ACCENT, space: 8 } },
      children: [new D.TextRun({ text: l, size: 22, color: INK, font: FONT })]
    }));
  });

  children.push(sectionTitle("INTRODUCCIÓN"));
  (r.intro || []).forEach(function (b, i, a) {
    var p = textPara(b, (b.t === "li" || b.t === "ol") && a[i + 1] && a[i + 1].t === b.t);
    if (p) children.push(p);
  });

  children.push(sectionTitle((r.planosTitulo || "PLANO CON EL REGISTRO DE LAS AFECTACIONES").toUpperCase()));
  var pb = r.planos || [];
  pb.forEach(function (b, i) {
    if (b.t === "img") {
      if (!b.img) return;
      var legend = b.role === "fijo";
      children.push(imagePara(b.img, legend ? 420 : 600, legend ? 260 : 420, false));
    } else {
      var next = pb[i + 1];
      var p = textPara(b, next && next.t === "img");
      if (p) children.push(p);
    }
  });

  var spaces = r.espacios || [];
  if (spaces.length) {
    children.push(sectionTitle("INFORME FOTOGRÁFICO"));
    spaces.forEach(function (sp) {
      if (!sp.photos.length) return;
      children.push(new D.Paragraph({
        heading: D.HeadingLevel.HEADING_2,
        keepNext: true,
        spacing: { before: 280, after: 120 },
        shading: { type: D.ShadingType.CLEAR, color: "auto", fill: BAND },
        border: { left: { style: D.BorderStyle.SINGLE, size: 36, color: ACCENT, space: 6 } },
        children: [new D.TextRun({ text: sp.name.toUpperCase(), bold: true, size: 24, color: NAVY, font: FONT })]
      }));
      children.push(photoTable(sp.photos));
    });
  }

  if (r.presupuesto && (r.presupuesto.img || r.presupuesto.nota)) {
    children.push(sectionTitle("PRESUPUESTO"));
    if (r.presupuesto.img) children.push(imagePara(r.presupuesto.img, 620, 700, !!r.presupuesto.nota));
    if (r.presupuesto.nota && r.presupuesto.img) {
      children.push(new D.Paragraph({
        spacing: { after: 120 },
        children: [new D.TextRun({ text: r.presupuesto.nota, italics: true, size: 20, color: GRAY, font: FONT })]
      }));
    } else if (r.presupuesto.nota) {
      var np = textPara({ t: "p", x: r.presupuesto.nota }, false);
      if (np) children.push(np);
    }
  }

  children.push(sectionTitle((r.conclTitulo || "CONCLUSIONES").toUpperCase()));
  var cb = r.concl || [];
  cb.forEach(function (b, i) {
    var last = i === cb.length - 1;
    var p = textPara(b, last && !!r.firma);
    if (p) children.push(p);
  });
  if (r.firma) children.push(imagePara(r.firma, 520, 220, false));

  var footer = new D.Footer({
    children: [new D.Paragraph({
      tabStops: [{ type: D.TabStopType.RIGHT, position: 9360 }],
      border: { top: { style: D.BorderStyle.SINGLE, size: 12, color: ACCENT, space: 6 } },
      children: [
        new D.TextRun({ text: "\tPágina ", size: 18, color: GRAY, font: FONT }),
        new D.TextRun({ children: [D.PageNumber.CURRENT], size: 18, color: GRAY, font: FONT }),
        new D.TextRun({ text: " de ", size: 18, color: GRAY, font: FONT }),
        new D.TextRun({ children: [D.PageNumber.TOTAL_PAGES], size: 18, color: GRAY, font: FONT })
      ]
    })]
  });

  return new D.Document({
    background: { color: "FFFFFF" },
    creator: "Generador de informes",
    title: "Informe " + (c.apartamento || ""),
    styles: {
      default: { document: { run: { font: FONT, size: 22, color: INK } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 32, bold: true, color: NAVY },
          paragraph: { spacing: { before: 0, after: 280 }, keepNext: true } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 24, bold: true, color: NAVY },
          paragraph: { spacing: { before: 280, after: 120 }, keepNext: true } }
      ]
    },
    numbering: {
      config: [{
        reference: "numeros",
        levels: [{ level: 0, format: D.LevelFormat.DECIMAL, text: "%1.", alignment: D.AlignmentType.START,
          style: { paragraph: { indent: { left: 540, hanging: 360 } } } }]
      }, {
        reference: "viñetas",
        levels: [{ level: 0, format: D.LevelFormat.BULLET, text: "•", alignment: D.AlignmentType.START,
          style: { paragraph: { indent: { left: 540, hanging: 270 } } } }]
      }]
    },
    sections: [{
      properties: {
        titlePage: true,
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1300, left: 1440, header: 500, footer: 500 } }
      },
      footers: { default: footer, first: new D.Footer({ children: [new D.Paragraph("")] }) },
      children: children
    }]
  });
}
if (typeof module !== "undefined") module.exports = buildReport2;
