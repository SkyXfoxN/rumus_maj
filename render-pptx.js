/* render-pptx.js — Bina fail .pptx daripada takrifan slaid yang SAMA
   dengan yang digunakan penjana HTML. */

(function () {
  window.renderPPTX = function (spec, filename) {
    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_WIDE";
    pres.company = "MyAduan Johor";
    let page = 0;

    const TB = { isTextBox: true, margin: 0 };

    function card(s, b, fill) {
      s.addShape(pres.ShapeType.roundRect, {
        x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: 0.06,
        fill: { color: fill }, line: { color: fill, width: 0 },
      });
    }

    function chart(s, b) {
      const common = {
        x: b.x, y: b.y, w: b.w, h: b.h,
        showLegend: false, catGridLine: { style: "none" },
        valAxisLineShow: false, catAxisLineShow: false,
        catAxisLabelFontFace: F.BODY, valAxisLabelFontFace: F.BODY,
      };
      if (b.k === "col" || b.k === "bar") {
        const horiz = b.k === "bar";
        /* PowerPoint membungkus label paksi sendiri, jadi label penuh
           dihantar tanpa dipendekkan. */
        const labels = horiz ? b.labels.slice().reverse() : b.labels;
        const values = horiz ? b.values.slice().reverse() : b.values;
        s.addChart("bar", [{ name: "Jumlah", labels, values }], Object.assign({}, common, {
          barDir: horiz ? "bar" : "col", barGapWidthPct: horiz ? 45 : 55,
          chartColors: [b.color],
          showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: FONT.nilai,
          dataLabelFontFace: F.BODY, dataLabelColor: C.TEXT,
          catAxisLabelColor: horiz ? C.TEXT : C.MUTED, catAxisLabelFontSize: FONT.paksi,
          valAxisHidden: true, valGridLine: { style: "none" },
          valAxisMinVal: 0, valAxisMaxVal: Math.max.apply(null, b.values) * (horiz ? 1.2 : 1.22),
        }));
      } else if (b.k === "donut") {
        /* Sama seperti penjana HTML - nombor dalam petunjuk supaya hirisan
           kecil tidak menyebabkan label bertindih. */
        /* PowerPoint tidak membenarkan kedudukan label ditetapkan per hirisan,
           dan penukar PDF mengabaikan "outEnd" - label hirisan nipis akan
           bertindan. Nilai diletak dalam petunjuk supaya sentiasa terbaca.
           Paparan pelayar pula melukis nombor di luar gelang dengan garis
           penunjuk, kerana di sana kedudukan boleh dikawal sepenuhnya. */
        const lab = b.labels.map((l, i) => l + "  (" + fmt(b.values[i]) + ")");
        s.addChart("doughnut", [{ name: "Pecahan", labels: lab, values: b.values }],
          Object.assign({}, common, {
            holeSize: 55, chartColors: b.colors || SERIES,
            showLegend: true, legendPos: "r", legendFontSize: FONT.petunjuk,
            legendFontFace: F.BODY, legendColor: C.TEXT,
            showValue: false,
          }));
      } else if (b.k === "line") {
        s.addChart("line", b.series.map((x) => ({ name: x.name, labels: b.labels, values: x.values })),
          Object.assign({}, common, {
            chartColors: b.series.map((x) => x.color),
            lineDataSymbolSize: 6, lineSize: 2.5,
            showLegend: true, legendPos: "t", legendFontSize: FONT.petunjuk,
            legendFontFace: F.BODY, legendColor: C.TEXT,
            showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "#,##0", dataLabelFontSize: FONT.nilai,
            dataLabelFontFace: F.BODY, dataLabelColor: C.TEXT,
            catAxisLabelColor: C.MUTED, catAxisLabelFontSize: FONT.paksi,
            valAxisLabelColor: C.MUTED, valAxisLabelFontSize: FONT.paksi,
            valGridLine: { color: C.LINE, size: 0.75 },
          }));
      }
    }

    function block(s, b) {
      if (b.t === "stat") {
        card(s, b, b.fill || C.MIST);
        s.addText(b.label.toUpperCase(), Object.assign({
          x: b.x + 0.26, y: b.y + 0.18, w: b.w - 0.52, h: 0.26,
          fontFace: F.BODY, fontSize: 9.5, color: b.labelColor || C.MUTED,
          bold: true, charSpacing: 1.2,
        }, TB));
        s.addText(fmt(b.value), Object.assign({
          x: b.x + 0.24, y: b.y + 0.42, w: b.w - 0.5, h: 0.62,
          fontFace: F.HEAD, fontSize: b.vsize || 40, color: b.color || C.INK, bold: true,
        }, TB));
        if (b.sub) {
          s.addText(b.sub, Object.assign({
            x: b.x + 0.26, y: b.y + b.h - 0.42, w: b.w - 0.52, h: 0.28,
            fontFace: F.BODY, fontSize: 10, color: b.subColor || C.MUTED,
          }, TB));
        }

      } else if (b.t === "card") {
        card(s, b, C.MIST);
        if (b.title) {
          s.addText(b.title, Object.assign({
            x: b.x + 0.28, y: b.y + 0.16, w: b.w - 0.56, h: 0.3,
            fontFace: F.BODY, fontSize: 12, color: C.INK, bold: true,
          }, TB));
        }
        if (b.note) {
          s.addText(b.note, Object.assign({
            x: b.x + 0.28, y: b.y + 0.44, w: b.w - 0.56, h: 0.26,
            fontFace: F.BODY, fontSize: 9.5, color: C.MUTED,
          }, TB));
        }
        if (b.body) {
          if (['col','bar','donut','line'].indexOf(b.body.k) >= 0) chart(s, b.body);
          else block(s, Object.assign({ t: b.body.k }, b.body));
        }

      } else if (b.t === "segment") {
        const barH = b.barH || 0.46;
        const tot = b.segs.reduce((a, x) => a + x.value, 0) || 1;
        let cx = b.x;
        b.segs.forEach((sg, i) => {
          const sw = Math.max(0.22, (sg.value / tot) * b.w);
          const last = i === b.segs.length - 1;
          const ww = last ? b.x + b.w - cx : sw;
          s.addShape(pres.ShapeType.rect, { x: cx, y: b.y, w: ww, h: barH, fill: { color: sg.color } });
          if (sg.value / tot > 0.07) {
            s.addText(fmt(sg.value), Object.assign({
              x: cx, y: b.y, w: ww, h: barH, align: "center", valign: "middle",
              fontFace: F.BODY, fontSize: 11, color: C.WHITE, bold: true,
            }, TB));
          }
          cx += sw;
        });
        const legY = b.y + barH + 0.22, cellW = b.w / b.segs.length;
        b.segs.forEach((sg, i) => {
          const lx = b.x + i * cellW;
          s.addShape(pres.ShapeType.rect, { x: lx, y: legY + 0.05, w: 0.15, h: 0.15, fill: { color: sg.color } });
          s.addText(sg.label, Object.assign({
            x: lx + 0.24, y: legY - 0.02, w: cellW - 0.3, h: 0.3,
            fontFace: F.BODY, fontSize: 9.5, color: C.TEXT,
          }, TB));
        });

      } else if (b.t === "insight") {
        s.addShape(pres.ShapeType.roundRect, {
          x: b.x, y: b.y, w: b.w, h: 0.62, rectRadius: 0.06,
          fill: { color: C.INSIGHT_BG }, line: { color: C.INSIGHT_BG, width: 0 },
        });
        s.addShape(pres.ShapeType.ellipse, { x: b.x + 0.22, y: b.y + 0.2, w: 0.22, h: 0.22, fill: { color: C.TEAL } });
        s.addText(b.text, Object.assign({
          x: b.x + 0.56, y: b.y + 0.09, w: b.w - 0.8, h: 0.44, valign: "middle",
          fontFace: F.BODY, fontSize: 10.5, color: C.INSIGHT_TX,
        }, TB));

      } else if (b.t === "funnel") {
        const gap = 0.5, bw = (b.w - gap * (b.steps.length - 1)) / b.steps.length;
        b.steps.forEach((st, i) => {
          const bx = b.x + i * (bw + gap);
          s.addShape(pres.ShapeType.roundRect, {
            x: bx, y: b.y, w: bw, h: b.h, rectRadius: 0.07,
            fill: { color: st.fill }, line: { color: st.fill, width: 0 },
          });
          s.addText(st.label.toUpperCase(), Object.assign({
            x: bx + 0.2, y: b.y + 0.2, w: bw - 0.4, h: 0.5,
            fontFace: F.BODY, fontSize: 9.5, color: st.labelColor, bold: true,
            charSpacing: 1, lineSpacingMultiple: 1.1,
          }, TB));
          s.addText(fmt(st.value), Object.assign({
            x: bx + 0.18, y: b.y + 0.72, w: bw - 0.36, h: 0.68,
            fontFace: F.HEAD, fontSize: 38, color: st.valueColor, bold: true,
          }, TB));
          if (st.note) {
            s.addText(st.note, Object.assign({
              x: bx + 0.2, y: b.y + b.h - 0.42, w: bw - 0.4, h: 0.3,
              fontFace: F.BODY, fontSize: 9.5, color: st.noteColor,
            }, TB));
          }
          if (i < b.steps.length - 1) {
            s.addShape(pres.ShapeType.rightArrow, {
              x: bx + bw + 0.09, y: b.y + b.h / 2 - 0.13, w: 0.32, h: 0.26,
              fill: { color: C.LINE },
            });
          }
        });

      } else if (b.t === "table") {
        const isLeft = (i) => (b.leftCols ? b.leftCols.indexOf(i) >= 0 : i === 0);
        const hdr = b.head.map((t, i) => ({
          text: t,
          options: {
            bold: true, color: C.WHITE, fill: { color: C.INK }, fontSize: 10.5,
            align: isLeft(i) ? "left" : "center", valign: "middle",
            fontFace: F.BODY, margin: [0.06, 0.1, 0.06, 0.1],
          },
        }));
        const body = b.rows.map((r, ri) => r.map((cell, i) => {
          const o = typeof cell === "object" && cell !== null ? cell : { text: cell };
          const teks = /^\d{4,}$/.test(String(o.text)) ? fmt(o.text) : String(o.text);
          return {
            text: teks,
            options: {
              color: o.color || C.TEXT, bold: o.bold || i === 0,
              fill: { color: ri % 2 === 0 ? C.WHITE : C.MIST },
              fontSize: b.fs || 10.5, fontFace: F.BODY,
              align: isLeft(i) ? "left" : "center", valign: "middle",
              margin: [0.06, 0.1, 0.06, 0.1],
            },
          };
        }));
        s.addTable([hdr].concat(body), {
          x: b.x, y: b.y, w: b.w, colW: b.colW,
          border: { type: "solid", color: C.LINE, pt: 0.5 },
          rowH: b.rowH || 0.32, autoPage: false,
        });

      } else if (b.t === "rank") {
        const max = Math.max.apply(null, b.items.map((i) => i[1])) || 1;
        b.items.forEach((it, i) => {
          const ry = b.y + i * 0.52;
          s.addText(it[0], Object.assign({
            x: b.x, y: ry, w: b.w - 1.9, h: 0.28,
            fontFace: F.BODY, fontSize: 10.5, color: C.TEXT,
          }, TB));
          s.addShape(pres.ShapeType.rect, {
            x: b.x + b.w - 1.85, y: ry + 0.06,
            w: Math.max(0.08, (it[1] / max) * 1.35), h: 0.17,
            fill: { color: b.color || C.OCEAN },
          });
          s.addText(fmt(it[1]), Object.assign({
            x: b.x + b.w - 0.42, y: ry, w: 0.42, h: 0.28, align: "right",
            fontFace: F.BODY, fontSize: 10.5, color: C.INK, bold: true,
          }, TB));
        });

      } else if (b.t === "finding") {
        if (!b.text) return;
        s.addText(b.text, Object.assign({
          x: b.x, y: b.y, w: b.w, h: b.h || 0.44, valign: "top",
          fontFace: F.BODY, fontSize: 9.5, color: C.MUTED, italic: true,
          lineSpacingMultiple: 1.15,
        }, TB));

      } else if (b.t === "para") {
        if (!b.text) return;
        s.addText(b.text, Object.assign({
          x: b.x, y: b.y, w: b.w, h: b.h, valign: "top",
          fontFace: F.BODY, fontSize: 11, color: C.TEXT, lineSpacingMultiple: 1.35,
        }, TB));

      } else if (b.t === "label") {
        s.addText(b.text, Object.assign({
          x: b.x, y: b.y, w: b.w, h: 0.28,
          fontFace: F.BODY, fontSize: 11, color: C.INK, bold: true,
        }, TB));
      }
    }

    spec.forEach((sl) => {
      const s = pres.addSlide();

      if (sl.kind === "title") {
        const i = sl.info;
        s.background = { color: C.INK };
        s.addShape(pres.ShapeType.rect, { x: 0, y: 5.9, w: G.W, h: 1.6, fill: { color: "0C1E2E" } });
        s.addText(i.sistem.toUpperCase(), Object.assign({ x: G.M, y: 1.75, w: 10, h: 0.32, fontFace: F.BODY, fontSize: 13, color: "7FB2C4", charSpacing: 3, bold: true }, TB));
        s.addText(i.tajuk, Object.assign({ x: G.M, y: 2.25, w: 10.5, h: 1.6, fontFace: F.HEAD, fontSize: 46, color: C.WHITE, bold: true, lineSpacingMultiple: 1.05 }, TB));
        s.addText(i.periode, Object.assign({ x: G.M, y: 4.15, w: 6, h: 0.6, fontFace: F.HEAD, fontSize: 26, color: C.ICE, italic: true }, TB));
        s.addText([{ text: "Disediakan oleh  ", options: { color: "6E8FA3" } }, { text: i.disediakan, options: { color: C.WHITE, bold: true } }],
          Object.assign({ x: G.M, y: 6.25, w: 6, h: 0.3, fontFace: F.BODY, fontSize: 12 }, TB));
        s.addText([{ text: "Sumber data  ", options: { color: "6E8FA3" } }, { text: i.sumber, options: { color: "C7D8E2" } }],
          Object.assign({ x: G.M, y: 6.62, w: 9, h: 0.3, fontFace: F.BODY, fontSize: 11 }, TB));

      } else if (sl.kind === "section") {
        s.background = { color: C.INK };
        s.addShape(pres.ShapeType.ellipse, { x: G.M, y: 2.05, w: 0.95, h: 0.95, fill: { color: C.TEAL } });
        s.addText(String(sl.num), Object.assign({ x: G.M, y: 2.05, w: 0.95, h: 0.95, align: "center", valign: "middle", fontFace: F.HEAD, fontSize: 34, color: C.WHITE, bold: true }, TB));
        s.addText(sl.title, Object.assign({ x: G.M, y: 3.3, w: 7.2, h: 1.5, fontFace: F.HEAD, fontSize: 42, color: C.WHITE, bold: true, lineSpacingMultiple: 1.05 }, TB));
        sl.items.forEach((t, i) => {
          const y = 2.15 + i * 0.62;
          s.addShape(pres.ShapeType.ellipse, { x: 8.15, y: y + 0.13, w: 0.16, h: 0.16, fill: { color: "5E9EB5" } });
          s.addText(t, Object.assign({ x: 8.53, y: y, w: 4.3, h: 0.45, fontFace: F.BODY, fontSize: 14, color: "C7D8E2" }, TB));
        });

      } else {
        page += 1;
        s.background = { color: C.WHITE };
        s.addText(sl.section.toUpperCase(), Object.assign({ x: G.M, y: 0.32, w: 8, h: 0.24, fontFace: F.BODY, fontSize: 10, color: C.TEAL, bold: true, charSpacing: 2.2 }, TB));
        s.addText(sl.title, Object.assign({ x: G.M, y: 0.56, w: 11.2, h: 0.52, fontFace: F.HEAD, fontSize: 27, color: C.INK, bold: true }, TB));
        s.addText(sl.periode, Object.assign({ x: 9.9, y: 0.34, w: 2.85, h: 0.26, align: "right", fontFace: F.BODY, fontSize: 10, color: C.MUTED }, TB));
        s.addText(String(page), Object.assign({ x: 11.9, y: 6.98, w: 0.85, h: 0.26, align: "right", fontFace: F.BODY, fontSize: 10, color: C.MUTED }, TB));
        s.addText("MyAduan Johor  \u00B7  Laporan Pemantauan Aduan", Object.assign({ x: G.M, y: 6.98, w: 7, h: 0.26, fontFace: F.BODY, fontSize: 10, color: C.MUTED }, TB));
        sl.blocks.forEach((b) => block(s, b));
      }
    });

    return pres.writeFile({ fileName: filename });
  };
})();
