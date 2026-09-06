/* render-html.js — Lukis slaid sebagai DOM. Digunakan untuk paparan
   dalam browser DAN untuk PDF (melalui cetakan browser). */

(function () {
  const charts = [];
  const esc = (s) => String(s).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
  const IN = (v) => v + "in";
  const box = (b) => `left:${IN(b.x)};top:${IN(b.y)};width:${IN(b.w)};` + (b.h ? `height:${IN(b.h)};` : "");

  function el(cls, style, html) {
    const d = document.createElement("div");
    d.className = cls;
    if (style) d.setAttribute("style", style);
    if (html != null) d.innerHTML = html;
    return d;
  }

  /* ── Carta ─────────────────────────────────────────────── */
  /* Chart.js mengukur font dalam px; takrifan tema dalam pt. */
  const PX = (pt) => Math.round(pt * (96 / 72));

  /* Label kategori yang panjang dibungkus kepada dua baris supaya
     terbaca penuh tanpa memakan lebar paksi. Chart.js menerima array
     rentetan sebagai label berbilang baris. */
  const HAD_BARIS = 22;
  const bungkus = (t) => {
    const teks = String(t);
    if (teks.length <= HAD_BARIS) return teks;
    const kata = teks.split(/\s+/);
    const baris = ["", ""];
    let i = 0;
    kata.forEach((k) => {
      if (i === 0 && (baris[0] + " " + k).trim().length > HAD_BARIS) i = 1;
      baris[i] = (baris[i] + " " + k).trim();
    });
    if (!baris[1]) return baris[0];
    if (baris[1].length > HAD_BARIS + 8) baris[1] = baris[1].slice(0, HAD_BARIS + 7) + "\u2026";
    return baris;
  };


  /* Hirisan bawah paras ini tidak dilabel - terlalu nipis untuk teks. */
  const HAD_PERATUS = 0.06;

  const NOANIM = {
    animation: false, responsive: true, maintainAspectRatio: false,
    devicePixelRatio: 3,   /* tajam semasa cetakan PDF */
  };
  const AXIS_OFF = { grid: { display: false }, border: { display: false } };

  /* Chart.js mengukur BEKAS, bukan kanvas. Tanpa bekas bersaiz tepat,
     kanvas jatuh ke saiz lalai 300x150 dan carta jadi herot. */
  function mkCanvas(parent, b) {
    const wrap = el("blk chartwrap", box(b));
    const c = document.createElement("canvas");
    wrap.appendChild(c);
    parent.appendChild(wrap);
    return c;
  }

  function drawChart(parent, body) {
    const cv = mkCanvas(parent, body);
    let cfg;

    if (body.k === "col" || body.k === "bar") {
      const horiz = body.k === "bar";
      cfg = {
        type: "bar",
        data: {
          labels: body.labels,
          datasets: [{ data: body.values, backgroundColor: hx(body.color), barPercentage: horiz ? 0.72 : 0.62, categoryPercentage: 0.9 }],
        },
        options: Object.assign({}, NOANIM, {
          indexAxis: horiz ? "y" : "x",
          layout: { padding: horiz ? { right: 34 } : { top: 24 } },
          plugins: {
            legend: { display: false },
            tooltip: { enabled: false },
            datalabels: {
              anchor: "end", align: horiz ? "right" : "top", offset: 2,
              color: hx(C.TEXT), font: { size: PX(FONT.nilai), family: F.BODY },
              formatter: (v) => fmt(v),
            },
          },
          scales: {
            x: horiz
              ? Object.assign({ display: false, beginAtZero: true }, AXIS_OFF)
              : Object.assign({ ticks: { color: hx(C.MUTED), font: { size: PX(FONT.paksi), family: F.BODY } } }, AXIS_OFF),
            y: horiz
              ? Object.assign({
                  afterFit: (sc) => { sc.width = Math.min(sc.width, sc.chart.width * 0.42); },
                  ticks: {
                    autoSkip: false, color: hx(C.TEXT),
                    font: { size: PX(FONT.paksi), family: F.BODY },
                    callback: function (v) { return bungkus(this.getLabelForValue(v)); },
                  },
                }, AXIS_OFF)
              : Object.assign({ display: false, beginAtZero: true }, AXIS_OFF),
          },
        }),
      };
    } else if (body.k === "donut") {
      /* Nombor diletak dalam petunjuk, bukan atas hirisan - hirisan bernilai
         satu atau dua terlalu nipis dan nombornya akan bertindih. */
      /* Peratus dilukis dalam hirisan, tetapi hanya bagi hirisan yang cukup
         besar untuk menampungnya. Hirisan kecil dibiar kosong - nombor
         penuhnya ada dalam petunjuk di sebelah. */
      const jumDonat = body.values.reduce((a, b) => a + b, 0) || 1;
      cfg = {
        type: "doughnut",
        data: {
          labels: body.labels.map((l, i) => l + "  (" + fmt(body.values[i]) + ")"),
          datasets: [{ data: body.values, backgroundColor: (body.colors || SERIES).map(hx), borderWidth: 0 }],
        },
        options: Object.assign({}, NOANIM, {
          cutout: "55%",
          plugins: {
            legend: { position: "right", labels: { boxWidth: 9, boxHeight: 9, color: hx(C.TEXT), font: { size: PX(FONT.petunjuk), family: F.BODY }, padding: 8 } },
            tooltip: { enabled: false },
            datalabels: {
              display: (ctx) => ctx.dataset.data[ctx.dataIndex] / jumDonat >= HAD_PERATUS,
              color: "#fff", font: { size: PX(FONT.nilai), weight: "bold", family: F.BODY },
              formatter: (v) => Math.round((v / jumDonat) * 100) + "%",
            },
          },
        }),
      };
    } else if (body.k === "line") {
      cfg = {
        type: "line",
        data: {
          labels: body.labels,
          datasets: body.series.map((s) => ({
            label: s.name, data: s.values,
            borderColor: hx(s.color), backgroundColor: hx(s.color),
            borderWidth: 2.5, pointRadius: 3.5, tension: 0.15, fill: false,
          })),
        },
        options: Object.assign({}, NOANIM, {
          layout: { padding: { top: 18 } },
          plugins: {
            legend: { position: "top", labels: { boxWidth: 14, color: hx(C.TEXT), font: { size: PX(FONT.paksi), family: F.BODY } } },
            tooltip: { enabled: false },
            datalabels: {
              display: true, align: "top", offset: 5,
              color: hx(C.TEXT), font: { size: PX(FONT.nilai), family: F.BODY },
              formatter: (v) => fmt(v),
            },
          },
          scales: {
            x: Object.assign({ ticks: { color: hx(C.MUTED), font: { size: PX(FONT.paksi), family: F.BODY } } }, AXIS_OFF),
            y: { beginAtZero: true, grid: { color: hx(C.LINE) }, border: { display: false }, ticks: { color: hx(C.MUTED), font: { size: PX(FONT.nilai), family: F.BODY } } },
          },
        }),
      };
    }
    charts.push(new Chart(cv, cfg));
  }

  /* ── Blok ──────────────────────────────────────────────── */
  function drawBlock(slide, b) {
    if (b.t === "stat") {
      const w = el("blk stat", `${box(b)};background:${hx(b.fill || C.MIST)}`);
      w.innerHTML =
        `<div class="stat-label" style="color:${hx(b.labelColor || C.MUTED)}">${esc(b.label)}</div>` +
        `<div class="stat-value" style="color:${hx(b.color || C.INK)};font-size:${b.vsize || 40}pt">${esc(fmt(b.value))}</div>` +
        (b.sub ? `<div class="stat-sub" style="color:${hx(b.subColor || C.MUTED)}">${esc(b.sub)}</div>` : "");
      slide.appendChild(w);

    } else if (b.t === "card") {
      const w = el("blk card", box(b));
      w.innerHTML = `<div class="card-title">${esc(b.title || "")}</div>` +
        (b.note ? `<div class="card-note">${esc(b.note)}</div>` : "");
      slide.appendChild(w);
      if (b.body) {
        if (['col','bar','donut','line'].indexOf(b.body.k) >= 0) drawChart(slide, b.body);
        else drawBlock(slide, Object.assign({ t: b.body.k }, b.body));
      }

    } else if (b.t === "segment") {
      const tot = b.segs.reduce((a, s) => a + s.value, 0) || 1;
      const w = el("blk segwrap", box(b));
      let bar = `<div class="segbar" style="height:${IN(b.barH || 0.46)}">`;
      b.segs.forEach((s) => {
        const p = (s.value / tot) * 100;
        bar += `<div class="seg" style="width:${p}%;background:${hx(s.color)}">` +
          (p > 7 ? `<span>${fmt(s.value)}</span>` : "") + "</div>";
      });
      bar += "</div><div class='seglegend'>";
      b.segs.forEach((s) => {
        bar += `<div class="segleg"><i style="background:${hx(s.color)}"></i>${esc(s.label)}</div>`;
      });
      w.innerHTML = bar + "</div>";
      slide.appendChild(w);

    } else if (b.t === "insight") {
      const w = el("blk insight", box(b));
      w.innerHTML = `<i></i><span>${esc(b.text)}</span>`;
      slide.appendChild(w);

    } else if (b.t === "funnel") {
      const gap = 0.5, bw = (b.w - gap * (b.steps.length - 1)) / b.steps.length;
      b.steps.forEach((st, i) => {
        const bx = b.x + i * (bw + gap);
        const w = el("blk fstep", `left:${IN(bx)};top:${IN(b.y)};width:${IN(bw)};height:${IN(b.h)};background:${hx(st.fill)}`);
        w.innerHTML =
          `<div class="stat-label" style="color:${hx(st.labelColor)}">${esc(st.label).replace(/\n/g, "<br>")}</div>` +
          `<div class="fstep-value" style="color:${hx(st.valueColor)}">${esc(fmt(st.value))}</div>` +
          (st.note ? `<div class="stat-sub" style="color:${hx(st.noteColor)}">${esc(st.note)}</div>` : "");
        slide.appendChild(w);
        if (i < b.steps.length - 1) {
          slide.appendChild(el("blk farrow", `left:${IN(bx + bw + 0.09)};top:${IN(b.y + b.h / 2 - 0.13)};width:${IN(0.32)};height:${IN(0.26)}`));
        }
      });

    } else if (b.t === "table") {
      const isLeft = (i) => (b.leftCols ? b.leftCols.indexOf(i) >= 0 : i === 0);
      const w = el("blk", box(b));
      let h = `<table class="tbl" style="font-size:${(b.fs || 10.5) * 0.95}pt"><colgroup>` +
        b.colW.map((c) => `<col style="width:${IN(c)}">`).join("") + "</colgroup><thead><tr>";
      b.head.forEach((t, i) => { h += `<th class="${isLeft(i) ? "" : "c"}">${esc(t)}</th>`; });
      h += "</tr></thead><tbody>";
      b.rows.forEach((r) => {
        h += `<tr style="height:${IN(b.rowH || 0.32)}">`;
        r.forEach((cell, i) => {
          const o = typeof cell === "object" && cell !== null ? cell : { text: cell };
          const teks = /^\d{4,}$/.test(String(o.text)) ? fmt(o.text) : o.text;
          h += `<td class="${isLeft(i) ? "" : "c"}" style="color:${hx(o.color || C.TEXT)};${o.bold || isLeft(i) ? "font-weight:700" : ""}">${esc(teks)}</td>`;
        });
        h += "</tr>";
      });
      w.innerHTML = h + "</tbody></table>";
      slide.appendChild(w);

    } else if (b.t === "rank") {
      const max = Math.max.apply(null, b.items.map((i) => i[1])) || 1;
      const w = el("blk", box(b));
      let h = "";
      b.items.forEach((it) => {
        h += `<div class="rank"><span class="rl">${esc(it[0])}</span>` +
          `<span class="rb"><i style="width:${(it[1] / max) * 100}%;background:${hx(b.color || C.OCEAN)}"></i></span>` +
          `<span class="rv">${fmt(it[1])}</span></div>`;
      });
      w.innerHTML = h;
      slide.appendChild(w);

    } else if (b.t === "label") {
      slide.appendChild(el("blk minilabel", box(b), esc(b.text)));

    } else if (b.t === "finding") {
      if (b.text) slide.appendChild(el("blk finding", box(b), esc(b.text)));

    } else if (b.t === "para") {
      if (b.text) slide.appendChild(el("blk para", box(b), esc(b.text).replace(/\n/g, "<br>")));
    }
  }

  /* ── Slaid ─────────────────────────────────────────────── */
  window.renderHTML = function (spec, mount) {
    charts.forEach((c) => c.destroy());
    charts.length = 0;
    mount.innerHTML = "";
    let page = 0;

    spec.forEach((sl) => {
      const s = el("slide" + (sl.kind === "content" ? "" : " dark"));

      if (sl.kind === "title") {
        const i = sl.info;
        s.innerHTML =
          `<div class="t-band"></div>` +
          `<div class="t-eyebrow">${esc(i.sistem)}</div>` +
          `<div class="t-title">${esc(i.tajuk)}</div>` +
          `<div class="t-period">${esc(i.periode)}</div>` +
          `<div class="t-foot1"><span>Disediakan oleh</span> <b>${esc(i.disediakan)}</b></div>` +
          `<div class="t-foot2"><span>Sumber data</span> ${esc(i.sumber)}</div>`;

      } else if (sl.kind === "section") {
        s.innerHTML =
          `<div class="s-num">${sl.num}</div><div class="s-title">${esc(sl.title)}</div>` +
          `<div class="s-items">` + sl.items.map((t) => `<div class="s-item"><i></i>${esc(t)}</div>`).join("") + "</div>";

      } else {
        page += 1;
        s.innerHTML =
          `<div class="c-eyebrow">${esc(sl.section)}</div>` +
          `<div class="c-title">${esc(sl.title)}</div>` +
          `<div class="c-period">${esc(sl.periode)}</div>` +
          `<div class="c-foot">MyAduan Johor &nbsp;&middot;&nbsp; Laporan Pemantauan Aduan</div>` +
          `<div class="c-page">${page}</div>`;
      }
      mount.appendChild(s);
      if (sl.blocks) sl.blocks.forEach((b) => drawBlock(s, b));
    });
    return page;
  };
})();
