/* app.js - Pengikat antara muka: muat naik -> takrifan slaid -> paparan & eksport. */

(function () {
  let DATA = JSON.parse(JSON.stringify(window.DEFAULT_DATA));
  let SPEC = [];
  let BUKU_MENTAH = null;     /* fail mentah disimpan supaya bulan boleh ditukar */

  /* Empat slot muat naik. 'terima' menyenaraikan jenis helaian yang sah
     bagi slot itu, supaya fail yang tersalah slot dapat dikesan awal. */
  const SLOT = [
    { id: "senarai", tajuk: "Senarai Aduan (Raw Data)",
      desc: "Eksport penuh dari sistem MAJ. Mengisi seksyen Status Semasa Aduan ikut PBT. Fail KML sempadan ahli majlis juga diterima di sini.",
      terima: ["senarai_penuh", "senarai", "ahli_map"] },
    { id: "cs", tajuk: "Customer Service",
      desc: "Helaian SISPAA dan EmelWS. Mengisi seksyen Customer Service.",
      terima: ["cs_sispaa", "cs_emel"] },
    { id: "media", tajuk: "Media Sosial",
      desc: "Log aduan media sosial dengan Case ID. Mengisi seksyen Media Sosial.",
      terima: ["media"] },
    { id: "lapangan", tajuk: "Lapangan",
      desc: "Helaian Aduan Pencegahan dan Lawatan Sistem (QC). Mengisi seksyen Lapangan.",
      terima: ["lap_cegah", "lap_qc"] },
  ];
  const NAMA_JENIS = {
    senarai_penuh: "senarai aduan penuh", senarai: "senarai aduan ringkas",
    ahli_map: "pemetaan zon ahli majlis",
    cs_sispaa: "SISPAA", cs_emel: "e-mel & WhatsApp",
    media: "media sosial", lap_cegah: "aduan pencegahan", lap_qc: "lawatan tapak (QC)",
  };
  const simpanan = {};   /* id slot -> [{nama, wb, jenis, sah}] */

  const $ = (id) => document.getElementById(id);
  const deck = $("deck"), notes = $("notes"), selBulan = $("selBulan"), panel = $("panel");

  function note(cls, html) {
    const d = document.createElement("div");
    d.className = "note " + cls;
    d.innerHTML = html;
    notes.appendChild(d);
  }
  const clearNotes = () => { notes.innerHTML = ""; };
  const esc = (s) => String(s).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
  const senarai = (arr, had) => "<ul>" + arr.slice(0, had || 10).map((x) => "<li>" + x + "</li>").join("") +
    (arr.length > (had || 10) ? "<li>... dan " + (arr.length - (had || 10)) + " lagi.</li>" : "") + "</ul>";

  /* ── Bina panel muat naik ───────────────────────────── */
  const slots = $("slots");
  SLOT.forEach((s, i) => {
    simpanan[s.id] = [];
    const el = document.createElement("div");
    el.className = "slot";
    el.id = "slot-" + s.id;
    el.innerHTML =
      '<div class="slot-num">' + (i + 1) + "</div>" +
      '<div class="slot-title">' + s.tajuk + "</div>" +
      '<div class="slot-desc">' + s.desc + "</div>" +
      '<button class="slot-btn" type="button">Pilih fail</button>' +
      '<div class="slot-files">Tiada fail</div>' +
      '<input type="file" accept=".xlsx,.xls,.kml" multiple>';
    slots.appendChild(el);

    const input = el.querySelector("input");
    el.querySelector(".slot-btn").addEventListener("click", () => input.click());
    input.addEventListener("change", (e) => muatSlot(s, Array.prototype.slice.call(e.target.files), e.target));
  });

  function lukisSlot(s) {
    const el = $("slot-" + s.id);
    const fail = simpanan[s.id];
    const box = el.querySelector(".slot-files");
    el.classList.toggle("isi", fail.length > 0);
    if (!fail.length) { box.textContent = "Tiada fail"; return; }
    box.innerHTML = "<b>" + fail.length + " fail dimuat naik</b>" +
      fail.map((f) => '<span class="fail' + (f.sah ? "" : " salah") + '">' +
        (f.sah ? "" : "\u26A0 ") + esc(f.nama) + "</span>").join("");
    kemasStatus();
  }

  function kemasStatus() {
    const jumlah = SLOT.reduce((a, s) => a + simpanan[s.id].length, 0);
    const salah = SLOT.reduce((a, s) => a + simpanan[s.id].filter((f) => !f.sah).length, 0);
    $("btnJana").disabled = jumlah === 0;
    $("janaNota").textContent = jumlah === 0
      ? "Belum ada fail dimuat naik."
      : jumlah + " fail sedia" + (salah ? " \u00B7 " + salah + " fail mungkin tersalah slot" : "") + ".";
  }

  function muatSlot(s, files, input) {
    if (!files.length) return;
    Promise.all(files.map((f) => new Promise((res, rej) => {
      const kml = /\.kml$/i.test(f.name);
      const fr = new FileReader();
      fr.onload = (ev) => {
        try {
          let wb;
          if (kml) {
            /* KML ditukar kepada buku kerja satu helaian supaya laluan
               pemprosesan seterusnya kekal sama dengan fail Excel. */
            const baris = parseKMLAhli(ev.target.result);
            if (!baris.length) throw new Error("tiada zon ahli majlis dijumpai");
            wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(baris), "Ahli_Majlis");
          } else {
            wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array", cellDates: true });
          }
          const jenis = kenalPastiFail(wb);
          res({ nama: f.name, wb: wb, jenis: jenis,
            sah: jenis.some((j) => s.terima.indexOf(j) >= 0) });
        } catch (err) { rej(new Error(f.name + ": " + err.message)); }
      };
      fr.onerror = () => rej(new Error("Gagal membaca " + f.name));
      if (kml) fr.readAsText(f); else fr.readAsArrayBuffer(f);
    }))).then((hasil) => {
      clearNotes();
      hasil.forEach((h) => {
        const sedia = simpanan[s.id].filter((x) => x.nama === h.nama).length;
        if (!sedia) simpanan[s.id].push(h);
      });
      lukisSlot(s);
      const salah = hasil.filter((h) => !h.sah);
      if (salah.length) {
        note("warn", "<b>Fail mungkin tersalah slot.</b>" + senarai(salah.map((h) =>
          esc(h.nama) + " dikenali sebagai " +
          (h.jenis.length ? h.jenis.map((j) => NAMA_JENIS[j] || j).join(", ") : "tidak dikenali") +
          ", bukan " + s.tajuk.toLowerCase())) +
          "Fail tetap akan diproses mengikut kandungannya, bukan mengikut slot.");
      }
    }).catch((err) => { clearNotes(); note("err", "<b>Gagal membaca fail.</b> " + err.message); })
      .then(() => { if (input) input.value = ""; });
  }

  $("btnKosong").addEventListener("click", function () {
    SLOT.forEach((s) => { simpanan[s.id] = []; lukisSlot(s); });
    clearNotes();
    kemasStatus();
  });

  $("btnPanel").addEventListener("click", () => panel.classList.toggle("tutup"));

  /* ── Jana slaid ─────────────────────────────────────── */
  $("btnJana").addEventListener("click", function () {
    const buku = [];
    SLOT.forEach((s) => simpanan[s.id].forEach((f) => buku.push({ nama: f.nama, wb: f.wb })));
    if (!buku.length) return;
    BUKU_MENTAH = buku;
    try {
      prosesMentah(null);
      panel.classList.add("tutup");
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (err) {
      clearNotes();
      note("err", "<b>Gagal memproses fail.</b> " + err.message);
    }
  });

  function prosesMentah(bulanPilihan) {
    const res = ingestRaw(BUKU_MENTAH, bulanPilihan);
    DATA = res.data;
    render();
    clearNotes();
    note("ok", "<b>" + BUKU_MENTAH.length + " fail diproses</b> untuk tempoh <b>" +
      res.data.info.periode + "</b>." + senarai(res.dikenali));
    if (res.warn.length) note("warn", "<b>Perlu perhatian:</b>" + senarai(res.warn));
    note("ok", "Tukar tempoh melalui pemilih bulan di bar atas. Seksyen tanpa data " +
      "dilangkau automatik &mdash; muat naik fail berkenaan untuk memaparkannya.");
    isiPemilihBulan(res.bulanTersedia, res.bulan);
  }

  /* ── Pemilih bulan ──────────────────────────────────── */
  function isiPemilihBulan(senaraiBulan, dipilih) {
    selBulan.innerHTML = "";
    senaraiBulan.forEach((b) => {
      const o = document.createElement("option");
      o.value = b.kunci; o.textContent = b.label;
      if (b.kunci === dipilih) o.selected = true;
      selBulan.appendChild(o);
    });
    selBulan.disabled = false;
  }
  selBulan.disabled = true;
  selBulan.addEventListener("change", function () {
    if (BUKU_MENTAH) prosesMentah(this.value);
  });

  /* ── Paparan ────────────────────────────────────────── */
  function fitSlides() {
    const avail = Math.min(window.innerWidth - 48, 1240);
    const scale = Math.min(1, avail / (13.333 * 96));
    document.querySelectorAll(".slide").forEach((s) => {
      s.style.transform = "scale(" + scale + ")";
      s.style.marginBottom = (7.5 * 96 * (scale - 1)) + "px";
    });
  }

  function render() {
    SPEC = buildSpec(DATA);
    renderHTML(SPEC, deck);
    $("count").textContent = SPEC.length + " slaid";
    $("period").textContent = DATA.info.periode || "-";
    fitSlides();
  }

  /* ── Templat Excel ──────────────────────────────────── */
  $("btnUpload").addEventListener("click", () => $("file").click());

  $("file").addEventListener("change", function (e) {
    const f = e.target.files[0];
    if (!f) return;
    clearNotes();
    const fr = new FileReader();
    fr.onload = function (ev) {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array", cellDates: true });
        const res = parseWorkbook(wb);
        DATA = res.data;
        BUKU_MENTAH = null;
        selBulan.disabled = true;
        selBulan.innerHTML = "";
        render();
        panel.classList.add("tutup");
        note("ok", "<b>" + esc(f.name) + "</b> berjaya dimuat naik. Slaid telah dikemas kini.");
        if (res.warn.length) note("warn", "<b>Perlu perhatian:</b>" + senarai(res.warn, 12));
      } catch (err) {
        note("err", "<b>Gagal membaca fail.</b> " + err.message);
      }
      e.target.value = "";
    };
    fr.readAsArrayBuffer(f);
  });

  $("btnTemplate").addEventListener("click", function () {
    buildTemplate(DATA, "Templat_Laporan_MAJ.xlsx");
  });

  /* ── Eksport ────────────────────────────────────────── */
  $("btnPptx").addEventListener("click", function () {
    const b = this;
    b.disabled = true; b.textContent = "Menjana...";
    const name = "Laporan_MAJ_" + String(DATA.info.periode || "Laporan").replace(/\s+/g, "_") + ".pptx";
    renderPPTX(SPEC, name)
      .then(() => { clearNotes(); note("ok", "Fail PowerPoint <b>" + name + "</b> telah dimuat turun."); })
      .catch((err) => { clearNotes(); note("err", "<b>Gagal menjana PPTX.</b> " + err.message); })
      .then(() => { b.disabled = false; b.textContent = "Muat turun PPTX"; });
  });

  $("btnPdf").addEventListener("click", function () {
    clearNotes();
    note("ok", "Dalam kotak cetakan: pilih <b>Save as PDF</b>, saiz kertas <b>Custom 13.33 x 7.5 inci</b> " +
      "(atau Landscape), dan matikan <b>Headers and footers</b>.");
    setTimeout(() => window.print(), 350);
  });

  window.addEventListener("resize", fitSlides);
  SLOT.forEach(lukisSlot);
  kemasStatus();
  render();
  note("ok", "Memaparkan <b>data demo</b>. Muat naik fail eksport mengikut slot di atas, " +
    "kemudian tekan <b>Jana slaid</b>.");
})();
