/* spec.js - Takrifan slaid.
   INI SATU-SATUNYA tempat susun atur ditakrifkan; penjana HTML dan PPTX
   kedua-duanya membacanya.

   Setiap tajuk dan setiap ayat dapatan mempunyai ID. Ayat automatik
   dijana daripada data, tetapi boleh ditindih melalui helaian "Teks"
   dalam fail Excel. */

(function () {
  const L = (o) => Object.keys(o || {});
  const V = (o) => Object.values(o || {});
  const num = (v) => (typeof v === "number" ? v : parseFloat(v) || 0);
  const pct = (a, b) => (b ? Math.round((a / b) * 100) : 0);
  const dtxt = (d) => (d >= 0 ? "\u25B2" : "\u25BC") + " " + Math.abs(d).toFixed(1) + "%  vs bulan lepas";
  const dcol = (d) => (d >= 0 ? C.MOSS : C.CRIMSON);
  const ageSegs = (o) => L(o).map((l, i) => ({ label: l, value: num(o[l]), color: AGEING_COLORS[i] }));

  window.buildSpec = function (D) {
    /* "Tidak dinyatakan" bukan satu tahap kesukaran sebenar - buang
       supaya carta hanya menunjukkan tahap yang bermakna. */
    const tanpaKosong = (o) => {
      const x = {};
      Object.keys(o || {}).forEach((k) => { if (!/tidak dinyatakan/i.test(k)) x[k] = o[k]; });
      return x;
    };
    const g = (sek, st) => (D.d[sek] && D.d[sek][st]) || {};
    /* Setiap siri trend membawa senarai bulannya sendiri, kerana
       seksyen berbeza boleh meliputi tempoh berbeza. */
    const TR = (k) => {
      const s = D.trend.siri[k] || { bulan: [], nilai: [] };
      return { bulan: s.bulan || [], nilai: (s.nilai || []).map(num) };
    };
    const periode = D.info.periode;
    const teks = D.teks || {};
    const REG = [];
    const slides = [];

    /* Teks boleh ubah: kosong dalam Excel = guna ayat automatik */
    function T(key, auto, desc) {
      REG.push({ key: key, desc: desc || "", auto: String(auto == null ? "" : auto) });
      const v = teks[key];
      return v == null || String(v).trim() === "" ? auto : String(v);
    }

    /* Kad + carta + ayat dapatan. Kedudukan carta dikira automatik. */
    function CC(o) {
      const headH = o.note ? 0.78 : 0.58;
      const findH = o.find === null ? 0 : (o.findH || 0.5);
      const pad = 0.22;
      const ix = o.x + pad, iw = o.w - pad * 2;
      const iy = o.y + headH, ih = o.h - headH - findH - 0.1;
      const card = {
        t: "card", x: o.x, y: o.y, w: o.w, h: o.h,
        title: T(o.id + ".tajuk", o.title, "Tajuk kad"),
        note: o.note ? T(o.id + ".nota", o.note, "Nota kecil") : null,
      };
      const out = [card];
      if (o.body) {
        if (o.body.k === "segment") {
          out.push({ t: "segment", x: ix + 0.06, y: iy + (ih - 0.85) / 2,
            w: iw - 0.12, barH: 0.6, segs: o.body.segs });
        } else {
          card.body = Object.assign({}, o.body, { x: ix, y: iy, w: iw, h: ih });
        }
      }
      if (findH) {
        out.push({ t: "finding", x: o.x + 0.28, y: o.y + o.h - findH + 0.02,
          w: o.w - 0.56, h: findH - 0.06,
          text: T(o.id + ".dapatan", o.find, "Ayat dapatan") });
      }
      return out;
    }

    const SW = 3.9, SG = 0.25, HW = (G.COL - 0.25) / 2;
    const R1 = 1.25, RH = 2.72, R2 = 4.13;   /* grid dua baris */

    /* ═══ TAJUK ═══ */
    slides.push({ kind: "title", info: D.info });


    /* Adakah set data ini wujud dan berisi? Digunakan untuk melangkau
       blok atau slaid apabila fail eksport tidak mengandungi medan berkenaan. */
    const ada = (o) => o && Object.keys(o).length > 0 &&
      Object.keys(o).reduce((a, k) => a + num(o[k]), 0) > 0;

    /* ═══════════ 1. MEDIA SOSIAL ═══════════ */
    const mR = g("Media", "Ringkasan");
    const mPlat = g("Media", "Platform"), mKat = g("Media", "Kategori");
    const mPBT = g("Media", "PBT"), mHal = g("Media", "Halaman");
    const mStat = g("Media", "Status"), mAge = g("Media", "Ageing");
    const mDMp = g("Media", "DM_Platform");
    const mCap = num(mR.Capture), mDM = num(mR.DM);
    const mDaftar = num(mR.Daftar_MAJ) || mCap;
    const mAktif = num(mR.Aktif), mSel = num(mR.Selesai);

    if (mCap) {
      slides.push({
        kind: "section", num: 1, title: T("S1.tajuk", "Media Sosial", "Tajuk seksyen"),
        items: ["Extract aduan dari media sosial", "Daftarkan aduan ke dalam MAJ",
          "Respons komen pengadu", "Jawab DM & pertanyaan"],
      });

      /* 1 - Statistik Keseluruhan */
      slides.push({
        kind: "content", section: "Media Sosial",
        title: T("M1.tajuk", "Statistik Keseluruhan Gerak Kerja Media Sosial", "Tajuk slaid"), periode,
        blocks: [
          { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Aduan di-capture", value: mCap, color: C.INK,
            sub: mR.Capture_Delta == null ? "Bulan " + periode : dtxt(num(mR.Capture_Delta)),
            subColor: mR.Capture_Delta == null ? C.MUTED : dcol(num(mR.Capture_Delta)) },
          { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Didaftarkan dalam MAJ", value: mDaftar, color: C.OCEAN, sub: pct(mDaftar, mCap) + "% daripada aduan di-capture" },
          { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Purata aduan sehari", value: Math.round(mCap / 21), fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: "Berasaskan 21 hari bekerja", subColor: C.ICE },
        ].concat(CC({
          id: "M1.trend", x: G.M, y: 2.85, w: G.CW, h: 4.0,
          title: "Trend bulanan aduan di-capture", note: "Enam bulan terakhir",
          body: { k: "col", labels: TR("Media_Capture").bulan, values: TR("Media_Capture").nilai, color: C.TEAL },
          find: fTrend(TR("Media_Capture").nilai, TR("Media_Capture").bulan, "aduan dijumpai"),
        })),
      });

      /* 2 - Pecahan Sumber & Kategori */
      const kedua = ada(mPBT) ? { set: mPBT, tajuk: "Pecahan ikut PBT", jenis: "col" }
        : ada(mHal) ? { set: mHal, tajuk: "Halaman / kumpulan sumber", jenis: "bar" } : null;
      slides.push({
        kind: "content", section: "Media Sosial",
        title: T("M2.tajuk", "Pecahan Sumber & Kategori", "Tajuk slaid"), periode,
        blocks: (kedua
          ? CC({ id: "M2.platform", x: G.M, y: R1, w: G.COL, h: RH, findH: 0.44,
              title: "Sumber platform",
              body: { k: "donut", labels: L(mPlat), values: V(mPlat).map(num), colors: SERIES },
              find: fTop(mPlat, "aduan") }).concat(CC({
              id: "M2.kedua", x: G.M, y: R2, w: G.COL, h: RH, findH: 0.44,
              title: kedua.tajuk,
              body: { k: kedua.jenis, labels: L(kedua.set), values: V(kedua.set).map(num), color: C.SLATE },
              find: fTop(kedua.set, "aduan") }))
          : CC({ id: "M2.platform", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
              title: "Sumber platform",
              body: { k: "donut", labels: L(mPlat), values: V(mPlat).map(num), colors: SERIES },
              find: fTop(mPlat, "aduan") })
        ).concat(CC({
          id: "M2.kategori", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
          title: "Kategori aduan dijumpai", note: "Sepuluh teratas",
          body: { k: "bar", labels: L(mKat), values: V(mKat).map(num), color: C.OCEAN },
          find: fTop(mKat, "aduan"),
        })),
      });

      /* 3 - Status Semasa Aduan */
      if (ada(mStat)) {
        const kanan = ada(mAge)
          ? [{ t: "stat", x: G.COL_X2, y: G.TOP, w: HW, h: 1.35, label: "Masih aktif", value: mAktif, color: C.AMBER, sub: pct(mAktif, mDaftar) + "% daripada aduan didaftar" },
             { t: "stat", x: G.COL_X2 + HW + 0.25, y: G.TOP, w: HW, h: 1.35, label: "Telah selesai", value: mSel, color: C.MOSS, sub: pct(mSel, mDaftar) + "% daripada aduan didaftar" }]
              .concat(CC({ id: "M3.ageing", x: G.COL_X2, y: 2.85, w: G.COL, h: 4.0,
                title: "Usia aduan masih aktif", note: "Tempoh aduan tergantung sejak didaftarkan",
                body: { k: "segment", segs: ageSegs(mAge) }, find: fAge(mAge) }))
          : [{ t: "stat", x: G.COL_X2, y: G.TOP, w: G.COL, h: 1.75, label: "Masih aktif", value: mAktif, fill: C.MIST2, color: C.AMBER, vsize: 48, sub: pct(mAktif, mDaftar) + "% daripada aduan didaftarkan masih dalam proses" },
             { t: "stat", x: G.COL_X2, y: 3.25, w: G.COL, h: 1.75, label: "Telah selesai atau ditolak", value: mSel, fill: C.MIST2, color: C.MOSS, vsize: 48, sub: pct(mSel, mDaftar) + "% daripada aduan didaftarkan telah tamat" },
             { t: "insight", x: G.COL_X2, y: 5.5, w: G.COL, text: T("M3.insight", "Status diambil daripada senarai aduan MAJ melalui padanan ID aduan.", "Ayat insight") }];
        slides.push({
          kind: "content", section: "Media Sosial",
          title: T("M3.tajuk", "Status Semasa Aduan", "Tajuk slaid"), periode,
          blocks: CC({
            id: "M3.status", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
            title: "Pecahan status", note: fmt(mDaftar) + " aduan didaftarkan melalui media sosial",
            body: { k: "donut", labels: L(mStat), values: V(mStat).map(num), colors: warnaStatus(L(mStat)) },
            find: fStatus(mStat, ["Diselesaikan", "Ditolak"]),
          }).concat(kanan),
        });
      }

      /* 4 - DM & Pertanyaan (hanya jika data DM tersedia) */
      if (mDM) {
        slides.push({
          kind: "content", section: "Media Sosial",
          title: T("M4.tajuk", "DM & Pertanyaan", "Tajuk slaid"), periode,
          blocks: [
            { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "DM diterima", value: mDM, color: C.INK, sub: dtxt(num(mR.DM_Delta)), subColor: dcol(num(mR.DM_Delta)) },
            { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Dijadikan aduan rasmi", value: num(mR.DM_Jadi_Aduan), color: C.OCEAN, sub: pct(num(mR.DM_Jadi_Aduan), mDM) + "% daripada DM diterima" },
            { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Selesai terus di DM", value: num(mR.DM_Selesai_Terus), color: C.MOSS, sub: "Tidak perlu didaftarkan sebagai aduan" },
          ].concat(CC({
            id: "M4.trend", x: G.M, y: 2.85, w: G.COL, h: 4.0, title: "Trend bulanan DM diterima",
            body: { k: "col", labels: TR("Media_DM").bulan, values: TR("Media_DM").nilai, color: C.SLATE },
            find: fTrend(TR("Media_DM").nilai, TR("Media_DM").bulan, "DM"),
          })).concat(CC({
            id: "M4.platform", x: G.COL_X2, y: 2.85, w: G.COL, h: 4.0, title: "Pecahan DM ikut platform",
            body: { k: "donut", labels: L(mDMp), values: V(mDMp).map(num), colors: SERIES },
            find: fTop(mDMp, "DM"),
          })),
        });
      }
    }

    /* ═══════════ 2. CUSTOMER SERVICE ═══════════ */
    const cR = g("CS", "Ringkasan");
    const cDit = num(cR.Diterima), cAkt = num(cR.Aktif);
    const cTut = num(cR.Ditutup), cBaki = num(cR.Belum_Tutup);
    const cAge = g("CS", "Ageing_Aktif");
    const cPBT = g("CS", "PBT"), cKat = g("CS", "Kategori"), cStat = g("CS", "Status_Aktif");
    const cIsu = g("CS", "Isu_Emel");
    const cEmel = num(cR.Emel_Dijawab), cWA = num(cR.WA_Dijawab);

    if (cDit) {
      slides.push({
        kind: "section", num: 2, title: T("S2.tajuk", "Perkhidmatan Pelanggan", "Tajuk seksyen"),
        items: ["Daftar aduan SISPAA ke MAJ", "Tutup aduan MAJ kembali ke SISPAA",
          "Jawab e-mel pengadu", "Jawab pertanyaan WhatsApp"],
      });

      slides.push({
        kind: "content", section: "Perkhidmatan Pelanggan",
        title: T("C1.tajuk", "Statistik Keseluruhan Pengurusan Aduan SISPAA", "Tajuk slaid"), periode,
        blocks: [{
          t: "funnel", x: G.M, y: G.TOP, w: G.CW, h: 1.9, steps: [
            { label: "Diterima &\ndidaftar", value: cDit, fill: C.INK, labelColor: C.ICE, valueColor: C.WHITE, note: cR.Diterima_Delta == null ? "Bulan " + periode : dtxt(num(cR.Diterima_Delta)), noteColor: C.ICE },
            { label: "Masih aktif\ndi MAJ", value: cAkt, fill: C.MIST, labelColor: C.MUTED, valueColor: C.AMBER, note: "Kerja belum siap", noteColor: C.MUTED },
            { label: "Ditutup di\nSISPAA", value: cTut, fill: C.MIST, labelColor: C.MUTED, valueColor: C.MOSS, note: "Kitaran lengkap", noteColor: C.MUTED },
            { label: "Siap tetapi\nbelum ditutup", value: cBaki, fill: C.PALE_RED, labelColor: C.DARK_RED, valueColor: C.CRIMSON, note: "Perlu tindakan Perkhidmatan Pelanggan", noteColor: C.DARK_RED },
          ],
        }].concat(CC({
          id: "C1.trend", x: G.M, y: 3.42, w: G.CW, h: 3.43,
          title: "Trend bulanan aduan SISPAA diterima",
          body: { k: "col", labels: TR("CS_Diterima").bulan, values: TR("CS_Diterima").nilai, color: C.TEAL },
          find: fTrend(TR("CS_Diterima").nilai, TR("CS_Diterima").bulan, "aduan"),
        })),
      });

      slides.push({
        kind: "content", section: "Perkhidmatan Pelanggan",
        title: T("C2.tajuk", "Pecahan Aduan SISPAA", "Tajuk slaid"), periode,
        blocks: (ada(cPBT)
          ? CC({ id: "C2.pbt", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
              title: "Pecahan ikut PBT", note: fmt(cDit) + " aduan diterima",
              body: { k: "donut", labels: L(cPBT), values: V(cPBT).map(num), colors: SERIES },
              find: fTop(cPBT, "aduan") }).concat(CC({
              id: "C2.kategori", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
              title: "Kategori aduan", note: "Sepuluh teratas",
              body: { k: "bar", labels: L(cKat), values: V(cKat).map(num), color: C.OCEAN },
              find: fTop(cKat, "aduan") }))
          : CC({ id: "C2.kategori", x: G.M, y: G.TOP, w: G.CW, h: 5.60,
              title: "Kategori aduan", note: fmt(cDit) + " aduan diterima pada " + periode,
              body: { k: "bar", labels: L(cKat), values: V(cKat).map(num), color: C.OCEAN },
              find: fTop(cKat, "aduan") })),
      });

      if (ada(cStat)) {
        const cLewat = num(cAge["16-30 hari"]) + num(cAge["30+ hari"]);
        const kananCS = ada(cAge)
          ? CC({ id: "C3.ageing", x: G.COL_X2, y: G.TOP, w: G.COL, h: RH,
              title: "Usia aduan aktif", note: "Tempoh aduan tergantung",
              body: { k: "segment", segs: ageSegs(cAge) }, find: fAge(cAge) }).concat([
              { t: "stat", x: G.COL_X2, y: R2, w: HW, h: 1.5, label: "Melebihi 15 hari", value: cLewat, color: C.AMBER, sub: pct(cLewat, cAkt) + "% daripada aduan aktif" },
              { t: "stat", x: G.COL_X2 + HW + 0.25, y: R2, w: HW, h: 1.5, label: "Melebihi 30 hari", value: num(cAge["30+ hari"]), color: C.CRIMSON, sub: "Perlu diangkat dalam mesyuarat" },
              { t: "insight", x: G.COL_X2, y: 5.95, w: G.COL, text: T("C3.insight", "Aduan yang telah selesai dikeluarkan dari kiraan ini - elak kira dua kali.", "Ayat insight") }])
          : [{ t: "stat", x: G.COL_X2, y: G.TOP, w: G.COL, h: 1.75, label: "Masih aktif di MAJ", value: cAkt, fill: C.MIST2, color: C.AMBER, vsize: 48, sub: pct(cAkt, cDit) + "% daripada aduan diterima bulan ini" },
             { t: "stat", x: G.COL_X2, y: 3.25, w: G.COL, h: 1.75, label: "Siap tetapi belum ditutup di SISPAA", value: cBaki, fill: C.PALE_RED, color: C.CRIMSON, labelColor: C.DARK_RED, vsize: 48, sub: "Kerja telah siap, rekod belum dikemaskini", subColor: C.DARK_RED },
             { t: "insight", x: G.COL_X2, y: 5.5, w: G.COL, text: T("C3.insight", "Aduan yang telah selesai dikeluarkan dari kiraan ini - elak kira dua kali.", "Ayat insight") }];
        slides.push({
          kind: "content", section: "Perkhidmatan Pelanggan",
          title: T("C3.tajuk", "Status Semasa Aduan Aktif", "Tajuk slaid"), periode,
          blocks: CC({
            id: "C3.status", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
            title: "Pecahan status", note: fmt(cAkt) + " aduan masih dalam proses",
            body: { k: "donut", labels: L(cStat), values: V(cStat).map(num), colors: warnaStatus(L(cStat)) },
            find: fTop(cStat, "aduan"),
          }).concat(kananCS),
        });
      }

      if (cEmel + cWA) {
        const kananE = ada(cIsu)
          ? [{ t: "card", x: 8.55, y: 2.85, w: 4.2, h: 4.0, title: T("C4.isu.tajuk", "Isu paling kerap ditanya", "Tajuk kad") },
             { t: "rank", x: 8.83, y: 3.5, w: 3.64, items: L(cIsu).slice(0, 6).map((k) => [k, num(cIsu[k])]), color: C.OCEAN }]
          : [{ t: "card", x: 8.55, y: 2.85, w: 4.2, h: 4.0, title: T("C4.catatan.tajuk", "Catatan", "Tajuk kad") },
             { t: "para", x: 8.83, y: 3.5, w: 3.64, h: 3.1, text: T("C4.catatan", "Tulis catatan bulan ini di sini melalui helaian Teks dalam fail Excel.", "Catatan manual") }];
        slides.push({
          kind: "content", section: "Perkhidmatan Pelanggan",
          title: T("C4.tajuk", "E-mel & WhatsApp", "Tajuk slaid"), periode,
          blocks: [
            { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "E-mel dijawab", value: cEmel, color: C.INK },
            { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "WhatsApp & panggilan", value: cWA, color: C.INK },
            { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Jumlah keseluruhan dijawab", value: cEmel + cWA, fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: "Semua saluran digabungkan", subColor: C.ICE },
          ].concat(CC({
            id: "C4.trend", x: G.M, y: 2.85, w: 7.7, h: 4.0,
            title: "Trend bulanan", note: "E-mel berbanding WhatsApp & panggilan",
            body: { k: "line", labels: TR("CS_Emel").bulan, series: [
              { name: "E-mel", values: TR("CS_Emel").nilai, color: C.OCEAN },
              { name: "WhatsApp / Panggilan", values: TR("CS_WA").nilai, color: C.TEAL }] },
            find: fSplit(cEmel, cWA, "e-mel", "WhatsApp & panggilan"),
          })).concat(kananE),
        });
      }
    }

    /* ═══════════ 3. LAPANGAN ═══════════ */
    const lR = g("Lawatan", "Ringkasan"), lPBT = g("Lawatan", "PBT");
    const lKat = g("Lawatan", "Kategori"), lZon = g("Lawatan", "Zon");
    const lStat = g("Lawatan", "Status"), lAge = g("Lawatan", "Ageing_Tersangkut");
    const lKes = tanpaKosong(g("Lawatan", "Kesukaran")), lSum = g("Lawatan", "Sumber");
    const lJab = g("Lawatan", "Jabatan");
    const lTot = num(lR.Total), lBrg = num(lR.Bergerak), lTsk = num(lR.Tersangkut);

    const pR = g("Pencegahan", "Ringkasan"), pPBT = g("Pencegahan", "PBT");
    const pKat = g("Pencegahan", "Kategori"), pZon = g("Pencegahan", "Zon");
    const pAgn = g("Pencegahan", "Agensi");
    const pStat = g("Pencegahan", "Status"), pAge = g("Pencegahan", "Ageing_Aktif");
    const pTot = num(pR.Total), pAkt = num(pR.Aktif), pSel = num(pR.Selesai);

    if (lTot || pTot) {
      slides.push({
        kind: "section", num: 3, title: T("S3.tajuk", "Lapangan", "Tajuk seksyen"),
        items: ["QC aduan & tanda perlu lawatan", "Laksana lawatan tapak & rekod hasil",
          "Daftar aduan pencegahan"],
      });
    }

    if (lTot) {
      slides.push({
        kind: "content", section: "Lapangan \u00B7 Lawatan Tapak (QC)",
        title: T("L1.tajuk", "Statistik Keseluruhan Lawatan Tapak (QC)", "Tajuk slaid"), periode,
        blocks: [
          { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Lawatan Tapak (QC)", value: lTot, fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: lR.Delta == null ? "Bulan " + periode : dtxt(num(lR.Delta)), subColor: C.ICE },
          { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Purata sehari bekerja", value: Math.round(lTot / 21), color: C.INK, sub: "Berasaskan 21 hari bekerja" },
          { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "PBT tertinggi", value: L(lPBT)[0] || "-", color: C.TEAL, vsize: 34, sub: fmt(num(V(lPBT)[0])) + " lawatan dilaksana" },
        ].concat(TR("Lawatan").nilai.length >= 2 ? CC({
          id: "L1.trend", x: G.M, y: 2.85, w: 7.7, h: 4.0,
          title: "Trend bulanan Lawatan Tapak (QC)",
          body: { k: "col", labels: TR("Lawatan").bulan, values: TR("Lawatan").nilai, color: C.TEAL },
          find: fTrend(TR("Lawatan").nilai, TR("Lawatan").bulan, "lawatan"),
        }) : CC({
          id: "L1.sumber", x: G.M, y: 2.85, w: 7.7, h: 4.0,
          title: "Sumber aduan dilawat",
          note: "Hanya satu bulan data tersedia, carta trend dilangkau",
          body: { k: "bar", labels: L(lSum), values: V(lSum).map(num), color: C.OCEAN },
          find: fTop(lSum, "lawatan"),
        })).concat(L(lPBT).length > 1 ? CC({
          id: "L1.pbt", x: 8.55, y: 2.85, w: 4.2, h: 4.0, title: "Pecahan ikut PBT",
          body: { k: "donut", labels: L(lPBT), values: V(lPBT).map(num), colors: SERIES },
          find: fTop(lPBT, "lawatan"),
        }) : CC({
          id: "L1.jabatan", x: 8.55, y: 2.85, w: 4.2, h: 4.0,
          title: "Pecahan ikut jabatan", note: L(lPBT)[0] + " sahaja dalam data ini",
          body: { k: "bar", labels: L(lJab), values: V(lJab).map(num), color: C.SLATE },
          find: fTop(lJab, "lawatan"),
        })),
      });

      slides.push({
        kind: "content", section: "Lapangan \u00B7 Lawatan Tapak (QC)",
        title: T("L2.tajuk", "Pecahan Aduan Dilawat", "Tajuk slaid"), periode,
        blocks: CC({
          id: "L2.kategori", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
          title: "Kategori aduan dilawat", note: "Sepuluh teratas",
          body: { k: "bar", labels: L(lKat), values: V(lKat).map(num), color: C.OCEAN },
          find: fTop(lKat, "lawatan"),
        }).concat(ada(lZon) ? CC({
          id: "L2.zon", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
          title: "Pecahan ikut zon", note: "Menunjukkan kawasan tumpuan lawatan",
          body: { k: "bar", labels: L(lZon), values: V(lZon).map(num), color: C.TEAL },
          find: fTop(lZon, "lawatan"),
        }) : CC({
          id: "L2.kesukaran", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
          title: "Tahap kesukaran aduan dilawat", note: "Aduan belum selesai - termasuk Selesai Sementara, tidak termasuk Diselesaikan dan Ditolak",
          body: { k: "donut", labels: L(lKes), values: V(lKes).map(num), colors: L(lKes).map((x) => KESUKARAN_WARNA[x] || C.MUTED) },
          find: fTop(lKes, "aduan"),
        })),
      });

      if (ada(lStat)) {
        const bawahL = ada(lAge)
          ? CC({ id: "L3.ageing", x: G.M, y: 3.0, w: G.COL, h: 3.25,
              title: "Usia aduan masih tersangkut", note: "Tempoh tergantung selepas lawatan",
              body: { k: "segment", segs: ageSegs(lAge) }, find: fAge(lAge) })
          : CC({ id: "L3.kesukaran", x: G.M, y: 3.0, w: G.COL, h: 3.85,
              title: "Tahap kesukaran aduan dilawat", note: "Aduan belum selesai - termasuk Selesai Sementara, tidak termasuk Diselesaikan dan Ditolak",
              body: { k: "donut", labels: L(lKes), values: V(lKes).map(num), colors: L(lKes).map((x) => KESUKARAN_WARNA[x] || C.MUTED) },
              find: fTop(lKes, "aduan") });
        slides.push({
          kind: "content", section: "Lapangan \u00B7 Lawatan Tapak (QC)",
          title: T("L3.tajuk", "Status Aduan Selepas Lawatan", "Tajuk slaid"), periode,
          blocks: [
            { t: "stat", x: G.M, y: G.TOP, w: HW, h: 1.5, label: "Telah bergerak", value: lBrg, color: C.MOSS, sub: pct(lBrg, lTot) + "% daripada aduan dilawat" },
            { t: "stat", x: G.M + HW + 0.25, y: G.TOP, w: HW, h: 1.5, label: "Masih tersangkut", value: lTsk, color: C.CRIMSON, sub: pct(lTsk, lTot) + "% daripada aduan dilawat" },
          ].concat(bawahL).concat(CC({
            id: "L3.status", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
            title: "Status semasa aduan dilawat", note: fmt(lTot) + " aduan telah dilawat pada " + periode,
            body: { k: "donut", labels: L(lStat), values: V(lStat).map(num), colors: warnaStatus(L(lStat)) },
            find: fStatus(lStat, ["Diselesaikan", "Selesai Sementara"]),
          })),
        });
      }
    }

    if (pTot) {
      const statCegah = [
        { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Aduan pencegahan didaftar", value: pTot, fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: pR.Delta == null ? "Bulan " + periode : dtxt(num(pR.Delta)), subColor: C.ICE },
      ];
      if (pAkt || pSel) {
        statCegah.push({ t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Dalam tindakan", value: pAkt, color: C.AMBER, sub: pct(pAkt, pTot) + "% daripada jumlah didaftar" });
        statCegah.push({ t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Telah selesai", value: pSel, color: C.MOSS, sub: pct(pSel, pTot) + "% daripada jumlah didaftar" });
      } else {
        statCegah.push({ t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Purata sehari bekerja", value: Math.round(pTot / 21), color: C.INK, sub: "Berasaskan 21 hari bekerja" });
        statCegah.push({ t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "PBT tertinggi", value: L(pPBT)[0] || "-", color: C.TEAL, vsize: 34, sub: fmt(num(V(pPBT)[0])) + " aduan pencegahan" });
      }

      slides.push({
        kind: "content", section: "Lapangan \u00B7 Aduan Pencegahan",
        title: T("L4.tajuk", "Statistik Keseluruhan Aduan Pencegahan", "Tajuk slaid"), periode,
        blocks: statCegah.concat(TR("Pencegahan").nilai.length >= 2 ? CC({
          id: "L4.trend", x: G.M, y: 2.85, w: 7.7, h: 4.0,
          title: "Trend bulanan aduan pencegahan",
          note: "Isu dikesan sendiri oleh pegawai semasa turun padang",
          body: { k: "col", labels: TR("Pencegahan").bulan, values: TR("Pencegahan").nilai, color: C.MOSS },
          find: fTrend(TR("Pencegahan").nilai, TR("Pencegahan").bulan, "aduan pencegahan"),
        }) : CC({
          id: "L4.kategori2", x: G.M, y: 2.85, w: 7.7, h: 4.0,
          title: "Kategori isu dikesan",
          note: "Hanya satu bulan data tersedia, carta trend dilangkau",
          body: { k: "bar", labels: L(pKat), values: V(pKat).map(num), color: C.MOSS },
          find: fTop(pKat, "isu"),
        })).concat(L(pPBT).length > 1 ? CC({
          id: "L4.pbt", x: 8.55, y: 2.85, w: 4.2, h: 4.0, title: "Pecahan ikut PBT",
          body: { k: "donut", labels: L(pPBT), values: V(pPBT).map(num), colors: SERIES },
          find: fTop(pPBT, "aduan"),
        }) : CC({
          id: "L4.agensi", x: 8.55, y: 2.85, w: 4.2, h: 4.0, title: "Agensi bertanggungjawab",
          body: { k: "bar", labels: L(pAgn), values: V(pAgn).map(num), color: C.SLATE },
          find: fTop(pAgn, "isu"),
        })),
      });

      const keduaP = ada(pZon) ? { set: pZon, tajuk: "Pecahan ikut zon", nota: "Zon dengan isu berulang wajar dijadualkan rondaan tetap", jenis: "col" }
        : ada(pAgn) ? { set: pAgn, tajuk: "Agensi bertanggungjawab", nota: "Pihak yang menerima aduan pencegahan", jenis: "bar" } : null;
      slides.push({
        kind: "content", section: "Lapangan \u00B7 Aduan Pencegahan",
        title: T("L5.tajuk", "Pecahan Isu Dikesan", "Tajuk slaid"), periode,
        blocks: CC({
          id: "L5.kategori", x: G.M, y: G.TOP, w: keduaP ? G.COL : G.CW, h: 5.60,
          title: "Kategori isu dikesan", note: "Susunan tertinggi ke terendah",
          body: { k: "bar", labels: L(pKat), values: V(pKat).map(num), color: C.MOSS },
          find: fTop(pKat, "isu"),
        }).concat(keduaP ? CC({
          id: "L5.kedua", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
          title: keduaP.tajuk, note: keduaP.nota,
          body: { k: keduaP.jenis, labels: L(keduaP.set), values: V(keduaP.set).map(num), color: C.TEAL },
          find: fTop(keduaP.set, "isu"),
        }) : []),
      });

      if (ada(pStat)) {
        slides.push({
          kind: "content", section: "Lapangan \u00B7 Aduan Pencegahan",
          title: T("L6.tajuk", "Status Semasa", "Tajuk slaid"), periode,
          blocks: CC({
            id: "L6.status", x: G.M, y: G.TOP, w: G.COL, h: 5.60,
            title: "Pecahan status", note: fmt(pTot) + " aduan pencegahan didaftarkan",
            body: { k: "donut", labels: L(pStat), values: V(pStat).map(num), colors: [C.AMBER, C.MOSS] },
            find: fSplit(pSel, pAkt, "telah selesai", "masih dalam tindakan"),
          }).concat([
            { t: "stat", x: G.COL_X2, y: G.TOP, w: HW, h: 1.5, label: "Dalam tindakan", value: pAkt, color: C.AMBER, sub: "Belum selesai sepenuhnya" },
            { t: "stat", x: G.COL_X2 + HW + 0.25, y: G.TOP, w: HW, h: 1.5, label: "Telah selesai", value: pSel, color: C.MOSS, sub: "Kerja pembaikan selesai" },
          ]).concat(ada(pAge) ? CC({
            id: "L6.ageing", x: G.COL_X2, y: 3.0, w: G.COL, h: 3.85,
            title: "Usia aduan dalam tindakan",
            body: { k: "segment", segs: ageSegs(pAge) }, find: fAge(pAge),
          }) : []),
        });
      }
    }

    /* ═══════════ 4. STATUS SEMASA ADUAN (IKUT PBT) ═══════════ */
    const PBTS = (D.pbt && D.pbt.senarai) || [];
    if (PBTS.length) {
      const tab = (nama) => ((D.pbt && D.pbt[nama]) || []);
      const kadarWarna = (v) => (v < 50 ? C.CRIMSON : v < 70 ? C.AMBER : C.MOSS);
      const jumA = (a) => a.reduce((x, y) => x + y, 0);

      slides.push({
        kind: "section", num: 4,
        title: T("S4.tajuk", "Status Semasa Aduan", "Tajuk seksyen"),
        items: ["Perbandingan empat PBT", "Statistik & status setiap PBT",
          "Kategori, sumber & prestasi jabatan",
          "Kualiti penyelesaian & aduan tertunggak"],
      });

      const R = (k) => g("PBT", k + "_Ringkasan");
      const jumSemua = jumA(PBTS.map((p) => num(R(p.kod).Jumlah)));
      const selSemua = jumA(PBTS.map((p) => num(R(p.kod).Selesai)));
      const slaSemua = jumA(PBTS.map((p) => num(R(p.kod).Lepas_SLA)));

      const cmpRows = PBTS.map((p) => {
        const r = R(p.kod), jml = num(r.Jumlah), sel = num(r.Selesai), sla = num(r.Lepas_SLA);
        const kadar = pct(sel, jml);
        return [p.kod, String(jml), String(sel),
          { text: kadar + "%", color: kadarWarna(kadar), bold: true },
          { text: String(sla), color: C.CRIMSON, bold: true },
          { text: pct(sla, jml) + "%", color: kadarWarna(100 - pct(sla, jml)) }];
      });
      const terbaik = PBTS.slice().sort((a, b) =>
        pct(num(R(b.kod).Selesai), num(R(b.kod).Jumlah)) - pct(num(R(a.kod).Selesai), num(R(a.kod).Jumlah)))[0];
      const terbanyakSLA = PBTS.slice().sort((a, b) => num(R(b.kod).Lepas_SLA) - num(R(a.kod).Lepas_SLA))[0];

      if (PBTS.length > 1) slides.push({
        kind: "content", section: "Status Semasa Aduan",
        title: T("P0.tajuk", "Perbandingan " + PBTS.length + " PBT", "Tajuk slaid"), periode,
        blocks: [
          { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Jumlah aduan seluruh Johor", value: jumSemua, fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: PBTS.length + " PBT digabungkan", subColor: C.ICE },
          { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Telah selesai", value: selSemua, color: C.MOSS, sub: pct(selSemua, jumSemua) + "% daripada jumlah aduan" },
          { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Melepasi SLA", value: slaSemua, color: C.CRIMSON, sub: pct(slaSemua, jumSemua) + "% daripada jumlah aduan" },
          { t: "table", x: G.M, y: 2.95, w: G.CW,
            head: ["PBT", "Jumlah Aduan", "Telah Selesai", "Kadar Selesai", "Melepasi SLA", "% Melepasi SLA"],
            rows: cmpRows, colW: [2.9, 2.0, 2.0, 1.9, 1.8, 1.6], rowH: 0.6, fs: 12.5, leftCols: [0] },
          { t: "finding", x: G.M + 0.02, y: 6.2, w: G.CW, h: 0.5,
            text: T("P0.dapatan", "Kadar selesai tertinggi: " + (terbaik ? terbaik.kod : "-") +
              ". Bilangan aduan melepasi SLA tertinggi: " + (terbanyakSLA ? terbanyakSLA.kod : "-") + ".", "Ayat dapatan") },
        ],
      });

      PBTS.forEach((p) => {
        const k = p.kod, EY = "Status Semasa Aduan \u00B7 " + k;
        const r = R(k), jml = num(r.Jumlah), sel = num(r.Selesai), sla = num(r.Lepas_SLA);
        const stat = g("PBT", k + "_Status");
        const ditolak = num(stat["Ditolak"]);
        const aktif = jml - sel - ditolak;
        const kat = g("PBT", k + "_Kategori"), sum = g("PBT", k + "_Sumber");
        const kes = tanpaKosong(g("PBT", k + "_Kesukaran")), nil = g("PBT", k + "_Penilaian");

        /* 1 - Statistik Keseluruhan */
        slides.push({
          kind: "content", section: EY,
          title: T(k + "1.tajuk", "Statistik Keseluruhan Aduan " + k, "Tajuk slaid"), periode,
          blocks: [
            { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Jumlah aduan", value: jml, fill: C.INK, color: C.WHITE, labelColor: C.ICE,
            sub: r.Delta == null ? p.nama : dtxt(num(r.Delta)), subColor: C.ICE },
            { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Telah selesai", value: sel, color: C.MOSS, sub: pct(sel, jml) + "% daripada jumlah aduan" },
            { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Melepasi SLA", value: sla, fill: C.PALE_RED, color: C.CRIMSON, labelColor: C.DARK_RED, sub: pct(sla, jml) + "% daripada jumlah aduan", subColor: C.DARK_RED },
          ].concat(TR(k + "_Aduan").nilai.length >= 2 ? CC({
            id: k + "1.trend", x: G.M, y: 2.85, w: G.CW, h: 4.0,
            title: "Trend bulanan aduan diterima", note: p.nama,
            body: { k: "col", labels: TR(k + "_Aduan").bulan, values: TR(k + "_Aduan").nilai, color: C.TEAL },
            find: fTrend(TR(k + "_Aduan").nilai, TR(k + "_Aduan").bulan, "aduan"),
          }) : CC({
            id: k + "1.sumber", x: G.M, y: 2.85, w: G.CW, h: 4.0,
            title: "Sumber aduan", note: p.nama + " \u00B7 hanya satu bulan data tersedia, carta trend dilangkau",
            body: { k: "bar", labels: L(g("PBT", k + "_Sumber")), values: V(g("PBT", k + "_Sumber")).map(num), color: C.OCEAN },
            find: fTop(g("PBT", k + "_Sumber"), "aduan"),
          })),
        });

        /* 2 - Pecahan Status */
        slides.push({
          kind: "content", section: EY,
          title: T(k + "2.tajuk", "Pecahan Status Aduan", "Tajuk slaid"), periode,
          blocks: CC({
            id: k + "2.status", x: G.M, y: G.TOP, w: 7.7, h: 5.60,
            title: "Status semasa setiap aduan",
            note: "Menunjukkan di peringkat mana " + fmt(jml) + " aduan berada",
            body: { k: "donut", labels: L(stat), values: V(stat).map(num),
              colors: warnaStatus(L(stat)) },
            find: fTop(stat, "aduan"),
          }).concat([
            { t: "stat", x: 8.55, y: G.TOP, w: 4.2, h: 1.6, label: "Masih aktif", value: aktif, color: C.AMBER, sub: pct(aktif, jml) + "% masih dalam proses" },
            { t: "stat", x: 8.55, y: 3.05, w: 4.2, h: 1.6, label: "Telah selesai", value: sel, color: C.MOSS, sub: "Diselesaikan & selesai sementara" },
            { t: "stat", x: 8.55, y: 5.10, w: 4.2, h: 1.6, label: "Ditolak", value: ditolak, color: C.MUTED, sub: pct(ditolak, jml) + "% daripada jumlah aduan" },
          ]),
        });

        /* 3 - Kategori & Sumber */
        slides.push({
          kind: "content", section: EY,
          title: T(k + "3.tajuk", "Kategori & Sumber Aduan", "Tajuk slaid"), periode,
          blocks: CC({
            id: k + "3.kategori", x: G.M, y: G.TOP,
            w: TR(k + "_Aduan").nilai.length < 2 ? G.CW : G.COL, h: 5.60,
            title: "Kategori aduan",
            note: "Tidak termasuk " + fmt(num(r.Lain_Lain)) + " aduan berkategori Lain-lain",
            body: { k: "bar", labels: L(kat), values: V(kat).map(num), color: C.OCEAN },
            find: fTop(kat, "aduan"),
          }).concat(TR(k + "_Aduan").nilai.length < 2 ? [] : CC({
            id: k + "3.sumber", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
            title: "Sumber aduan", note: "Saluran pengadu membuat laporan",
            body: { k: "donut", labels: L(sum), values: V(sum).map(num), colors: SERIES },
            find: fTop(sum, "aduan"),
          })),
        });

        /* 4 - Prestasi Jabatan */
        const jb = tab("jabatan").filter((x) => x.pbt === k && num(x.jumlah) >= 3)
          .map((x) => {
            const t = num(x.jumlah), sl = num(x.selesai);
            return { nama: x.nama, jumlah: t, selesai: sl, proses: num(x.proses),
              purata: num(x.purata), kadar: t ? (sl / t) * 100 : 0 };
          }).sort((a, b) => b.jumlah - a.jumlah);
        const besar = jb.filter((x) => x.jumlah >= 10);
        const jBaik = besar.slice().sort((a, b) => b.kadar - a.kadar)[0];
        const jLemah = besar.slice().sort((a, b) => a.kadar - b.kadar)[0];

        slides.push({
          kind: "content", section: EY,
          title: T(k + "4.tajuk", "Prestasi Jabatan", "Tajuk slaid"), periode,
          blocks: [
            { t: "table", x: G.M, y: 1.35, w: G.CW,
              head: ["Jabatan", "Jumlah", "Selesai", "Dalam Proses", "Purata (hari)", "Kadar Selesai"],
              rows: jb.map((x) => [x.nama, String(x.jumlah), String(x.selesai), String(x.proses),
                x.purata.toFixed(1),
                { text: x.kadar.toFixed(1) + "%", color: kadarWarna(x.kadar), bold: true }]),
              colW: [4.4, 1.5, 1.5, 1.8, 1.6, 1.4], rowH: 0.5, fs: 11.5, leftCols: [0] },
            { t: "finding", x: G.M + 0.02, y: 6.15, w: G.CW, h: 0.55,
              text: T(k + "4.dapatan", jBaik && jLemah
                ? "Kadar selesai tertinggi: " + jBaik.nama + " (" + jBaik.kadar.toFixed(1) + "%). Memerlukan perhatian: " +
                  jLemah.nama + " (" + jLemah.kadar.toFixed(1) + "% daripada " + jLemah.jumlah +
                  " aduan). Jabatan dengan kurang daripada 3 aduan tidak disenaraikan."
                : "Jabatan dengan kurang daripada 3 aduan tidak disenaraikan.", "Ayat dapatan") },
            { t: "finding", x: G.M + 0.02, y: 6.55, w: G.CW, h: 0.3,
              text: "Purata hari dikira dalam hari bekerja, dari tarikh aduan hingga tarikh kemaskini terakhir. Aduan yang belum ditugaskan kepada mana-mana jabatan tidak disenaraikan." },
          ],
        });

        /* 5 - Kualiti Penyelesaian */
        const totNilai = jumA(V(nil).map(num));
        const puas = num(nil["Sangat Baik"]) + num(nil["Memuaskan"]) +
          num(nil["Cemerlang"]) + num(nil["Baik"]);
        slides.push({
          kind: "content", section: EY,
          title: T(k + "5.tajuk", "Kualiti Penyelesaian", "Tajuk slaid"), periode,
          blocks: CC({
            id: k + "5.kesukaran", x: G.M, y: G.TOP, w: ada(nil) ? G.COL : G.CW, h: 5.60,
            title: "Tahap kesukaran aduan", note: "Aduan belum selesai - termasuk Selesai Sementara, tidak termasuk Diselesaikan dan Ditolak",
            body: { k: "donut", labels: L(kes), values: V(kes).map(num), colors: L(kes).map((x) => KESUKARAN_WARNA[x] || C.MUTED) },
            find: fTop(kes, "aduan"),
          }).concat(ada(nil) ? CC({
            id: k + "5.penilaian", x: G.COL_X2, y: G.TOP, w: G.COL, h: 5.60,
            title: "Penilaian pengadu",
            note: fmt(totNilai) + " maklum balas diterima daripada " + fmt(jml) + " aduan",
            body: { k: "donut", labels: L(nil), values: V(nil).map(num), colors: L(nil).map((x) => PENILAIAN_WARNA[x] || C.MUTED) },
            find: pct(puas, totNilai) + "% pengadu berpuas hati dengan penyelesaian yang diberikan.",
          }) : []),
        });

        /* 6 - Kawasan & Aduan Tertunggak */
        const zn = tab("zon").filter((x) => x.pbt === k)
          .sort((a, b) => num(b.jumlah) - num(a.jumlah)).slice(0, 5);
        const tg = tab("tertunggak").filter((x) => x.pbt === k)
          .sort((a, b) => num(b.hari) - num(a.hari)).slice(0, 5);

        slides.push({
          kind: "content", section: EY,
          title: T(k + "6.tajuk", "Kawasan & Aduan Tertunggak", "Tajuk slaid"), periode,
          blocks: [
            { t: "label", x: G.M, y: 1.22, w: 8, text: T(k + "6.label1", "Lima zon ahli majlis dengan aduan tertinggi", "Label jadual") },
            { t: "table", x: G.M, y: 1.55, w: G.CW,
              head: ["Zon / Kawasan", "Jumlah", "Selesai", "Dalam Proses", "Kadar Selesai"],
              rows: zn.map((x) => {
                const t = num(x.jumlah), sl = num(x.selesai), kd = t ? (sl / t) * 100 : 0;
                return [x.zon, String(t), String(sl), String(num(x.proses)),
                  { text: kd.toFixed(1) + "%", color: kadarWarna(kd), bold: true }];
              }),
              colW: [4.4, 1.9, 1.9, 2.2, 1.8], rowH: 0.40, fs: 11, leftCols: [0] },
            { t: "label", x: G.M, y: 4.22, w: 8, text: T(k + "6.label2", "Lima aduan tertunggak terlama", "Label jadual") },
            { t: "table", x: G.M, y: 4.55, w: G.CW,
              head: ["No. Rujukan", "Tajuk Aduan", "Status", "Tahap", "Tertunggak"],
              rows: tg.map((x) => [x.ruj,
                String(x.tajuk).length > 44 ? String(x.tajuk).slice(0, 42) + "..." : x.tajuk,
                x.status, x.tahap,
                { text: num(x.hari) + " hari", color: C.CRIMSON, bold: true }]),
              colW: [2.4, 4.9, 1.9, 1.5, 1.5], rowH: 0.38, fs: 10, leftCols: [0, 1] },
          ],
        });
      });
    }


    /* ═══════════ 5. AHLI MAJLIS (MBJB SAHAJA) ═══════════ */
    /* Aduan Kompleks bertempoh SLA 365 hari bekerja - mustahil selesai
       dalam tempoh mingguan. Ia dikeluarkan daripada kiraan kadar selesai
       supaya angka mencerminkan kerja yang benar-benar boleh disiapkan. */
    const AHLI = (D.ahli || []).map((a) => {
      const jml = num(a.jumlah), sel = num(a.selesai), kom = num(a.kompleks);
      const tlk = num(a.ditolak);
      /* Kadar = Selesai / (Jumlah - Ditolak). Aduan ditolak tidak mungkin
         "selesai", jadi ia dikeluarkan daripada penyebut. Aduan Kompleks
         DIKEKALKAN kerana ia boleh dan memang diselesaikan - mengeluarkannya
         daripada penyebut sahaja menghasilkan kadar melebihi 100%.
         Lajur Kompleks kekal sebagai konteks. */
      const bolehSelesai = Math.max(0, jml - tlk);
      return { nama: a.nama || "", zon: a.zon, jumlah: jml, selesai: sel, kompleks: kom,
        proses: num(a.proses), tertunggak: num(a.tertunggak),
        baru: num(a.baru), ditolak: tlk, bolehSelesai: bolehSelesai,
        kadar: bolehSelesai ? Math.min(100, (sel / bolehSelesai) * 100) : 0 };
    });

    if (AHLI.length) {
      const A_PBT = D.info.ahli_pbt || "MBJB";
      const A_MGG = D.info.ahli_minggu || periode;
      const A_JLT = D.info.ahli_julat || "";
      const kadarWarna2 = (v) => (v < 40 ? C.CRIMSON : v < 65 ? C.AMBER : C.MOSS);
      const jumA2 = (a) => a.reduce((x, y) => x + y, 0);
      const PER_SLAID = 12;

      slides.push({
        kind: "section", num: 5,
        title: T("S5.tajuk", "Prestasi Zon Ahli Majlis", "Tajuk seksyen"),
        items: [A_PBT + " \u00B7 " + A_MGG, A_JLT || "",
          AHLI.length + " zon ahli majlis", "Disusun mengikut jumlah aduan"].filter(Boolean),
      });

      const EY5 = "Ahli Majlis \u00B7 " + A_PBT;
      const aJum = jumA2(AHLI.map((a) => a.jumlah));
      const aSel = jumA2(AHLI.map((a) => a.selesai));
      const aPrs = jumA2(AHLI.map((a) => a.proses));
      const aTtg = jumA2(AHLI.map((a) => a.tertunggak));
      const aBar = jumA2(AHLI.map((a) => a.baru));
      const aKom = jumA2(AHLI.map((a) => a.kompleks));
      const aTlk = jumA2(AHLI.map((a) => a.ditolak));
      const adaNama = AHLI.some((a) => a.nama);
      const aBoleh = Math.max(0, aJum - aTlk);
      /* Disusun mengikut jumlah aduan - zon paling banyak aduan di atas. */
      const susun = AHLI.slice().sort((a, b) =>
        (b.jumlah - a.jumlah) || (b.tertunggak - a.tertunggak) || (a.kadar - b.kadar));
      const nBaik = susun.filter((a) => a.kadar >= 65).length;
      const nSed = susun.filter((a) => a.kadar >= 40 && a.kadar < 65).length;
      const nLemah = susun.filter((a) => a.kadar < 40).length;
      const aKat = g("Ahli", "Kategori"), aSum = g("Ahli", "Sumber");
      const aKes = tanpaKosong(g("Ahli", "Kesukaran")), aStat = g("Ahli", "Status");

      /* 1 - Statistik Keseluruhan */
      slides.push({
        kind: "content", section: EY5,
        title: T("A1.tajuk", "Statistik Keseluruhan Aduan Ahli Majlis " + A_PBT, "Tajuk slaid"), periode: A_MGG,
        blocks: [
          { t: "stat", x: G.M, y: R1, w: SW, h: 1.35, label: "Jumlah aduan diterima", value: aJum, fill: C.INK, color: C.WHITE, labelColor: C.ICE, sub: "Bersumber Ahli Majlis, merentas " + AHLI.length + " zon", subColor: C.ICE },
          { t: "stat", x: G.M + SW + SG, y: R1, w: SW, h: 1.35, label: "Telah selesai", value: aSel, color: C.MOSS, sub: pct(aSel, aBoleh) + "% daripada " + aBoleh + " aduan (tolak aduan ditolak)" },
          { t: "stat", x: G.M + (SW + SG) * 2, y: R1, w: SW, h: 1.35, label: "Tertunggak", value: aTtg, fill: C.PALE_RED, color: C.CRIMSON, labelColor: C.DARK_RED, sub: pct(aTtg, aJum) + "% melepasi tarikh akhir", subColor: C.DARK_RED },
          { t: "card", x: G.M, y: 2.85, w: G.CW, h: 1.35, title: T("A1.status.tajuk", "Status keseluruhan aduan", "Tajuk kad") },
          { t: "segment", x: G.M + 0.3, y: 3.32, w: G.CW - 0.6, barH: 0.5, segs: [
            { label: "Selesai (" + aSel + ")", value: aSel, color: C.MOSS },
            { label: "Dalam proses (" + aPrs + ")", value: aPrs, color: C.OCEAN },
            { label: "Baru diterima (" + aBar + ")", value: aBar, color: C.AMBER },
            { label: "Tertunggak (" + aTtg + ")", value: aTtg, color: C.CRIMSON },
          ].concat(aTlk ? [{ label: "Ditolak (" + aTlk + ")", value: aTlk, color: C.MUTED }] : []) },
        ].concat(CC({
          id: "A1.taburan", x: G.M, y: 4.35, w: G.COL, h: 2.5,
          title: "Taburan prestasi zon",
          body: { k: "col", labels: ["Kadar 65% ke atas", "Kadar 40-64%", "Kadar bawah 40%"],
            values: [nBaik, nSed, nLemah], color: C.SLATE },
          findH: 0.62,
          find: nBaik + " zon mencapai 65% ke atas; " + nLemah + " zon di bawah 40%. " +
            "Kadar dikira selepas mengeluarkan " + aTlk + " aduan ditolak daripada penyebut.",
        })).concat(ada(aKes) ? CC({
          id: "A1.kesukaran", x: G.COL_X2, y: 4.35, w: G.COL, h: 2.5,
          title: "Tahap kesukaran",
          body: { k: "donut", labels: L(aKes), values: V(aKes).map(num),
            colors: L(aKes).map((x) => KESUKARAN_WARNA[x] || C.MUTED) },
          find: fTop(aKes, "aduan"),
        }) : (ada(aSum) ? CC({
          id: "A1.sumber", x: G.COL_X2, y: 4.35, w: G.COL, h: 2.5,
          title: "Sumber aduan",
          body: { k: "bar", labels: L(aSum), values: V(aSum).map(num), color: C.OCEAN },
          find: fTop(aSum, "aduan"),
        }) : [])),
      });

      /* 2 - Kategori & Status (dilangkau jika data tiada) */
      if (ada(aKat) || ada(aStat)) {
        const duaLajur = ada(aKat) && ada(aStat);
        slides.push({
          kind: "content", section: EY5,
          title: T("A2.tajuk", "Kategori & Status Aduan", "Tajuk slaid"), periode: A_MGG,
          blocks: (ada(aKat) ? CC({
            id: "A2.kategori", x: G.M, y: G.TOP, w: duaLajur ? G.COL : G.CW, h: 5.60,
            title: "Kategori aduan", note: "Merentas semua zon ahli majlis",
            body: { k: "bar", labels: L(aKat), values: V(aKat).map(num), color: C.OCEAN },
            find: fTop(aKat, "aduan"),
          }) : []).concat(ada(aStat) ? CC({
            id: "A2.status", x: duaLajur ? G.COL_X2 : G.M, y: G.TOP,
            w: duaLajur ? G.COL : G.CW, h: 5.60,
            title: "Status semasa aduan", note: "Kedudukan aduan pada " + A_MGG,
            body: { k: "donut", labels: L(aStat), values: V(aStat).map(num),
              colors: warnaStatus(L(aStat)) },
            find: fStatus(aStat, ["Diselesaikan", "Selesai Sementara"]),
          }) : []),
        });
      }

      /* 3+ - Jadual prestasi zon, 12 baris setiap slaid */
      const bhg = Math.ceil(susun.length / PER_SLAID);
      for (let i = 0; i < bhg; i++) {
        const sub = susun.slice(i * PER_SLAID, (i + 1) * PER_SLAID);
        const mula = i * PER_SLAID + 1;
        slides.push({
          kind: "content", section: EY5,
          title: T("A3." + i + ".tajuk", "Prestasi Zon Ahli Majlis" +
            (bhg > 1 ? " (" + (i + 1) + " daripada " + bhg + ")" : ""), "Tajuk slaid"),
          periode: A_MGG,
          blocks: [
            { t: "label", x: G.M, y: 1.22, w: 9,
              text: "Kedudukan " + mula + " hingga " + (mula + sub.length - 1) +
                " daripada " + susun.length + " zon, disusun mengikut jumlah aduan diterima" },
            { t: "table", x: G.M, y: 1.55, w: G.CW,
              head: ["#"].concat(adaNama ? ["Ahli Majlis"] : [])
                .concat(["Zon / Kawasan", "Jumlah", "Kompleks", "Selesai", "Proses", "Tertunggak", "Kadar"]),
              rows: sub.map((a, j) => [String(mula + j)]
                .concat(adaNama ? [a.nama] : [])
                .concat([a.zon, String(a.jumlah),
                    { text: String(a.kompleks), color: a.kompleks ? C.SLATE : C.MUTED },
                  String(a.selesai), String(a.proses),
                  { text: String(a.tertunggak), color: a.tertunggak ? C.CRIMSON : C.MUTED, bold: !!a.tertunggak },
                  a.jumlah
                    ? { text: a.kadar.toFixed(0) + "%", color: kadarWarna2(a.kadar), bold: true }
                    : { text: "\u2014", color: C.MUTED }])),
              colW: adaNama
                ? [0.55, 3.2, 2.7, 0.95, 1.05, 0.95, 0.95, 1.05, 0.8]
                : [0.6, 4.2, 1.2, 1.25, 1.15, 1.15, 1.4, 1.25],
              rowH: 0.355, fs: 10, leftCols: adaNama ? [1, 2] : [1] },
            { t: "finding", x: G.M + 0.02, y: 6.32, w: G.CW, h: 0.45,
              text: T("A3." + i + ".dapatan", i === 0
                ? "Kadar selesai = Selesai / (Jumlah - Ditolak). Lajur Kompleks ditunjukkan sebagai konteks kerana aduan Kompleks bertempoh SLA 365 hari bekerja."
                : "Zon di bahagian bawah senarai menerima aduan paling sedikit melalui ahli majlis.", "Ayat dapatan") },
          ],
        });
      }
    }

    window.TEXT_KEYS = REG;
    return slides;
  };
})();
