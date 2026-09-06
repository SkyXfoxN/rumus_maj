/* excel.js — Jana templat Excel dan hurai fail yang dimuat naik.
   Skema 4 helaian, format panjang supaya senang diedit. */

(function () {
  const SHEETS = {
    INFO: "Info", DATA: "Data", TREND: "Trend", TEKS: "Teks",
    PBT: "PBT_Senarai", JAB: "PBT_Jabatan", ZON: "PBT_Zon", TGK: "PBT_Tertunggak",
    AHLI: "Ahli_Majlis",
  };

  const INFO_KEYS = [
    ["tajuk", "Tajuk Laporan"], ["sistem", "Nama Sistem"], ["periode", "Periode"],
    ["disediakan", "Disediakan Oleh"], ["tarikh_jana", "Tarikh Jana"], ["sumber", "Sumber Data"],
    ["ahli_pbt", "Ahli Majlis - PBT"], ["ahli_minggu", "Ahli Majlis - Minggu"],
    ["ahli_julat", "Ahli Majlis - Julat Tarikh"],
  ];

  /* ── Jana templat daripada data semasa ─────────────────── */
  window.buildTemplate = function (D, filename) {
    const wb = XLSX.utils.book_new();

    const info = [["Kunci", "Nilai"]];
    INFO_KEYS.forEach(([k, label]) => info.push([label, D.info[k] || ""]));
    const wsI = XLSX.utils.aoa_to_sheet(info);
    wsI["!cols"] = [{ wch: 22 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(wb, wsI, SHEETS.INFO);

    const data = [["Seksyen", "Set", "Label", "Nilai"]];
    Object.keys(D.d).forEach((sek) => {
      Object.keys(D.d[sek]).forEach((set) => {
        Object.keys(D.d[sek][set]).forEach((lbl) => {
          data.push([sek, set, lbl, D.d[sek][set][lbl]]);
        });
      });
    });
    const wsD = XLSX.utils.aoa_to_sheet(data);
    wsD["!cols"] = [{ wch: 18 }, { wch: 22 }, { wch: 32 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsD, SHEETS.DATA);

    const trend = [["Bulan", "Siri", "Nilai"]];
    Object.keys(D.trend.siri).forEach((siri) => {
      const s = D.trend.siri[siri] || {};
      (s.bulan || []).forEach((b, i) => trend.push([b, siri, (s.nilai || [])[i]]));
    });
    const wsT = XLSX.utils.aoa_to_sheet(trend);
    wsT["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsT, SHEETS.TREND);

    // Empat helaian PBT: senarai PBT dan tiga jadual
    const P = D.pbt || { senarai: [], jabatan: [], zon: [], tertunggak: [] };

    const shPBT = [["Kod", "Nama Penuh"]];
    P.senarai.forEach((r) => shPBT.push([r.kod, r.nama]));
    const wsP = XLSX.utils.aoa_to_sheet(shPBT);
    wsP["!cols"] = [{ wch: 10 }, { wch: 42 }];
    XLSX.utils.book_append_sheet(wb, wsP, SHEETS.PBT);

    const shJab = [["PBT", "Jabatan", "Jumlah", "Selesai", "Dalam Proses", "Purata (hari)"]];
    P.jabatan.forEach((r) => shJab.push([r.pbt, r.nama, r.jumlah, r.selesai, r.proses, r.purata]));
    const wsJ = XLSX.utils.aoa_to_sheet(shJab);
    wsJ["!cols"] = [{ wch: 8 }, { wch: 36 }, { wch: 9 }, { wch: 9 }, { wch: 13 }, { wch: 13 }];
    XLSX.utils.book_append_sheet(wb, wsJ, SHEETS.JAB);

    const shZon = [["PBT", "Zon / Kawasan", "Jumlah", "Selesai", "Dalam Proses"]];
    P.zon.forEach((r) => shZon.push([r.pbt, r.zon, r.jumlah, r.selesai, r.proses]));
    const wsZ = XLSX.utils.aoa_to_sheet(shZon);
    wsZ["!cols"] = [{ wch: 8 }, { wch: 30 }, { wch: 9 }, { wch: 9 }, { wch: 13 }];
    XLSX.utils.book_append_sheet(wb, wsZ, SHEETS.ZON);

    const shTgk = [["PBT", "No. Rujukan", "Tajuk Aduan", "Status", "Tahap", "Hari Tertunggak"]];
    P.tertunggak.forEach((r) => shTgk.push([r.pbt, r.ruj, r.tajuk, r.status, r.tahap, r.hari]));
    const wsT2 = XLSX.utils.aoa_to_sheet(shTgk);
    wsT2["!cols"] = [{ wch: 8 }, { wch: 18 }, { wch: 52 }, { wch: 18 }, { wch: 12 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, wsT2, SHEETS.TGK);

    const shAhli = [["Ahli Majlis", "Zon / Kawasan", "Jumlah", "Kompleks", "Selesai",
      "Dalam Proses", "Tertunggak", "Baru Diterima"]];
    (D.ahli || []).forEach((r) => shAhli.push([r.nama, r.zon, r.jumlah, r.kompleks,
      r.selesai, r.proses, r.tertunggak, r.baru]));
    const wsA = XLSX.utils.aoa_to_sheet(shAhli);
    wsA["!cols"] = [{ wch: 34 }, { wch: 28 }, { wch: 9 }, { wch: 10 }, { wch: 9 },
      { wch: 13 }, { wch: 12 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsA, SHEETS.AHLI);

    // Helaian Teks: senarai setiap tajuk & ayat yang boleh diubah.
    // buildSpec dipanggil dahulu supaya senarai ID lengkap.
    buildSpec(D);
    const seen = {};
    const teks = [["ID", "Perkara", "Teks Automatik (rujukan sahaja)", "Teks Ubahsuai"]];
    (window.TEXT_KEYS || []).forEach((t) => {
      if (seen[t.key]) return;
      seen[t.key] = 1;
      teks.push([t.key, t.desc, t.auto, (D.teks && D.teks[t.key]) || ""]);
    });
    const wsX = XLSX.utils.aoa_to_sheet(teks);
    wsX["!cols"] = [{ wch: 20 }, { wch: 16 }, { wch: 70 }, { wch: 70 }];
    XLSX.utils.book_append_sheet(wb, wsX, SHEETS.TEKS);

    XLSX.writeFile(wb, filename);
  };

  /* ── Hurai fail yang dimuat naik ───────────────────────── */
  window.parseWorkbook = function (wb) {
    const warn = [];
    const need = [SHEETS.INFO, SHEETS.DATA, SHEETS.TREND];
    need.forEach((n) => { if (!wb.Sheets[n]) throw new Error("Helaian '" + n + "' tiada dalam fail Excel."); });

    const rows = (name) => (wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false }) : []);

    // Info
    const info = {};
    const labelToKey = {};
    INFO_KEYS.forEach(([k, label]) => { labelToKey[label.toLowerCase()] = k; });
    rows(SHEETS.INFO).slice(1).forEach((r) => {
      const k = labelToKey[String(r[0] || "").trim().toLowerCase()];
      if (k) info[k] = String(r[1] == null ? "" : r[1]);
    });
    INFO_KEYS.forEach(([k, label]) => { if (!info[k]) { info[k] = ""; warn.push("Info '" + label + "' kosong."); } });

    // Data
    const d = {};
    rows(SHEETS.DATA).slice(1).forEach((r, i) => {
      const sek = String(r[0] || "").trim(), set = String(r[1] || "").trim(), lbl = String(r[2] || "").trim();
      if (!sek || !set || !lbl) { if (r.length) warn.push("Data baris " + (i + 2) + " tidak lengkap - dilangkau."); return; }
      const v = r[3];
      const nv = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
      d[sek] = d[sek] || {};
      d[sek][set] = d[sek][set] || {};
      d[sek][set][lbl] = isNaN(nv) ? 0 : nv;
    });

    // Trend: setiap siri membawa senarai bulannya sendiri, kerana
    // seksyen berbeza boleh meliputi tempoh berbeza.
    const siri = {};
    rows(SHEETS.TREND).slice(1).forEach((r) => {
      const b = String(r[0] || "").trim(), sr = String(r[1] || "").trim();
      if (!b || !sr) return;
      siri[sr] = siri[sr] || { bulan: [], nilai: [] };
      const v = r[2];
      siri[sr].bulan.push(b);
      siri[sr].nilai.push(typeof v === "number" ? v : parseFloat(v) || 0);
    });

    // PBT: senarai dan tiga jadual
    const n2 = (v) => (typeof v === "number" ? v : parseFloat(v) || 0);
    const pbt = {
      senarai: rows(SHEETS.PBT).slice(1)
        .map((r) => ({ kod: String(r[0] || "").trim(), nama: String(r[1] || "").trim() }))
        .filter((r) => r.kod),
      jabatan: rows(SHEETS.JAB).slice(1)
        .map((r) => ({ pbt: String(r[0] || "").trim(), nama: String(r[1] || "").trim(),
          jumlah: n2(r[2]), selesai: n2(r[3]), proses: n2(r[4]), purata: n2(r[5]) }))
        .filter((r) => r.pbt && r.nama),
      zon: rows(SHEETS.ZON).slice(1)
        .map((r) => ({ pbt: String(r[0] || "").trim(), zon: String(r[1] || "").trim(),
          jumlah: n2(r[2]), selesai: n2(r[3]), proses: n2(r[4]) }))
        .filter((r) => r.pbt && r.zon),
      tertunggak: rows(SHEETS.TGK).slice(1)
        .map((r) => ({ pbt: String(r[0] || "").trim(), ruj: String(r[1] || "").trim(),
          tajuk: String(r[2] || ""), status: String(r[3] || ""),
          tahap: String(r[4] || ""), hari: n2(r[5]) }))
        .filter((r) => r.pbt && r.ruj),
    };

    const ahli = rows(SHEETS.AHLI).slice(1)
      .map((r) => ({ nama: String(r[0] || "").trim(), zon: String(r[1] || "").trim(),
        jumlah: n2(r[2]), kompleks: n2(r[3]), selesai: n2(r[4]), proses: n2(r[5]),
        tertunggak: n2(r[6]), baru: n2(r[7]) }))
      .filter((r) => r.nama && r.zon);

    // Teks: hanya lajur "Teks Ubahsuai" dibaca. Kosong = guna ayat automatik,
    // supaya ayat sentiasa mengikut data terkini.
    const teks = {};
    rows(SHEETS.TEKS).slice(1).forEach((r) => {
      const id = String(r[0] || "").trim();
      const v = r[3] == null ? "" : String(r[3]).trim();
      if (id && v) teks[id] = v;
    });

    // Semakan asas
    const mustHave = [["Media", "Ringkasan"], ["Media", "Platform"], ["CS", "Ringkasan"],
      ["Lawatan", "Ringkasan"], ["Pencegahan", "Ringkasan"]];
    mustHave.forEach(([a, b]) => {
      if (!d[a] || !d[a][b]) warn.push("Set '" + a + " / " + b + "' tiada - slaid berkaitan mungkin kosong.");
    });
    if (!Object.keys(siri).length) warn.push("Helaian Trend kosong - carta trend tidak akan dipaparkan.");
    if (!ahli.length) warn.push("Helaian Ahli_Majlis kosong - seksyen Ahli Majlis akan dilangkau.");
    if (!pbt.senarai.length) {
      warn.push("Helaian PBT_Senarai kosong - seksyen Status Semasa Aduan akan dilangkau.");
    } else {
      pbt.senarai.forEach((x) => {
        if (!d.PBT || !d.PBT[x.kod + "_Ringkasan"]) {
          warn.push("Set 'PBT / " + x.kod + "_Ringkasan' tiada dalam helaian Data.");
        }
      });
    }

    return { data: { info, d, trend: { siri }, teks, pbt, ahli }, warn };
  };
})();
