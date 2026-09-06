/* ingest.js - Memproses fail eksport mentah terus kepada data laporan.
   Pengguna muat naik fail seperti yang dieksport dari sistem; modul ini
   mengenal pasti jenis setiap fail, mengagregat, dan menghasilkan struktur
   yang sama dengan templat Excel. */

(function () {
  const BULAN_MS = ["Januari", "Februari", "Mac", "April", "Mei", "Jun",
    "Julai", "Ogos", "September", "Oktober", "November", "Disember"];
  const BULAN_PENDEK = ["Jan", "Feb", "Mac", "Apr", "Mei", "Jun",
    "Jul", "Ogo", "Sep", "Okt", "Nov", "Dis"];

  const PBT_DARI_AGENSI = {
    "majlis bandaraya johor bahru": "MBJB",
    "majlis bandaraya iskandar puteri": "MBIP",
    "majlis bandaraya pasir gudang": "MBPG",
    "majlis perbandaran kulai": "MPKU",
  };

  const STATUS_SELESAI = ["diselesaikan", "selesai sementara"];
  const STATUS_TOLAK = ["ditolak"];

  /* ── Alat ──────────────────────────────────────────────── */
  const bersih = (v) => String(v == null ? "" : v).trim();
  const kunciBulan = (d) => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  const labelBulan = (k) => BULAN_PENDEK[parseInt(k.slice(5), 10) - 1];
  const labelPenuh = (k) => BULAN_MS[parseInt(k.slice(5), 10) - 1] + " " + k.slice(0, 4);

  /* Tarikh boleh datang sebagai objek Date, nombor siri Excel, atau teks
     seperti "19/8/2025 22.05". Kembalikan null jika tidak boleh dihurai. */
  function tarikh(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === "number") {
      const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      return isNaN(d.getTime()) ? null : d;
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  const kira = (arr, f) => {
    const o = {};
    arr.forEach((x) => { const k = f(x); if (k) o[k] = (o[k] || 0) + 1; });
    return o;
  };
  const susunTurun = (o, had) => {
    const k = Object.keys(o).sort((a, b) => o[b] - o[a]);
    const out = {};
    (had ? k.slice(0, had) : k).forEach((x) => { out[x] = o[x]; });
    return out;
  };
  const jumlah = (o) => Object.keys(o).reduce((a, k) => a + o[k], 0);

  const adalahSelesai = (s) => STATUS_SELESAI.indexOf(bersih(s).toLowerCase()) >= 0;
  const adalahTolak = (s) => STATUS_TOLAK.indexOf(bersih(s).toLowerCase()) >= 0;

  /* Aduan aktif untuk tujuan tahap kesukaran: dari Menunggu hingga
     Selesai Sementara. Aduan Diselesaikan dan Ditolak dikecualikan kerana
     tahap kesukaran hanya bermakna bagi kerja yang masih berjalan. */
  const adalahAktif = (s) => {
    const t = bersih(s).toLowerCase();
    return t !== "diselesaikan" && t !== "ditolak";
  };

  /* Hari bekerja antara dua tarikh (Isnin-Jumaat, tidak mengambil kira cuti umum) */
  function hariBekerja(a, b) {
    if (!a || !b) return 0;
    let n = 0;
    const d = new Date(a.getFullYear(), a.getMonth(), a.getDate());
    const akhir = new Date(b.getFullYear(), b.getMonth(), b.getDate());
    while (d < akhir) {
      d.setDate(d.getDate() + 1);
      const w = d.getDay();
      if (w !== 0 && w !== 6) n++;
    }
    return n;
  }

  /* Ambang SLA mengikut tahap kesukaran, dalam hari bekerja */
  const SLA = { mudah: 3, sederhana: 15, kompleks: 365 };
  const LABEL_KESUKARAN = {
    mudah: "Mudah (1-3 hari)", sederhana: "Sederhana (1-15 hari)",
    kompleks: "Kompleks (16-365 hari)",
  };

  /* ── Kenal pasti jenis setiap helaian ──────────────────── */
  function kenalPasti(namaHelaian, lajur) {
    const n = namaHelaian.toLowerCase();
    const L = lajur.map((x) => bersih(x).toLowerCase());
    const ada = (t) => L.indexOf(t) >= 0;

    if (n === "sispaa" || (ada("status sispaa") && ada("status maj"))) return "cs_sispaa";
    if (n === "emelws" || (ada("emel/chat/call") && ada("kategori"))) return "cs_emel";
    if (n.indexOf("pencegahan") >= 0) return "lap_cegah";
    if ((n.indexOf("ahli") >= 0 || n.indexOf("zon") >= 0) &&
        L.some((x) => x.indexOf("zon") >= 0 || x.indexOf("kawasan") >= 0) &&
        L.some((x) => x.indexOf("ahli majlis") >= 0 || x === "nama") &&
        !ada("perincian aduan")) return "ahli_map";
    if (ada("perincian aduan") && ada("zon ahli majlis")) return "senarai_penuh";
    if (n.indexOf("lawatan") >= 0 || (ada("tahap kesukaran") && ada("zon ahli majlis"))) return "lap_qc";
    if (ada("case id") && ada("platform")) return "media";
    if (ada("id aduan") && ada("status") && L.length <= 3) return "senarai";
    return null;
  }

  /* Hurai fail KML sempadan zon ahli majlis (dieksport dari GeoJB).
     Hanya medan ZON dan AHLI_MAJLIS diambil - geometri poligon diabaikan. */
  window.parseKMLAhli = function (teks) {
    const doc = new DOMParser().parseFromString(teks, "text/xml");
    const pm = doc.getElementsByTagName("Placemark");
    const out = [];
    for (let i = 0; i < pm.length; i++) {
      const sd = pm[i].getElementsByTagName("SimpleData");
      const d = {};
      for (let j = 0; j < sd.length; j++) {
        d[sd[j].getAttribute("name")] = (sd[j].textContent || "").trim();
      }
      const zon = d.ZON;
      const nama = d.AHLI_MAJLIS || d.AHLI_MAJLI || "";
      if (zon && nama) {
        out.push({ Zon: zon, "Ahli Majlis": nama, Parti: d.PARTI || "" });
      }
    }
    return out;
  };

  /* Kenal pasti jenis helaian dalam satu buku kerja - digunakan oleh
     antara muka untuk mengesahkan fail dimasukkan ke slot yang betul. */
  window.kenalPastiFail = function (wb) {
    const jenis = [];
    wb.SheetNames.forEach((sh) => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sh], { defval: null });
      if (!rows.length) return;
      const j = kenalPasti(sh, Object.keys(rows[0]));
      if (j && jenis.indexOf(j) < 0) jenis.push(j);
    });
    return jenis;
  };

  /* ── Proses semua fail ─────────────────────────────────── */
  window.ingestRaw = function (buku, bulanPilihan) {
    const warn = [];
    const src = { media: [], cs_sispaa: [], cs_emel: [], lap_cegah: [], lap_qc: [],
      senarai: [], senarai_penuh: [], ahli_map: [] };
    const dikenali = [];

    buku.forEach((b) => {
      let jumpa = false;
      b.wb.SheetNames.forEach((sh) => {
        const rows = XLSX.utils.sheet_to_json(b.wb.Sheets[sh], { defval: null });
        if (!rows.length) return;
        const jenis = kenalPasti(sh, Object.keys(rows[0]));
        if (!jenis) return;
        src[jenis] = src[jenis].concat(rows);
        dikenali.push(b.nama + " \u2192 " + sh + " (" + rows.length + " baris)");
        jumpa = true;
      });
      if (!jumpa) warn.push("Fail '" + b.nama + "' tidak dikenali - dilangkau.");
    });

    /* Rujukan status: ID Aduan -> Status */
    const rujStatus = {};
    src.senarai.concat(src.senarai_penuh).forEach((r) => {
      const id = bersih(r["ID Aduan"]);
      if (id) rujStatus[id] = bersih(r["Status"]);
    });

    /* ── Kumpul semua tarikh untuk tentukan bulan tersedia ── */
    const semuaBulan = {};
    const daftarBulan = (d) => { if (d) semuaBulan[kunciBulan(d)] = 1; };

    const media = src.media.map((r) => ({
      d: tarikh(r["Date Retrieved"]) || tarikh(r["Date Posting"]),
      platform: bersih(r["Platform"]) || "Tidak dinyatakan",
      kategori: bersih(r["Category"]) || "Tidak dinyatakan",
      sumber: bersih(r["Page / Group Source"]) || "Tidak dinyatakan",
      id: bersih(r["Case ID"]),
    }));
    media.forEach((r) => daftarBulan(r.d));

    const sispaa = src.cs_sispaa.map((r) => ({
      d: tarikh(r["Tarikh"]),
      id: bersih(r["ID Aduan"]),
      kategori: bersih(r["Kategori"]) || "Tidak dinyatakan",
      statusMAJ: bersih(r["Status MAJ"]),
      statusSISPAA: bersih(r["Status SISPAA"]).toUpperCase(),
    }));
    sispaa.forEach((r) => daftarBulan(r.d));

    const emel = src.cs_emel.map((r) => ({
      d: tarikh(r["TARIKH"]),
      kategori: bersih(r["KATEGORI"]) || "Tidak dinyatakan",
      saluran: bersih(r["EMEL/CHAT/CALL"]).toUpperCase(),
    }));
    emel.forEach((r) => daftarBulan(r.d));

    const cegah = src.lap_cegah.map((r) => ({
      d: tarikh(r["Tarikh"]),
      pbt: bersih(r["PBT"]) || "Tidak dinyatakan",
      agensi: bersih(r["Agensi"]) || "Tidak dinyatakan",
      kategori: bersih(r["Kategori"]) || "Tidak dinyatakan",
    }));
    cegah.forEach((r) => daftarBulan(r.d));

    let qcGunaTarikhLawatan = false;
    const qc = src.lap_qc.map((r) => {
      const ag = bersih(r["Agensi"]);
      const tl = tarikh(r["Tarikh Lawatan"]);
      if (tl) qcGunaTarikhLawatan = true;
      return {
        d: tl || tarikh(r["Tarikh aduan"]),
        dAduan: tarikh(r["Tarikh aduan"]),
        id: bersih(r["ID Aduan"]),
        pbt: PBT_DARI_AGENSI[ag.toLowerCase()] || ag || "Tidak dinyatakan",
        sumber: bersih(r["Sumber"]) || "Tidak dinyatakan",
        jabatan: bersih(r["Jabatan"]) || "Tidak dinyatakan",
        kategori: bersih(r["Kategori"]) || "Tidak dinyatakan",
        zon: bersih(r["Zon Ahli Majlis"]) || "Tidak dinyatakan",
        status: bersih(r["Status"]),
        tahap: bersih(r["Tahap Kesukaran"]) || "Tidak dinyatakan",
      };
    });
    qc.forEach((r) => daftarBulan(r.d));

    const penuh = src.senarai_penuh.map((r) => {
      const ag = bersih(r["Agensi"]);
      const tahap = bersih(r["Tahap Kesukaran"]);
      return {
        d: tarikh(r["Tarikh aduan"]),
        dk: tarikh(r["Tarikh Kemaskini"]),
        id: bersih(r["ID Aduan"]),
        pbt: PBT_DARI_AGENSI[ag.toLowerCase()] || ag || "Tidak dinyatakan",
        agensiPenuh: ag,
        sumber: bersih(r["Sumber"]) || "Tidak dinyatakan",
        jabatan: bersih(r["Jabatan"]) || "Tidak dinyatakan",
        kategori: bersih(r["Kategori"]) || "Tidak dinyatakan",
        subkategori: bersih(r["Sub Kategori"]),
        zon: bersih(r["Zon Ahli Majlis"]) || "Tidak dinyatakan",
        status: bersih(r["Status"]),
        tahap: tahap,
        tahapKey: tahap.toLowerCase(),
        rating: bersih(r["Rating"]),
      };
    });
    penuh.forEach((r) => daftarBulan(r.d));

    const bulanTersedia = Object.keys(semuaBulan).sort();
    if (!bulanTersedia.length) {
      throw new Error("Tiada tarikh sah dijumpai dalam fail yang dimuat naik.");
    }

    /* Bulan lalai: bulan penuh terakhir (bulan terkini biasanya belum lengkap) */
    const bulan = bulanPilihan && bulanTersedia.indexOf(bulanPilihan) >= 0
      ? bulanPilihan
      : (bulanTersedia.length > 1 ? bulanTersedia[bulanTersedia.length - 2]
        : bulanTersedia[bulanTersedia.length - 1]);

    /* Enam bulan sehingga bulan terpilih */
    const idx = bulanTersedia.indexOf(bulan);
    const trendBulan = bulanTersedia.slice(Math.max(0, idx - 5), idx + 1);
    /* Bulan sifar di hadapan siri bermakna rekod belum dikumpul lagi,
       bukan sifar sebenar - buang supaya carta trend tidak mengelirukan. */
    const siriBulanan = (rows) => {
      const c = kira(rows.filter((r) => r.d), (r) => kunciBulan(r.d));
      let mula = 0;
      while (mula < trendBulan.length && !c[trendBulan[mula]]) mula++;
      const guna = trendBulan.slice(mula);
      return { bulan: guna.map(labelBulan), nilai: guna.map((k) => c[k] || 0) };
    };
    const dalamBulan = (rows) => rows.filter((r) => r.d && kunciBulan(r.d) === bulan);
    const lepasKira = (rows) => {
      const k = bulanTersedia[idx - 1];
      return k ? rows.filter((r) => r.d && kunciBulan(r.d) === k).length : 0;
    };
    const delta = (kini, lepas) => (lepas ? ((kini - lepas) / lepas) * 100 : 0);

    const d = {};
    const trend = { siri: {} };
    const info = {
      tajuk: "Laporan Pemantauan Aduan Menyeluruh",
      sistem: "MyAduan Johor",
      periode: labelPenuh(bulan),
      disediakan: "Unit Pemantau Aduan",
      tarikh_jana: new Date().toLocaleDateString("ms-MY", { day: "2-digit", month: "long", year: "numeric" }),
      sumber: "Eksport sistem MAJ",
      ahli_pbt: "MBJB", ahli_minggu: "", ahli_julat: "",
    };

    /* ── MEDIA SOSIAL ──────────────────────────────────── */
    if (media.length) {
      const bln = dalamBulan(media);
      const berID = bln.filter((r) => r.id);
      const statusKira = {};
      let aktif = 0, selesai = 0, tanpaPadanan = 0;
      berID.forEach((r) => {
        const s = rujStatus[r.id];
        if (!s) { tanpaPadanan++; return; }
        statusKira[s] = (statusKira[s] || 0) + 1;
        if (adalahSelesai(s) || adalahTolak(s)) selesai++; else aktif++;
      });

      const mLepas = lepasKira(media);
      d.Media = {
        Ringkasan: {
          Capture: bln.length, Daftar_MAJ: berID.length,
          Aktif: aktif, Selesai: selesai,
        },
        Platform: susunTurun(kira(bln, (r) => r.platform)),
        Kategori: susunTurun(kira(bln, (r) => r.kategori), 10),
        Halaman: susunTurun(kira(bln, (r) => r.sumber), 6),
        Status: susunTurun(statusKira),
      };
      if (mLepas) d.Media.Ringkasan.Capture_Delta = Math.round(delta(bln.length, mLepas) * 10) / 10;
      trend.siri.Media_Capture = siriBulanan(media);

      if (!src.senarai.length && !src.senarai_penuh.length) {
        warn.push("Fail senarai aduan tidak dimuat naik - status aduan media sosial tidak dapat dipaparkan.");
      } else if (tanpaPadanan) {
        warn.push(tanpaPadanan + " ID aduan media sosial tiada dalam senarai aduan.");
      }
      if (bln.length - berID.length) {
        warn.push((bln.length - berID.length) + " aduan media sosial tiada Case ID - tidak dikira sebagai didaftarkan.");
      }
    }

    /* ── CUSTOMER SERVICE ──────────────────────────────── */
    if (sispaa.length) {
      const bln = dalamBulan(sispaa);
      const siapKerja = bln.filter((r) => adalahSelesai(r.statusMAJ) || adalahTolak(r.statusMAJ));
      const ditutup = bln.filter((r) => r.statusSISPAA === "SELESAI");
      const belumTutup = siapKerja.filter((r) => r.statusSISPAA !== "SELESAI");
      const aktifRows = bln.filter((r) => !adalahSelesai(r.statusMAJ) && !adalahTolak(r.statusMAJ));

      const cLepas = lepasKira(sispaa);
      d.CS = {
        Ringkasan: {
          Diterima: bln.length,
          Aktif: aktifRows.length,
          Ditutup: ditutup.length,
          Belum_Tutup: belumTutup.length,
        },
        Kategori: susunTurun(kira(bln, (r) => r.kategori), 10),
        Status_Aktif: susunTurun(kira(aktifRows, (r) => r.statusMAJ)),
      };
      if (cLepas) d.CS.Ringkasan.Diterima_Delta = Math.round(delta(bln.length, cLepas) * 10) / 10;
      trend.siri.CS_Diterima = siriBulanan(sispaa);
    }

    if (emel.length) {
      const bln = dalamBulan(emel);
      d.CS = d.CS || { Ringkasan: {} };
      d.CS.Ringkasan.Emel_Dijawab = bln.filter((r) => r.saluran === "EMEL").length;
      d.CS.Ringkasan.WA_Dijawab = bln.filter((r) => r.saluran !== "EMEL").length;
      d.CS.Isu_Emel = susunTurun(kira(bln, (r) => r.kategori), 6);
      trend.siri.CS_Emel = siriBulanan(emel.filter((r) => r.saluran === "EMEL"));
      trend.siri.CS_WA = siriBulanan(emel.filter((r) => r.saluran !== "EMEL"));
    }

    /* ── LAPANGAN: LAWATAN TAPAK (QC) ──────────────────── */
    if (qc.length) {
      const bln = dalamBulan(qc);
      if (!qcGunaTarikhLawatan) {
        warn.push("Helaian lawatan tiada lajur 'Tarikh Lawatan' - lawatan dikira mengikut " +
          "tarikh aduan, bukan tarikh lawatan sebenar.");
      }
      const bergerak = bln.filter((r) => adalahSelesai(r.status)).length;
      const lLepas = lepasKira(qc);
      d.Lawatan = {
        Ringkasan: {
          Total: bln.length, Bergerak: bergerak,
          Tersangkut: bln.length - bergerak,
        },
        PBT: susunTurun(kira(bln, (r) => r.pbt)),
        Kategori: susunTurun(kira(bln, (r) => r.kategori), 10),
        Zon: susunTurun(kira(bln, (r) => r.zon), 8),
        Status: susunTurun(kira(bln, (r) => r.status)),
        Sumber: susunTurun(kira(bln, (r) => r.sumber)),
        Kesukaran: susunTurun(kira(bln.filter((r) => adalahAktif(r.status)), (r) => r.tahap)),
        Jabatan: susunTurun(kira(bln, (r) => r.jabatan), 8),
      };
      if (lLepas) d.Lawatan.Ringkasan.Delta = Math.round(delta(bln.length, lLepas) * 10) / 10;
      trend.siri.Lawatan = siriBulanan(qc);
    }

    /* ── LAPANGAN: ADUAN PENCEGAHAN ────────────────────── */
    if (cegah.length) {
      const bln = dalamBulan(cegah);
      const pLepas = lepasKira(cegah);
      d.Pencegahan = {
        Ringkasan: { Total: bln.length },
        PBT: susunTurun(kira(bln, (r) => r.pbt)),
        Kategori: susunTurun(kira(bln, (r) => r.kategori), 10),
        Agensi: susunTurun(kira(bln, (r) => r.agensi), 8),
      };
      if (pLepas) d.Pencegahan.Ringkasan.Delta = Math.round(delta(bln.length, pLepas) * 10) / 10;
      trend.siri.Pencegahan = siriBulanan(cegah);
    }

    /* ── STATUS SEMASA ADUAN (IKUT PBT) ────────────────── */
    let pbtOut = null;
    if (penuh.length) {
      const bln = dalamBulan(penuh);
      /* Tarikh rujukan untuk mengira umur aduan yang masih aktif:
         hari terakhir bulan laporan. */
      const akhirBulan = new Date(parseInt(bulan.slice(0, 4), 10),
        parseInt(bulan.slice(5), 10), 0);

      bln.forEach((r) => {
        const tamat = adalahSelesai(r.status) || adalahTolak(r.status);
        const hingga = tamat && r.dk ? r.dk : akhirBulan;
        r.umur = hariBekerja(r.d, hingga);
        r.ambang = SLA[r.tahapKey] || SLA.sederhana;
        r.lepasSLA = r.umur > r.ambang;
        r.tamat = tamat;
      });

      const kodSenarai = [];
      const seenKod = {};
      bln.forEach((r) => {
        if (!seenKod[r.pbt]) { seenKod[r.pbt] = r.agensiPenuh || r.pbt; kodSenarai.push(r.pbt); }
      });

      pbtOut = { senarai: [], jabatan: [], zon: [], tertunggak: [] };

      kodSenarai.forEach((kod) => {
        const rows = bln.filter((r) => r.pbt === kod);
        const selesai = rows.filter((r) => adalahSelesai(r.status)).length;
        const lepasSLA = rows.filter((r) => r.lepasSLA).length;
        const lepasBulan = penuh.filter((r) => r.pbt === kod && r.d &&
          kunciBulan(r.d) === bulanTersedia[idx - 1]).length;

        const katSemua = kira(rows, (r) => r.kategori);
        const lainLain = Object.keys(katSemua)
          .filter((k) => k.toLowerCase().indexOf("lain") === 0)
          .reduce((a, k) => a + katSemua[k], 0);
        const kat = {};
        Object.keys(katSemua).forEach((k) => {
          if (k.toLowerCase().indexOf("lain") !== 0) kat[k] = katSemua[k];
        });

        const kesKira = {};
        rows.filter((r) => adalahAktif(r.status)).forEach((r) => {
          const lab = LABEL_KESUKARAN[r.tahapKey] || r.tahap || "Tidak dinyatakan";
          kesKira[lab] = (kesKira[lab] || 0) + 1;
        });
        const kes = {};
        ["Mudah (1-3 hari)", "Sederhana (1-15 hari)", "Kompleks (16-365 hari)", "Tidak dinyatakan"]
          .forEach((lab) => { if (kesKira[lab]) kes[lab] = kesKira[lab]; });
        Object.keys(kesKira).forEach((lab) => { if (!kes[lab]) kes[lab] = kesKira[lab]; });

        d.PBT = d.PBT || {};
        d.PBT[kod + "_Ringkasan"] = {
          Jumlah: rows.length, Selesai: selesai, Lepas_SLA: lepasSLA,
          Lain_Lain: lainLain,
        };
        if (lepasBulan) {
          d.PBT[kod + "_Ringkasan"].Delta = Math.round(delta(rows.length, lepasBulan) * 10) / 10;
        }
        d.PBT[kod + "_Status"] = susunTurun(kira(rows, (r) => r.status));
        d.PBT[kod + "_Kategori"] = susunTurun(kat, 10);
        d.PBT[kod + "_Sumber"] = susunTurun(kira(rows, (r) => r.sumber));
        d.PBT[kod + "_Kesukaran"] = kes;
        /* Susunan tetap dari terbaik ke terlemah supaya carta mudah dibaca */
        const ratKira = kira(rows.filter((r) => r.rating), (r) => r.rating);
        if (Object.keys(ratKira).length) {
          const rat = {};
          ["Cemerlang", "Sangat Baik", "Baik", "Memuaskan", "Sederhana", "Lemah", "Tidak Memuaskan"]
            .forEach((lab) => { if (ratKira[lab]) rat[lab] = ratKira[lab]; });
          Object.keys(ratKira).forEach((lab) => { if (!rat[lab]) rat[lab] = ratKira[lab]; });
          d.PBT[kod + "_Penilaian"] = rat;
        }

        /* Trend hanya bulan yang benar-benar ada data bagi PBT ini */
        const cAll = kira(penuh.filter((r) => r.pbt === kod && r.d), (r) => kunciBulan(r.d));
        const adaBulan = trendBulan.filter((k) => cAll[k]);
        if (adaBulan.length >= 2) {
          trend.siri[kod + "_Aduan"] = {
            bulan: adaBulan.map(labelBulan), nilai: adaBulan.map((k) => cAll[k]),
          };
        }

        pbtOut.senarai.push({ kod: kod, nama: seenKod[kod] || kod });

        const jabKira = {};
        rows.filter((r) => r.jabatan && r.jabatan !== "Tidak dinyatakan").forEach((r) => {
          const j = jabKira[r.jabatan] || (jabKira[r.jabatan] = { n: 0, s: 0, p: 0, hari: 0, nh: 0 });
          j.n++;
          if (adalahSelesai(r.status)) { j.s++; j.hari += r.umur; j.nh++; }
          else if (!adalahTolak(r.status)) j.p++;
        });
        Object.keys(jabKira).forEach((nama) => {
          const j = jabKira[nama];
          pbtOut.jabatan.push({ pbt: kod, nama: nama, jumlah: j.n, selesai: j.s,
            proses: j.p, purata: j.nh ? Math.round((j.hari / j.nh) * 10) / 10 : 0 });
        });

        const zonKira = {};
        rows.forEach((r) => {
          const z = zonKira[r.zon] || (zonKira[r.zon] = { n: 0, s: 0, p: 0 });
          z.n++;
          if (adalahSelesai(r.status)) z.s++;
          else if (!adalahTolak(r.status)) z.p++;
        });
        Object.keys(zonKira).forEach((z) => {
          pbtOut.zon.push({ pbt: kod, zon: z, jumlah: zonKira[z].n,
            selesai: zonKira[z].s, proses: zonKira[z].p });
        });

        /* Aduan tertunggak: masih aktif dan telah melepasi ambang SLA.
           Tajuk menggunakan sub kategori, bukan perincian aduan, kerana
           perincian boleh mengandungi maklumat peribadi pengadu. */
        rows.filter((r) => !r.tamat && r.lepasSLA)
          .sort((a, b) => b.umur - a.umur).slice(0, 5)
          .forEach((r) => {
            pbtOut.tertunggak.push({ pbt: kod, ruj: r.id,
              tajuk: r.subkategori || r.kategori, status: r.status,
              tahap: r.tahap, hari: r.umur });
          });
      });
    }

    /* ── AHLI MAJLIS (daripada Zon Ahli Majlis) ────────── */
    let ahliOut = [];
    if (penuh.length) {
      const blnA = dalamBulan(penuh);
      /* Pemetaan zon -> nama ahli majlis, jika helaian pemetaan dimuat naik */
      /* Nama zon dalam eksport aduan dan dalam KML boleh berbeza ruang
         atau huruf besar/kecil - normalkan sebelum dipadankan. */
      const kunciZon = (z) => bersih(z).toUpperCase().replace(/\s+/g, " ")
        .replace(/\s*\/\s*/g, "/");
      const petaNama = {};
      src.ahli_map.forEach((r) => {
        const k = Object.keys(r);
        const kz = k.find((x) => /zon|kawasan/i.test(x));
        const kn = k.find((x) => /ahli majlis|nama/i.test(x));
        if (kz && kn) {
          const z = kunciZon(r[kz]);
          if (z) petaNama[z] = bersih(r[kn]);
        }
      });

      /* PBT dominan menentukan label seksyen */
      const kiraPBT = kira(blnA, (r) => r.pbt);
      const pbtUtama = Object.keys(kiraPBT).sort((a, b) => kiraPBT[b] - kiraPBT[a])[0];
      info.ahli_pbt = pbtUtama || "";
      info.ahli_minggu = labelPenuh(bulan);
      info.ahli_julat = "";

      /* Seksyen ini melaporkan aduan yang dibawa masuk oleh ahli majlis,
         jadi hanya baris dengan Sumber = "Ahli Majlis" dikira. */
      const barisAhli = blnA.filter((r) => r.pbt === pbtUtama &&
        /ahli\s*majlis/i.test(r.sumber) && r.zon && r.zon !== "Tidak dinyatakan");

      if (barisAhli.length) {
        const kesA = {};
        barisAhli.filter((r) => adalahAktif(r.status)).forEach((r) => {
          const lab = LABEL_KESUKARAN[r.tahapKey] || r.tahap || "Tidak dinyatakan";
          kesA[lab] = (kesA[lab] || 0) + 1;
        });
        const kesUrut = {};
        ["Mudah (1-3 hari)", "Sederhana (1-15 hari)", "Kompleks (16-365 hari)", "Tidak dinyatakan"]
          .forEach((l) => { if (kesA[l]) kesUrut[l] = kesA[l]; });
        Object.keys(kesA).forEach((l) => { if (!kesUrut[l]) kesUrut[l] = kesA[l]; });

        d.Ahli = {
          Kategori: susunTurun(kira(barisAhli, (r) => r.kategori), 10),
          Kesukaran: kesUrut,
          Status: susunTurun(kira(barisAhli, (r) => r.status)),
        };
      }

      /* Senarai zon diambil daripada SEMUA aduan PBT ini pada bulan
         berkenaan, bukan hanya aduan bersumber Ahli Majlis. Zon tanpa
         aduan ahli majlis tetap disenaraikan dengan nilai sifar. */
      const zonKira = {};
      const kosong = () => ({ jumlah: 0, kompleks: 0, selesai: 0, selesai_nk: 0,
        ditolak: 0, tertunggak: 0, baru: 0 });
      blnA.filter((r) => r.pbt === pbtUtama && r.zon && r.zon !== "Tidak dinyatakan")
        .forEach((r) => { if (!zonKira[r.zon]) zonKira[r.zon] = kosong(); });
      Object.keys(petaNama).forEach((k) => {
        const adaZon = Object.keys(zonKira).some((z) => kunciZon(z) === k);
        if (!adaZon && pbtUtama) zonKira[k] = kosong();
      });

      barisAhli
        .forEach((r) => {
          const z = zonKira[r.zon] || (zonKira[r.zon] = kosong());
          z.jumlah++;
          const kompleks = r.tahapKey === "kompleks";
          if (kompleks) z.kompleks++;
          if (adalahSelesai(r.status)) { z.selesai++; if (!kompleks) z.selesai_nk++; }
          else if (adalahTolak(r.status)) z.ditolak++;
          else if (r.lepasSLA) z.tertunggak++;
          else if (/menunggu|baru/i.test(r.status)) z.baru++;
        });

      ahliOut = Object.keys(zonKira).map((z) => {
        const x = zonKira[z];
        return {
          nama: petaNama[kunciZon(z)] || "",
          zon: z, jumlah: x.jumlah, kompleks: x.kompleks,
          selesai: x.selesai, selesai_nk: x.selesai_nk,
          ditolak: x.ditolak, tertunggak: x.tertunggak, baru: x.baru,
          proses: Math.max(0, x.jumlah - x.selesai - x.ditolak - x.tertunggak - x.baru),
        };
      }).sort((a, b) => b.jumlah - a.jumlah);

      if (!barisAhli.length) {
        warn.push("Tiada aduan bersumber 'Ahli Majlis' pada " + labelPenuh(bulan) +
          " - seksyen Ahli Majlis dilangkau.");
      }
      if (ahliOut.length && !Object.keys(petaNama).length) {
        warn.push("Nama ahli majlis tiada dalam eksport - jadual prestasi zon memaparkan " +
          "kawasan sahaja. Muat naik helaian pemetaan (Zon | Ahli Majlis) untuk memaparkan nama.");
      }
    }

    /* Amaran bulan tidak lengkap: eksport selalunya diambil pertengahan bulan,
       jadi bulan terakhir boleh kelihatan jatuh mendadak tanpa sebab sebenar. */
    const semak = [["Media Sosial", "Media_Capture"], ["Customer Service", "CS_Diterima"],
      ["Lawatan Tapak (QC)", "Lawatan"], ["Aduan Pencegahan", "Pencegahan"]];
    semak.forEach((x) => {
      const s2 = trend.siri[x[1]];
      if (!s2 || s2.nilai.length < 4) return;
      const n = s2.nilai.length;
      const kini = s2.nilai[n - 1];
      const purata = (s2.nilai[n - 4] + s2.nilai[n - 3] + s2.nilai[n - 2]) / 3;
      if (purata > 10 && kini < purata * 0.4) {
        warn.push("Data " + x[0] + " bagi " + labelPenuh(bulan) + " (" + kini +
          ") jauh lebih rendah daripada purata tiga bulan sebelum (" + Math.round(purata) +
          ") - kemungkinan bulan belum lengkap dalam fail eksport.");
      }
    });

    Object.keys(d).forEach((sek) => {
      const kosong = Object.keys(d[sek]).filter((s) => s !== "Ringkasan" && !jumlah(d[sek][s]));
      kosong.forEach((s) => { delete d[sek][s]; });
    });

    return {
      data: { info: info, d: d, trend: trend, teks: {}, pbt: pbtOut, ahli: ahliOut },
      warn: warn,
      dikenali: dikenali,
      bulan: bulan,
      bulanTersedia: bulanTersedia.map((k) => ({ kunci: k, label: labelPenuh(k) })),
    };
  };
})();
