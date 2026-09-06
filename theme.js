/* theme.js — Palet & geometri. Dikongsi oleh penjana HTML dan PPTX. */
window.C = {
  INK: "12293D", SLATE: "234B69", TEAL: "13716E", OCEAN: "3C88A6",
  AMBER: "CE7C2A", CRIMSON: "9E3B3B", MOSS: "5C8A5E",
  MIST: "EFF3F6", MIST2: "E3EAEF", LINE: "D5DEE5",
  TEXT: "1B2A36", MUTED: "6B7F8D", WHITE: "FFFFFF",
  PALE_RED: "F3E3E3", DARK_RED: "8A5A5A", ICE: "9BC7D6",
  INSIGHT_BG: "E7EFF0", INSIGHT_TX: "184C4B",
};

window.SERIES = [C.TEAL, C.OCEAN, C.AMBER, C.SLATE, C.MOSS, C.CRIMSON,
  "8FAEBD", "B99A55", "7FA8A6", "9A7A6B"];

window.AGEING_COLORS = [C.TEAL, C.OCEAN, C.AMBER, C.CRIMSON];

/* Warna status aduan - hijau selesai, merah ditolak, jingga sedang diproses */
window.STATUS_WARNA = {
  "Menunggu": "8FAEBD",
  "Baru": "8FAEBD",
  "Ditugaskan": "3C88A6",
  "Telah Ditugaskan": "3C88A6",
  "Sedang Diproses": "CE7C2A",
  "Dalam Siasatan": "B99A55",
  "Dalam Tindakan": "234B69",
  "Pelaksanaan": "6E5FA0",
  "Pengesahan": "9A7A6B",
  "Sedang Disahkan": "7FA8A6",
  "Selesai Sementara": "8FB48F",
  "Diselesaikan": "5C8A5E",
  "Ditolak": "9E3B3B"
};

/* Status yang tiada dalam peta di atas diberi warna berbeza secara
   berturutan, supaya tiada dua status berkongsi warna yang sama. */
window.warnaStatus = function (label) {
  const guna = {};
  const baki = SERIES.filter((c) => Object.keys(STATUS_WARNA)
    .every((k) => STATUS_WARNA[k] !== c));
  let i = 0;
  return label.map((l) => {
    let c = STATUS_WARNA[l];
    if (!c || guna[c]) c = baki[i++ % baki.length] || C.MUTED;
    guna[c] = 1;
    return c;
  });
};

/* Warna tahap kesukaran - hijau mudah, jingga sederhana, merah kompleks */
window.KESUKARAN_WARNA = {
  "Mudah (1-3 hari)": C.MOSS, "Mudah": C.MOSS,
  "Sederhana (1-15 hari)": C.AMBER, "Sederhana": C.AMBER,
  "Kompleks (16-365 hari)": C.CRIMSON, "Kompleks": C.CRIMSON,
  "Tidak dinyatakan": C.MUTED,
};

/* Warna penilaian pengadu - hijau baik, merah lemah */
window.PENILAIAN_WARNA = {
  "Cemerlang": C.MOSS, "Sangat Baik": C.MOSS,
  "Baik": C.OCEAN, "Memuaskan": C.OCEAN,
  "Sederhana": C.AMBER,
  "Lemah": C.CRIMSON, "Tidak Memuaskan": C.CRIMSON,
};

/* Saiz font dalam carta, dalam pt. Laraskan di sini sahaja - HTML dan
   PPTX kedua-duanya membacanya. HTML menukar pt ke px secara automatik. */
window.FONT = {
  paksi: 11,      /* label paksi (nama kategori, bulan) */
  nilai: 11,      /* nombor pada bar, donat dan garisan */
  petunjuk: 11,   /* teks petunjuk (legend) */
};

window.F = { HEAD: "Cambria", BODY: "Calibri" };

/* Geometri slaid dalam inci (13.333 x 7.5 landscape) */
window.G = {
  W: 13.333, H: 7.5, M: 0.55, CW: 12.2,
  TOP: 1.25, BOT: 6.85,
  COL: 5.98, COL_X2: 6.87,
  ROW_H: 2.68, ROW_Y2: 4.17,
};

/* Pemisah ribuan: 1155 -> 1,155. Digunakan pada setiap nombor yang
   dipaparkan, dalam kedua-dua penjana. */
window.fmt = (n) => {
  const v = typeof n === "number" ? n : parseFloat(n);
  if (!isFinite(v)) return String(n);
  return String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
};

window.hx = (c) => "#" + c;
