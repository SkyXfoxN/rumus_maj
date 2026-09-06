/* insights.js - Menghasilkan ayat dapatan secara automatik daripada data.
   Setiap ayat boleh ditindih secara manual melalui helaian "Teks" dalam Excel. */

(function () {
  const sum = (a) => a.reduce((x, y) => x + y, 0);
  const pc = (a, b) => (b ? Math.round((a / b) * 100) : 0);
  const cap = (t) => (t ? t.charAt(0).toUpperCase() + t.slice(1) : t);

  /* Taburan: item tertinggi dan bahagiannya */
  window.fTop = function (obj, unit) {
    const k = Object.keys(obj || {});
    if (!k.length) return "";
    const v = k.map((x) => obj[x]);
    const tot = sum(v);
    let hi = 0;
    v.forEach((x, i) => { if (x > v[hi]) hi = i; });
    const dua = k.length > 1 && k.length <= 5
      ? " Dua teratas mewakili " + pc(v[hi] + Math.max.apply(null, v.filter((_, i) => i !== hi)), tot) + "%."
      : "";
    return k[hi] + " mendahului dengan " + fmt(v[hi]) + " " + (unit || "aduan") +
      " (" + pc(v[hi], tot) + "% daripada " + fmt(tot) + ")." + dua;
  };

  /* Trend: arah bulan semasa dan purata */
  window.fTrend = function (vals, bulan, unit) {
    if (!vals || vals.length < 2) return "";
    const n = vals.length, kini = vals[n - 1], lepas = vals[n - 2];
    const dlt = lepas ? ((kini - lepas) / lepas) * 100 : 0;
    const arah = dlt >= 0 ? "meningkat" : "menurun";
    const purata = Math.round(sum(vals) / n);
    const tinggi = Math.max.apply(null, vals);
    return fmt(kini) + " " + (unit || "aduan") + " bulan ini, " + arah + " " +
      Math.abs(dlt).toFixed(1) + "% berbanding " + bulan[n - 2] +
      ". Purata " + n + " bulan: " + fmt(purata) +
      (kini === tinggi ? ". Ini paras tertinggi dalam tempoh ini." : ".");
  };

  /* Ageing: berapa yang melepasi ambang */
  window.fAge = function (obj) {
    const k = Object.keys(obj || {});
    if (!k.length) return "";
    const tot = sum(k.map((x) => obj[x]));
    const lama = (obj["16-30 hari"] || 0) + (obj["30+ hari"] || 0);
    const sgtLama = obj["30+ hari"] || 0;
    return fmt(lama) + " aduan (" + pc(lama, tot) + "%) telah melebihi 15 hari" +
      (sgtLama ? ", termasuk " + fmt(sgtLama) + " yang melebihi 30 hari" : "") + ".";
  };

  /* Status: berapa masih bergerak berbanding selesai */
  window.fStatus = function (obj, selesaiKeys) {
    const k = Object.keys(obj || {});
    if (!k.length) return "";
    const tot = sum(k.map((x) => obj[x]));
    const sel = (selesaiKeys || []).reduce((a, x) => a + (obj[x] || 0), 0);
    if (sel) {
      return pc(sel, tot) + "% telah diselesaikan; baki " + fmt(tot - sel) +
        " masih dalam proses.";
    }
    return window.fTop(obj, "aduan");
  };

  /* Perbandingan dua angka */
  window.fSplit = function (a, b, la, lb) {
    const t = a + b;
    return cap(la) + " " + pc(a, t) + "% (" + fmt(a) + "); " + lb + " " + pc(b, t) + "% (" + fmt(b) + ").";
  };
})();
