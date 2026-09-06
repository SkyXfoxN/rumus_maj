# Penjana Laporan Pemantauan Aduan — MyAduan Johor

Laman statik. Muat naik fail Excel, lihat 49 slaid terus dalam pelayar, muat turun sebagai **PowerPoint** atau **PDF landscape saiz slaid** (13.33" x 7.5").

Tiada pelayan, tiada pangkalan data. Semua pemprosesan berlaku dalam pelayar pengguna.

---

## Pasang di Netlify

1. Buka [app.netlify.com/drop](https://app.netlify.com/drop)
2. Seret **seluruh folder ini** ke dalam kotak
3. Siap — Netlify beri URL

Untuk kemas kini: seret folder yang sama sekali lagi.

**Penting:** Jangan letak fail Excel berisi data aduan sebenar di dalam folder ini. Fail yang diletak di sini boleh dicapai oleh sesiapa yang tahu URL. Excel sepatutnya dimuat naik dari komputer setiap kali digunakan — dengan cara itu data tidak pernah meninggalkan pelayar.

Kalau mahu had capaian, aktifkan **Password protection** atau **Netlify Identity** di Site settings.

---

## Cara guna

Ada dua cara memasukkan data.

### A — Muat naik data eksport (paling cepat)

Panel muat naik ada **empat slot**. Masukkan fail eksport ikut jenis, kemudian tekan **Jana slaid**.

| Slot | Isi | Mengisi seksyen |
|---|---|---|
| 1. Senarai Aduan (Raw Data) | Eksport penuh dari sistem MAJ, dan fail KML sempadan ahli majlis | Status Semasa Aduan ikut PBT, nama ahli majlis |
| 2. Perkhidmatan Pelanggan | Helaian SISPAA & EmelWS | Perkhidmatan Pelanggan |
| 3. Media Sosial | Log aduan media sosial dengan Case ID | Media Sosial |
| 4. Lapangan | Helaian Aduan Pencegahan & Lawatan Sistem (QC) | Lapangan |

**Satu fail boleh mengandungi beberapa helaian.** Sistem mengenal pasti mengikut nama helaian dan tajuk lajur, bukan nama fail. Jadi satu buku kerja gabungan dengan helaian `Media`, `SISPAA`, `EmelWS`, `Aduan Pencegahan` dan `Aduan Lawatan Sistem (QC)` boleh dimasukkan ke mana-mana slot dan tetap diproses dengan betul. Dua fail sudah memadai untuk deck penuh: satu buku gerak kerja gabungan, satu senarai aduan penuh.

Setiap slot boleh menerima **beberapa fail sekali gus** — masukkan eksport bulan semasa dan bulan-bulan sebelum supaya carta trend ada sekurang-kurangnya dua titik.

Slot yang berisi bertukar hijau. Kalau fail dimasukkan ke slot yang salah, sistem beri amaran dan tandakan fail berkenaan — tetapi ia tetap diproses mengikut kandungan sebenar, bukan mengikut slot. Slot hanyalah panduan susunan.

Fail yang dikenali:

| Fail | Helaian | Digunakan untuk |
|---|---|---|
| Data Media Sosial | apa-apa helaian dengan lajur `Category` + `Platform` + `Case ID` | Seksyen Media Sosial |
| Data Perkhidmatan Pelanggan | `SISPAA`, `EmelWS` | Seksyen Perkhidmatan Pelanggan |
| Data Lapangan | `Aduan Pencegahan`, `Aduan Lawatan Sistem (QC)` | Seksyen Lapangan |
| Senarai aduan (ringkas) | dua lajur: `ID Aduan` + `Status` | Status aduan media sosial (padanan ID) |
| Senarai aduan (penuh) | 38 lajur, ada `Perincian aduan` + `Zon Ahli Majlis` | Seluruh seksyen Status Semasa Aduan (ikut PBT) |
| Sempadan ahli majlis | fail `.kml` dengan medan `ZON` + `AHLI_MAJLIS` | Nama ahli majlis dalam jadual prestasi zon |

Fail boleh dicampur. Muat naik eksport media sosial, CS, lapangan, dan senarai aduan penuh sekali gus — setiap satu mengisi bahagiannya sendiri dan bersama-sama membentuk satu deck lengkap.

Selepas **Jana slaid** ditekan, panel bertutup sendiri dan **pemilih bulan** di bar atas menjadi aktif. Tukar bulan, slaid dijana semula serta-merta tanpa perlu muat naik semula.

**Seksyen tanpa data dilangkau automatik.** Kalau fail media sosial tidak dimuat naik, seksyen itu tidak muncul langsung. Begitu juga blok dalam slaid — kalau eksport tiada medan PBT, blok berkenaan digantikan dengan pecahan lain yang ada.

### B — Templat Excel (kawalan penuh)

| Butang | Fungsi |
|---|---|
| **Templat Excel** | Muat turun fail Excel dengan struktur betul dan data semasa sebagai contoh |
| **Muat naik templat** | Baca fail yang telah diisi, kemas kini semua slaid serta-merta |

Butang **Muat naik data** di bar atas membuka dan menutup panel slot.
| **Muat turun PPTX** | Jana fail PowerPoint yang boleh diedit |
| **Muat turun PDF** | Buka kotak cetakan pelayar |

### Melaras saiz font dalam carta

Semua saiz font carta ditakrifkan dalam `assets/theme.js`, dalam **pt**:

```js
window.FONT = {
  paksi: 11,      // label paksi (nama kategori, bulan)
  nilai: 11,      // nombor pada bar, donat dan garisan
  petunjuk: 11,   // teks petunjuk (legend)
};
```

Tukar nilai di sini, kedua-dua paparan HTML dan fail PPTX berubah serentak. Penjana HTML menukar pt kepada px secara automatik.

### Seksyen Ahli Majlis

Seksyen ini melaporkan **aduan yang bersumberkan Ahli Majlis** sahaja — iaitu baris dengan `Sumber` = "Ahli Majlis" dalam eksport senarai aduan. Ia bukan semua aduan dalam zon berkenaan, tetapi aduan yang dibawa masuk melalui ahli majlis.

Kalau tiada aduan bersumber Ahli Majlis dalam bulan terpilih, seksyen ini dilangkau dan amaran diberikan.

### Nama ahli majlis

Eksport senarai aduan mengandungi `Zon Ahli Majlis` tetapi bukan nama ahli majlis. Ada dua cara memasukkan nama:

**Fail KML (paling mudah)** — muat naik fail sempadan zon ahli majlis dari GeoJB (contoh `AM2025.kml`) ke slot 1. Sistem membaca medan `ZON` dan `AHLI_MAJLIS` terus daripada fail; geometri poligon diabaikan. Tiada penukaran format diperlukan.

**Helaian Excel** — dua lajur, `Zon` dan `Ahli Majlis`. Nama helaian mesti mengandungi perkataan "Ahli" atau "Zon".

Nama zon dinormalkan sebelum dipadankan (huruf besar, ruang berganda, ruang sekeliling garis miring), jadi "SETIA  AUSTIN" dan "Setia Austin" dianggap sama.

Tanpa pemetaan, jadual prestasi zon tetap dijana, cuma lajur nama digugurkan.

### Tetapan cetakan PDF

Dalam kotak cetakan Chrome:
- Destination: **Save as PDF**
- Paper size: **Custom** 13.33 x 7.5 inci, atau **Landscape**
- Margins: **None**
- Options: **Background graphics** dihidupkan, **Headers and footers** dimatikan

---

## Struktur fail Excel

Empat helaian. Nama helaian dan tajuk lajur **tidak boleh diubah**.

**`Info`** — Kunci | Nilai
Tajuk laporan, periode, siapa sediakan, dan sebagainya.

**`Data`** — Seksyen | Set | Label | Nilai
Semua pecahan angka. Contoh:

| Seksyen | Set | Label | Nilai |
|---|---|---|---|
| Media | Platform | Facebook | 97 |
| Media | Platform | Instagram | 65 |
| CS | Ringkasan | Diterima | 318 |

Susunan baris menentukan susunan dalam carta. Untuk tukar susunan kategori, susun semula baris dalam Excel.

**`Trend`** — Bulan | Siri | Nilai
Data untuk semua carta trend. Setiap siri boleh meliputi tempoh berbeza — seksyen Media/CS/Lapangan guna enam bulan, seksyen PBT guna sepanjang tahun.

**`PBT_Senarai`** — Kod | Nama Penuh
Menentukan PBT mana yang dipaparkan dan susunannya. Buang satu baris = PBT itu hilang dari deck. Tambah baris = PBT baharu muncul (asalkan datanya ada dalam helaian Data).

**`PBT_Jabatan`** — PBT | Jabatan | Jumlah | Selesai | Dalam Proses | Purata (hari)
Kadar selesai dikira sendiri dan diwarnakan: merah bawah 50%, kuning 50–70%, hijau atas 70%. Jabatan dengan kurang daripada 3 aduan tidak disenaraikan.

**`PBT_Zon`** — PBT | Zon / Kawasan | Jumlah | Selesai | Dalam Proses
Lima zon tertinggi setiap PBT dipaparkan.

**`PBT_Tertunggak`** — PBT | No. Rujukan | Tajuk Aduan | Status | Tahap | Hari Tertunggak
Lima aduan terlama setiap PBT dipaparkan.

**`Ahli_Majlis`** — Ahli Majlis | Zon / Kawasan | Jumlah | Kompleks | Selesai | Dalam Proses | Tertunggak | Baru Diterima
Laporan mingguan, **MBJB sahaja** buat masa ini. Tempoh mingguan diisi dalam helaian `Info`.

Dua perkara dikendalikan automatik:

**Aduan Kompleks dikeluarkan dari kiraan kadar selesai.** SLA Kompleks ialah 365 hari bekerja — mustahil disiapkan dalam tempoh mingguan. Kalau ia dikira, kadar selesai setiap zon tertekan sepanjang tahun tanpa sebab. Kadar dikira sebagai `Selesai ÷ (Jumlah − Kompleks)`.

**Zon disusun mengikut jumlah aduan diterima**, paling banyak di atas. 12 baris setiap slaid; bilangan slaid bertambah sendiri mengikut bilangan zon.

**`Teks`** — ID | Perkara | Teks Automatik | Teks Ubahsuai
Setiap tajuk slaid, tajuk kad, dan ayat dapatan di bawah setiap carta.

Lajur **Teks Automatik** hanya rujukan — ia menunjukkan ayat yang sistem jana sendiri daripada data. Isi lajur **Teks Ubahsuai** hanya kalau mahu tulis sendiri.

Biarkan kosong = ayat dijana automatik dan sentiasa ikut data terkini.
Isi = ayat kau yang digunakan, kekal sampai kau padam.

Kalau ada set yang hilang, aplikasi tetap berjalan tetapi memaparkan amaran dan slaid berkenaan akan kosong.

---

## Struktur kod

```
index.html
assets/
  theme.js          Palet warna & geometri slaid
  spec.js           TAKRIFAN 30 SLAID  <- edit di sini
  render-html.js    Lukis slaid untuk paparan & PDF
  render-pptx.js    Bina fail PowerPoint
  excel.js          Jana templat & hurai fail
  app.js            Pengikat antara muka
  app.css           Gaya slaid & peraturan cetakan
  ingest.js         Pemprosesan fail eksport mentah
  insights.js       Enjin ayat dapatan automatik
  data-default.js   Data demo
  vendor/           Pustaka (disimpan setempat, tiada CDN)
```

**Prinsip:** susun atur ditakrifkan **sekali sahaja** dalam `spec.js`. Kedua-dua penjana membacanya. Ubah kedudukan atau kandungan slaid di situ, dan PPTX serta PDF berubah serentak.

### Tukar warna

Semua dalam `assets/theme.js`:

```js
window.C = {
  INK: "12293D",     // navy — slaid tajuk & pembahagi
  TEAL: "13716E",    // aksen data utama
  AMBER: "CE7C2A",   // amaran
  CRIMSON: "9E3B3B", // kritikal
  ...
};
```

### Tambah slaid

Dalam `spec.js`, tambah objek ke dalam `slides`:

```js
slides.push({
  kind: "content", section: "Media Sosial", title: "Tajuk Baru", periode,
  blocks: [
    { t: "stat", x: G.M, y: 1.25, w: 3.9, h: 1.35,
      label: "Label", value: 123, color: C.INK, sub: "Nota kecil" },
    { t: "card", x: G.M, y: 2.85, w: G.CW, h: 4.0, title: "Tajuk kad",
      body: { k: "col", x: G.M + 0.3, y: 3.45, w: G.CW - 0.6, h: 3.2,
              labels: [...], values: [...], color: C.TEAL } },
  ],
});
```

Semua kedudukan dalam **inci**. Slaid ialah 13.333 x 7.5 inci.

Jenis blok: `stat`, `card`, `segment`, `insight`, `funnel`, `table`, `rank`, `label`
Jenis carta (dalam `body`): `col`, `bar`, `donut`, `line`

---

## Diuji

| Perkara | Status |
|---|---|
| Takrifan 30 slaid, kedudukan blok | Lulus — tiada blok terkeluar sempadan |
| Penjanaan PPTX | Lulus — fail sah, dibuka tanpa ralat |
| Kitaran Excel (jana templat, baca semula) | Lulus — 268 nilai + 100 baris jadual, 0 perbezaan |
| Tindihan teks manual | Lulus — 218 baris boleh diubah |
| Binaan DOM untuk paparan HTML | Lulus — 50 carta, tiada ralat |
| **Paparan visual dalam pelayar sebenar** | **Belum diuji** |
| **PDF melalui cetakan pelayar** | **Belum diuji** |

Dua perkara terakhir tidak dapat diuji kerana persekitaran pembinaan tiada pelayar. Buka `index.html` dan semak paparan serta cetakan sebelum guna untuk laporan sebenar.

---

## Belum sedia

| Perkara | Tindakan |
|---|---|
| Sumber data Media Sosial | Bina helaian log |
| Sumber data e-mel / WhatsApp | Bina helaian log |
| Penanda aduan pencegahan | Tambah kolum baru **di hujung** Master_Engine |
| Butiran escalation | Tentukan penerima Peringatan 1-3 dan tindakan selepas Peringatan 3 |
| Data dalam `data-default.js` | Semua **dummy** — bukan angka sebenar |


---

## Susunan deck

| Seksyen | Slaid |
|---|---|
| Media Sosial | 4 |
| Perkhidmatan Pelanggan | 4 |
| Lapangan | 6 |
| Status Semasa Aduan (ikut PBT) | 25 |
| Ahli Majlis (MBJB) | 4 |
| Tajuk + 5 pembahagi seksyen | 6 |
| **Jumlah** | **49** |

Seksyen PBT: satu slaid perbandingan, kemudian enam slaid setiap PBT dengan format identik — Statistik, Status, Kategori & Sumber, Prestasi Jabatan, Kualiti Penyelesaian, Kawasan & Tertunggak.

Seksyen Ahli Majlis: statistik keseluruhan, kategori & tahap kesukaran, kemudian jadual prestasi zon. Bilangan slaid jadual bertambah sendiri mengikut bilangan zon (12 baris setiap slaid). Untuk memasukkan PBT lain, tambah barisnya dalam helaian `Ahli_Majlis` dan tukar medan PBT dalam helaian `Info`.

## Seksyen yang disorokkan

Escalation telah dikeluarkan buat masa ini. Kodnya boleh dipulangkan semula bila diperlukan.


---

## Apa yang boleh dan tidak boleh dijana daripada data mentah

Fail eksport tidak mengandungi setiap medan yang templat Excel boleh isi. Yang berikut **tiada** dalam eksport semasa, jadi blok berkenaan dilangkau:

| Tiada dalam eksport | Kesan |
|---|---|
| PBT bagi aduan media sosial | Digantikan dengan pecahan halaman / kumpulan sumber |
| PBT bagi aduan SISPAA | Slaid kategori jadi lebar penuh |
| Ageing (tiada tarikh tutup) | Blok ageing dilangkau; digantikan kad ringkasan |
| DM media sosial | Slaid DM dilangkau sepenuhnya |
| Status aduan pencegahan | Slaid status pencegahan dilangkau |
| Zon bagi aduan pencegahan | Digantikan dengan pecahan agensi |
| Nama ahli majlis | Jadual prestasi zon memaparkan kawasan sahaja. Muat naik helaian pemetaan untuk memaparkan nama. |

Untuk memaparkan blok-blok ini, guna templat Excel dan isi nilainya secara manual.

## Amaran bulan tidak lengkap

Eksport selalunya diambil pertengahan bulan. Kalau bilangan bagi bulan terpilih jauh lebih rendah daripada purata tiga bulan sebelum, sistem memberi amaran — supaya penurunan mendadak dalam carta tidak disalah tafsir sebagai penurunan sebenar.

Sebab itu bulan lalai ialah **bulan penuh terakhir**, bukan bulan terkini.


---

## Seksyen PBT daripada senarai aduan penuh

Eksport senarai aduan penuh (38 lajur) mengisi seluruh seksyen Status Semasa Aduan. Beberapa nilai dikira, bukan diambil terus:

**Lawatan Tapak (QC)** dikira mengikut lajur `Tarikh Lawatan` apabila ada — iaitu bila lawatan benar-benar dilaksana. Jika lajur itu tiada, sistem berundur kepada `Tarikh aduan` dan memberi amaran, kerana kiraan itu mengukur perkara yang berbeza.

**Melepasi SLA** — umur setiap aduan dikira dalam **hari bekerja** (Isnin–Jumaat, tidak mengambil kira cuti umum), dari `Tarikh aduan` hingga `Tarikh Kemaskini` bagi aduan yang telah tamat, atau hingga hari terakhir bulan laporan bagi aduan yang masih aktif. Ambang mengikut `Tahap Kesukaran`: Mudah 3, Sederhana 15, Kompleks 365 hari bekerja.

**Purata hari** dalam jadual prestasi jabatan menggunakan asas yang sama, dikira hanya ke atas aduan yang telah diselesaikan.

**Aduan tanpa jabatan** tidak disenaraikan dalam jadual prestasi jabatan. Ia bukan jabatan yang berprestasi rendah — ia aduan yang belum ditugaskan langsung, dan memasukkannya akan mencemarkan kadar selesai.

**Tajuk aduan tertunggak** menggunakan `Sub Kategori`, bukan `Perincian aduan`. Perincian ialah teks bebas yang boleh mengandungi nama, alamat, atau nombor telefon pengadu — ia tidak sesuai untuk laporan yang diedarkan.

**Nama, telefon, e-mel, alamat dan koordinat pengadu tidak digunakan langsung** dalam mana-mana slaid.

Apabila hanya satu bulan data tersedia bagi sesuatu PBT, carta trend digantikan dengan pecahan sumber aduan, dan perbandingan bulan lepas disembunyikan.


## Pengendalian data tidak lengkap

Beberapa keadaan dikendalikan sendiri supaya carta tidak mengelirukan:

| Keadaan | Tindakan sistem |
|---|---|
| Bulan sifar di hadapan siri trend | Dibuang. Sifar sebelum rekod mula dikumpul bukan sifar sebenar. |
| Hanya satu bulan data bagi sesuatu seksyen | Carta trend digantikan dengan pecahan lain (sumber, kategori, atau jabatan). |
| Tiada bulan sebelum untuk dibandingkan | Teks "vs bulan lepas" disembunyikan, digantikan label tempoh. |
| Hanya satu PBT dalam data | Carta pecahan PBT digantikan dengan pecahan jabatan atau agensi. |
| Tiada nama ahli majlis | Lajur nama digugurkan; jadual memaparkan zon sahaja. |
| Label kategori panjang | Dibungkus kepada dua baris pada paksi carta, bukan dipotong. PowerPoint membungkus sendiri. |
| Hirisan donat terlalu kecil | Peratus hanya dilukis dalam hirisan 6% ke atas; hirisan lebih nipis dibiar kosong kerana teks tidak muat. Nombor penuh setiap kategori ada dalam petunjuk. |
| Zon tanpa aduan bersumber Ahli Majlis | Tetap disenaraikan dengan nilai sifar dan kadar dipaparkan sebagai sempang, supaya senarai zon sentiasa lengkap. |
| Bulan terpilih jauh lebih rendah daripada purata tiga bulan sebelum | Amaran diberikan — kemungkinan eksport diambil pertengahan bulan. |


## Kadar selesai zon ahli majlis

`Kadar = Selesai / (Jumlah - Ditolak)`

Aduan ditolak dikeluarkan daripada penyebut kerana ia tidak mungkin menjadi "selesai". Aduan **Kompleks dikekalkan** dalam kiraan — ia boleh dan memang diselesaikan, dan mengeluarkannya daripada penyebut sahaja menghasilkan kadar melebihi 100%. Lajur Kompleks kekal dipaparkan sebagai konteks, kerana zon dengan banyak aduan Kompleks memang mengambil masa lebih lama.


## Label carta donat

Setiap donat memaparkan dua perkara:

- **Dalam hirisan** — peratus, tetapi hanya bagi hirisan **6% ke atas**. Hirisan lebih nipis dibiar kosong; teks tidak muat dan akan bertindih dengan label jiran.
- **Dalam petunjuk** — nama kategori diikuti bilangan sebenar, contoh `Ditugaskan (17)`.

Jadi hirisan besar boleh dibaca sekilas pandang, dan nilai tepat setiap kategori tetap ada tanpa mengotorkan carta.

Ambang 6% boleh diubah dalam `assets/render-html.js`:

```js
const HAD_PERATUS = 0.06;
```

Fail PPTX meletak nilai dalam petunjuk sahaja, tanpa label hirisan. PowerPoint tidak membenarkan kedudukan atau paparan label ditetapkan per hirisan dari fail, jadi label hirisan nipis akan bertindan pada sesetengah penukar.


## Pemformatan nombor

Setiap nombor yang dipaparkan menggunakan pemisah ribuan: `1155` menjadi `1,155`. Ini terpakai pada kad statistik, corong, bar bersegmen, label carta, petunjuk donat, sel jadual, dan ayat dapatan automatik.

Fungsinya ada dalam `assets/theme.js`:

```js
window.fmt = (n) => ...
```

## Warna status aduan

Setiap status mempunyai warnanya sendiri. Status yang tiada dalam peta `STATUS_WARNA` diberi warna berbeza secara berturutan daripada palet siri, jadi tiada dua status berkongsi warna walaupun sistem MAJ menambah status baharu kemudian.

Fungsi `warnaStatus()` dalam `assets/theme.js` menguruskan ini.

## Tahap kesukaran

Carta tahap kesukaran mengira **aduan yang belum selesai sahaja** — dari Menunggu hingga Selesai Sementara. Aduan Diselesaikan dan Ditolak dikecualikan, kerana tahap kesukaran hanya bermakna bagi kerja yang masih berjalan.

Kategori "Tidak dinyatakan" tidak dipaparkan; ia bukan satu tahap kesukaran sebenar.

Ambil perhatian bahawa kad "Masih aktif" pada slaid lain menggunakan takrifan berbeza — ia tidak mengira Selesai Sementara sebagai aktif. Nota di bawah tajuk carta menyatakan asas kiraan supaya kedua-dua angka tidak disalah banding.
