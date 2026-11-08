# TERA — Teliti · Efisien · Rapi · Akurat

**Smart reporting untuk laporan lab pabrik kelapa sawit.**
Tempel laporan dari WhatsApp → TERA membaca dan memeriksa → mengisi file Excel milik Anda secara otomatis.
Hanya sel yang dituju yang berubah. Rumus, format, dan sheet lain tidak disentuh.

Versi: **v4.6** · Berjalan 100% di browser (tanpa server, data tidak dikirim ke mana pun) · Cocok untuk HP dan komputer · Mode terang/gelap

---

## Daftar isi
1. [Apa yang dilakukan TERA](#1-apa-yang-dilakukan-tera)
2. [Mulai cepat (5 langkah)](#2-mulai-cepat-5-langkah)
3. [Fitur satu per satu](#3-fitur-satu-per-satu)
4. [Cara kerja draf dan unduhan](#4-cara-kerja-draf-dan-unduhan)
5. [Opsi: ubah tanggal di Excel sesuai laporan](#5-opsi-ubah-tanggal-di-excel-sesuai-laporan)
6. [Pemeriksaan dan pengaman](#6-pemeriksaan-dan-pengaman)
7. [Pengaturan](#7-pengaturan)
8. [Menjalankan di komputer](#8-menjalankan-di-komputer)
9. [Publikasi ke Vercel](#9-publikasi-ke-vercel)
10. [Mengubah kode](#10-mengubah-kode)
11. [Penyimpanan data dan privasi](#11-penyimpanan-data-dan-privasi)
12. [Pemecahan masalah](#12-pemecahan-masalah)
13. [Pertanyaan yang sering muncul](#13-pertanyaan-yang-sering-muncul)
14. [Batasan yang perlu diketahui](#14-batasan-yang-perlu-diketahui)
15. [Struktur proyek](#15-struktur-proyek)
16. [Riwayat versi](#16-riwayat-versi)

---

## 1. Apa yang dilakukan TERA

Analis lab biasanya menerima laporan berupa teks WhatsApp, lalu mengetik ulang angkanya ke Excel bulanan. Itu lambat dan mudah salah ketik. TERA mengerjakan bagian mengetik itu dengan aman:

1. **Membaca** teks laporan (tanggal, shift, nama titik, angka).
2. **Memeriksa** (tanggal vs hari, rentang angka, ODM vs OWM, duplikat, bulan file, dan lainnya).
3. **Menampilkan pratinjau** sel mana yang akan diisi, dilewati, atau diblokir, sebelum ada yang ditulis.
4. **Mengisi** hanya sel tujuan di file Excel Anda, lalu **membaca ulang** hasilnya untuk memastikan benar.
5. Menyimpan catatan di **Riwayat** dan mengunduh Excel **sekali** setelah semua tanggal selesai.

### Enam menu

| Menu | Fungsi | Tujuan di Excel |
|---|---|---|
| **Oil Losses** | Laporan Moist / OWM / ODM per titik, per shift | Sheet harian `01`–`31`. Shift I mulai baris 78, Shift II mulai baris 144 |
| **Kualitas** | Laporan kualitas pagi (CPO, OST, COT) | Halaman 1 sheet harian, hanya kolom-kolom yang dituju |
| **Limbah** | pH dan suhu effluent Kolam 1–16 | Sheet `Harian Limbah <Bulan> <Tahun>`, baris sesuai tanggal |
| **Kernel** | Rata-rata sampling kernel per shift | Hanya ditampilkan di layar (isi ke Excel menyusul) |
| **Riwayat** | Catatan setiap isian + status "Belum diunduh" | — |
| **Pengaturan** | Batas nilai, nama alias titik, cadangan | — |

Fitur Absen dan Serah Terima sudah dihapus. Tidak ada tombol "isi contoh".

### Dua file Excel yang terpisah
TERA punya dua tempat file:
- **File losses**: dipakai oleh **Oil Losses** dan **Kualitas** (workbook dengan sheet `01`–`31`).
- **File limbah**: dipakai oleh **Limbah** (workbook dengan sheet `Harian Limbah …`).

Keduanya bisa dimuat bersamaan. Jenis file dikenali otomatis: jika Anda salah memasukkan file ke tempat yang salah, TERA memindahkannya ke tempat yang benar dan memberi tahu Anda.

---

## 2. Mulai cepat (5 langkah)

1. **Pilih file Excel** (tombol *Pilih file Excel* atau seret file ke kotak). Hanya `.xlsx`.
2. **Tempel laporan** dari WhatsApp ke kotak teks (atau tekan *Tempel*).
3. **Periksa pratinjau.** Hijau berarti siap, kuning berarti perlu dicek, merah berarti diblokir sampai diperbaiki.
4. Tekan **Isi ke draf** (atau `Ctrl + Enter`). Ulangi untuk tanggal atau laporan berikutnya.
5. Setelah selesai, tekan **Unduh Excel** satu kali. Buka file hasilnya di Excel agar rekap terhitung ulang.

Anda boleh berhenti di tengah bulan (misalnya sampai tanggal 14) lalu mengunduh. Tidak perlu menunggu bulan penuh.

---

## 3. Fitur satu per satu

### 3.1 Oil Losses
- **Contoh format laporan:**
  ```
  Lap Oil Losses
  Rabu, 04.11.2026
  Shift I

  Line I
  Fibre Press 1
  Moist = 32.03
  Owm = 3.84
  Odm = 5.65
  ```
- TERA membaca tanggal, shift (I atau II), nama titik, dan angka. Nama titik yang tidak dikenal ditampilkan sebagai "tidak dikenali"; Anda bisa memetakannya sekali ke titik yang benar (tersimpan sebagai alias).
- **Tujuan:** sheet sesuai tanggal (tanggal 4 → sheet `04`). Shift I → baris dasar 78, Shift II → baris dasar 144. Di tiap titik: Moist = baris dasar, OWM = +1, ODM = +2.
- **22 titik:** Fibre Press 1–6, Nut Line I/II, Before BP, EBP 1–2, Heavy Phase Centrifuge 1–5, Crude EBP, Underflow Juice Tank, Condensat 1–2, Inlet Recovery, Final Effluent.
- Titik yang tidak ada di laporan tidak diubah di Excel dan ditampilkan sebagai "belum ada di laporan".

### 3.2 Kualitas
Laporan kualitas produksi pagi. Bagian yang dibaca:

| Bagian laporan | Tujuan di sheet harian |
|---|---|
| CPO Produksi (Moist, Dirt, FFA) | Kolom **P** baris 50 / 52 / 53 |
| OST 1, 2, 3, 4, D, E | Kolom **R, S, T, U, V, W** baris 50 / 52 / 53 |
| Komposisi COT (FFA, Oil, Emulsi, Moist, Dirt) | Sel **O19, P19, Q19, R19, S19** |
| Kualitas Kernel Produksi, Losses Kernel, Moisture Bottom | Dibaca dan ditampilkan, **belum ditulis** (letaknya di lembar shift; menunggu fitur Kernel Losses) |

Halaman 1 banyak berisi rumus. TERA hanya menulis sel kolom yang dituju dan **tidak pernah menulis sel berumus**. Komposisi COT diperiksa apakah jumlahnya sekitar 100%.

### 3.3 Limbah
- **Contoh format laporan:**
  ```
  Laporan PH&Suhu.Effluent Sabtu, 7-11-2026
  Kolam NO1=4.47 Suhu:41.2
  Kolam NO2=4.43 Suhu:40.9
  …
  Kolam NO16=8.47 Suhu:27.3
  ```
- Sheet dipilih otomatis dari **bulan dan tahun laporan** (mis. `Harian Limbah November 2026`). Nama sheet dicocokkan secara fleksibel (spasi berlebih, singkatan bulan). Bila ada beberapa sheet bernama mirip, dipilih yang tanggalnya cocok.
- **Tata letak dideteksi per sheet** dari tulisan `KOLAM n`, sub-judul `PH` / `SUHU`, dan header `Tanggal`. Jadi bulan yang kolomnya bergeser tetap terbaca benar. Baris tanggal juga **dicek** terhadap tanggal di kolom tanggal; jika tidak sama, pengisian dihentikan agar tidak salah hari.
- Kolam yang tidak ada di laporan dilewati dan sel Excelnya tidak diubah.

### 3.4 Kernel
Menghitung rata-rata sampling kernel dari chat WhatsApp (format seperti `[4/11, 11.27] Nama: …` atau `4/11/26, 11.27 - Nama: …`, dengan baris `Sampling jam 11.27`).
- Rata-rata dikelompokkan **per shift**. Jam mulai shift bawaan **07:00** (Shift I) dan **19:00** (Shift II); bisa diubah di pengaturan Kernel.
- Setiap baris sampling bisa dikecualikan dari rata-rata dengan satu ketukan.
- Hasilnya hanya ditampilkan. **Belum ada pengisian ke Excel** (Kernel Losses menyusul).

### 3.5 Riwayat
- Mencatat setiap isian: tanggal, jenis (Shift / Kualitas / Limbah), jumlah sel, sheet, dan nama file.
- Isian yang **belum diunduh** diberi tanda kuning **Belum diunduh**. Cocok untuk mengecek tanggal mana yang terlewat sebelum mengunduh.
- Disimpan di browser perangkat itu (maksimal 500 catatan terakhir).

---

## 4. Cara kerja draf dan unduhan

Sebelumnya setiap isian langsung mengunduh file, sehingga banyak file menumpuk dan lambat. Sekarang:

- **Isi ke draf**: perubahan diterapkan di memori, lalu **diverifikasi** dengan membuka ulang hasilnya. **Tidak ada unduhan.**
- Semua isian menumpuk di **draf**. Kartu file menampilkan daftarnya, misalnya `1 Nov · Shift I`, `1 Nov · Kualitas`, `7 Nov · Limbah`, beserta jumlahnya.
- **Unduh Excel** menulis semuanya **sekali**.
- **Batalkan terakhir** membatalkan isian terakhir di draf (sampai 10 langkah). File di komputer tidak tersentuh.
- **Draf tersimpan di perangkat.** Jika tab tertutup atau halaman dimuat ulang, draf dipulihkan dan Anda diberi tahu.
- Menutup tab saat draf belum diunduh memunculkan peringatan browser. Mengganti file saat draf belum diunduh meminta konfirmasi.

### Nama file unduhan
Nama dibuat rapi dan mudah dibedakan:

```
UPDATE LAPORAN BY TERA (Oil Losses s.d. 04 Nov 2026).xlsx
UPDATE LAPORAN BY TERA (Oil Losses + Kualitas s.d. 14 Nov 2026).xlsx
UPDATE LAPORAN BY TERA (Limbah s.d. 07 Nov 2026).xlsx
```
Isi kurung = jenis laporan yang masuk + tanggal laporan terakhir. Jika Anda mengunduh dua kali dengan isi yang sama, unduhan kedua diberi akhiran `- ke-2`.

### Simpan langsung (opsional, Chrome/Edge)
Tombol **Simpan langsung** menyambungkan TERA ke file di komputer, sehingga tombol unduh menjadi **Simpan ke file** dan file ditimpa di tempat (nama tidak berubah). Jika file di komputer sudah berubah sejak dibuka (misalnya Anda menyimpannya dari Excel), TERA **tidak menimpanya** dan mengunduh file baru agar pekerjaan Anda tidak hilang. Fitur ini hanya tambahan; cara biasa lewat *Pilih file Excel* adalah yang paling andal.

---

## 5. Opsi: ubah tanggal di Excel sesuai laporan

Tersedia di **Oil Losses** dan **Kualitas**. Bawaannya **mati** (centang manual).

**Masalahnya:** header setiap sheet harian punya bagian `DATE : hari / bulan / tahun`. Saat bulan baru, bagian ini sering lupa diganti sehingga masih tertulis bulan atau tahun lama (mis. `oktober`).

**Yang dilakukan opsi ini:** menyamakan hari, bulan, dan tahun di header Excel dengan tanggal di laporan yang Anda tempel.

Contoh: laporan *Rabu, 04.11.2026* ditempel ke file yang header sheet `04`-nya masih `oktober`.
- Header sheet `04`: `4 / oktober / 2026` → `4 / november / 2026`
- Sheet `01` (induk bulan dan tahun bagi sheet lain) ikut disamakan.

Aturan keamanannya:
- Sel **berumus tidak disentuh** (mengikuti sheet `01`).
- Gaya sel dipertahankan; huruf teks bulan mengikuti gaya yang ada (`oktober` → `november`, `Oktober` → `November`, `OKTOBER` → `NOVEMBER`).
- Perubahan **ditampilkan lebih dulu** di pratinjau. Jika bulan induk di sheet `01` akan berubah, TERA meminta konfirmasi agar Anda tidak sengaja memasukkan laporan ke file bulan yang salah.
- Peringatan bulan (sel AC8) tidak muncul lagi saat opsi ini aktif karena sudah ditangani.
- Opsi ini tidak ada di **Limbah**, karena sheet dan barisnya sudah dipilih langsung dari tanggal laporan.

---

## 6. Pemeriksaan dan pengaman

| Pengaman | Penjelasan |
|---|---|
| **Hanya sel target** | Penulisan dilakukan langsung di XML file `.xlsx`. Hanya sel yang dituju diganti; sheet lain, rumus, gaya, dan gambar tidak diubah. Excel diminta menghitung ulang rumus saat file dibuka. |
| **Rumus tak pernah ditulis** | Sel berumus ditandai *diblokir* di pratinjau. |
| **Sel sudah berisi dilewati** | Jika sel sudah berisi angka berbeda, sel itu **dilewati**, kecuali Anda mengaktifkan *Timpa* di Pengaturan. |
| **Input dua kali aman** | Laporan yang sama ditempel lagi akan dikenali sebagai *sama* atau *dilewati*, bukan ditulis dobel. |
| **Tanggal dan hari** | Hari yang tertulis (mis. "Rabu") dicek terhadap tanggalnya. Tanggal masa depan atau terlalu lampau diberi peringatan. |
| **Bulan file** | Bulan di file dibandingkan dengan bulan laporan. |
| **Rentang angka** | Persen di atas 100 atau angka negatif diblokir; batas khusus bisa diatur per titik. |
| **Cek silang ODM** | ODM harus ≈ OWM ÷ (1 − Moist/100). Selisih diberi peringatan "salah ketik?" (hanya peringatan, tidak memblokir). |
| **Duplikat dan konflik** | Titik yang muncul dua kali dengan angka berbeda ditandai. |
| **Lonjakan** | Perubahan besar dibanding hari sebelumnya diberi peringatan. |
| **Verifikasi pasca-tulis** | Setelah menulis, hasil dibuka ulang dan setiap sel dibaca kembali. Jika ada yang tidak sama, pengisian dibatalkan. |
| **Konfirmasi** | Hal yang berisiko (mis. menimpa sel berisi, bulan induk berubah) meminta konfirmasi lebih dulu. |

Arti status sel di pratinjau: **isi** (akan ditulis), **timpa** (akan menggantikan angka lama), **sama** (sudah berisi angka itu), **lewati** (sudah berisi angka lain), **blokir** (rumus/sel gabungan/alamat tidak cocok), **menunggu** (belum ada data).

---

## 7. Pengaturan

| Pengaturan | Bawaan | Fungsi |
|---|---|---|
| **Cek silang ODM** | Aktif | Peringatan jika ODM tidak cocok dengan OWM ÷ (1 − Moist). |
| **Hitung nilai ketiga** | Mati | Jika hanya dua dari Moist/OWM/ODM ada, hitung yang ketiga (ditandai sebagai hasil hitung). |
| **Timpa sel berisi angka** | Mati | Mengganti angka lama. Bawaan mati karena paling aman. |
| **Batas nilai** | Kosong | Batas umum dan batas per titik untuk Moist/OWM/ODM. |
| **Alias titik** | — | Nama titik versi analis yang dipetakan ke titik yang benar (dibuat dari pratinjau). |
| **Tema** | Otomatis | Otomatis / Terang / Gelap. |
| **Cadangan** | — | Simpan/pulihkan pengaturan, alias, batas, dan riwayat ke file JSON. |

---

## 8. Menjalankan di komputer

### Paling mudah: tanpa install
Klik dua kali `siap-pakai/index.html`. Aplikasi langsung terbuka di browser.

### Dengan Live Server di VS Code
1. Buka folder proyek di VS Code.
2. Pasang ekstensi **Live Server** (Ritwick Dey): `Ctrl+Shift+X` → cari "Live Server" → Install.
3. Buka folder `siap-pakai/`, klik kanan `index.html` → **Open with Live Server**.
4. Browser terbuka di `http://127.0.0.1:5500`.

Gunakan **Chrome atau Edge** bila ingin memakai *Simpan langsung*. Untuk cara biasa, browser modern apa pun cukup.

---

## 9. Publikasi ke Vercel

TERA adalah situs statis, jadi tidak perlu build di Vercel.

1. Siapkan folder yang berisi **tepat dua file di level paling atas**: `index.html` dan `vercel.json` (isi folder `siap-pakai/`). Jika ZIP, jangan membungkusnya dalam subfolder.
2. Di Vercel: **Add New → Project**, unggah folder itu.
3. **Framework Preset: Other.** Kosongkan *Build Command* dan *Output Directory*.
4. Deploy. Setelah selesai, lihat teks versi di sidebar kiri bawah (mis. `v4.6`) untuk memastikan versi terbaru yang tampil.

`vercel.json` sudah mengatur agar browser selalu mengambil versi terbaru (`no-cache`) dan menambah header keamanan dasar.

**Jika masih terlihat versi lama:** tekan `Ctrl + Shift + R`, atau buka di jendela penyamaran. Pastikan deployment yang aktif adalah yang terbaru (tab *Deployments* di Vercel) dan `index.html` ada di akar, bukan di dalam subfolder.

---

## 10. Mengubah kode

Butuh Node.js 18 atau lebih baru.

```bash
npm install        # sekali saja
npm run build      # menghasilkan dist/index.html (satu file, siap dipakai offline)
npm test           # menjalankan tes otomatis
```
Setelah build, salin `dist/index.html` ke `siap-pakai/index.html`.

- **Tumpukan teknologi:** React 19, Tailwind CSS v4, esbuild, JSZip (disertakan di `vendor/`).
- **Logika inti** ada di `src/core/` sebagai fungsi murni yang teruji (`node --test`), terpisah dari tampilan.
- **Jika template Excel berubah** (baris/kolom/sheet bergeser), ubah `src/core/config.js`: baris dasar shift, kolom titik, sel tanggal. Untuk Kualitas lihat `src/core/quality.js`; untuk Limbah tata letak dideteksi otomatis.
- Tes mencakup parser, perencana pengisian, penulisan XLSX, Kualitas, Limbah, Kernel, nama file unduhan, dan perubahan tanggal header.

---

## 11. Penyimpanan data dan privasi

- **Tidak ada server.** File Excel dan teks laporan diproses di browser Anda dan tidak dikirim ke mana pun.
- Yang disimpan di browser perangkat ini (localStorage/IndexedDB): pengaturan, batas, alias, riwayat, teks laporan yang belum diisi, dan **draf isian** beserta salinan file yang sedang dikerjakan.
- Data **tidak** berpindah antar perangkat atau antar browser. Membersihkan data situs di browser akan menghapusnya. Gunakan **Cadangan** di Pengaturan untuk menyimpan pengaturan dan riwayat.
- Selama masih ada draf yang belum diunduh, jangan membersihkan data browser.

---

## 12. Pemecahan masalah

| Gejala | Penyebab dan solusi |
|---|---|
| Tombol *Pilih file Excel* tidak bereaksi | Pakai tombol **Pilih file Excel** (pemilih bawaan browser) atau seret file ke kotak. *Simpan langsung* hanya tambahan dan tidak berfungsi di semua browser. |
| "Sheet “04” tidak ada" | File yang dimuat bukan file losses bulan itu. TERA biasanya memindahkan file ke tempat yang benar secara otomatis; jika tidak, pastikan Anda memilih workbook dengan sheet `01`–`31`. |
| Peringatan merah soal bulan (AC8) | Header sheet `01` masih bulan lama. Ubah AC8 jadi tanggal 1 bulan itu, atau centang **Ubah tanggal di Excel sesuai laporan**. |
| "ODM … tidak cocok dengan hitungan OWM ÷ (1 − Moist)" | Peringatan salah ketik. Cek angka laporan; jika memang benar, lanjutkan. Tidak memblokir. |
| Sel dilewati padahal ingin diganti | Sel sudah berisi angka lain. Aktifkan **Timpa sel yang sudah berisi angka** di Pengaturan. |
| "Tanggal tidak terbaca" | Sertakan tanggal seperti `04.11.2026`, `4-11-2026`, atau `04 November 2026`. |
| "Ada lebih dari satu tanggal / shift" | Tempel **satu** laporan saja per kali isi. |
| File `.xls` atau Google Sheets | Hanya `.xlsx`. Dari Google Sheets: *File → Unduh → Microsoft Excel (.xlsx)*. File `.xls` lama: buka di Excel lalu *Simpan Sebagai .xlsx*. |
| Angka rekap tidak berubah | Buka file hasil di Excel; rumus dihitung ulang saat dibuka. Beberapa penampil (pratinjau email/HP) menampilkan nilai lama. |
| Draf hilang | Draf ada di browser dan perangkat yang dipakai mengisi. Jika data situs dibersihkan, draf ikut hilang. |
| Vercel masih versi lama | Lihat bagian 9. |

---

## 13. Pertanyaan yang sering muncul

**Apakah rumus Excel saya aman?**
Ya. Rumus tidak pernah ditulis. Pengisian hanya mengubah sel tujuan, dan hasilnya diperiksa ulang. Pada pengujian dengan template nyata, perbandingan sebelum dan sesudah hanya menunjukkan perubahan di sel yang dituju.

**Kalau laporan yang sama ditempel dua kali?**
Aman. Sel yang sudah berisi angka sama dikenali sebagai *sama*; yang berbeda dilewati (kecuali *Timpa* aktif). Tidak ada penulisan ganda.

**Boleh berhenti di tanggal 14 lalu mengunduh?**
Boleh. Unduh kapan saja ada isian di draf. Nama file menunjukkan tanggal terakhir (`s.d. 14 Nov 2026`). Lanjutkan dari tanggal 15 dan unduh lagi nanti.

**Bagaimana mengetahui tanggal mana yang belum diisi?**
Lihat kartu file (daftar isian di draf) dan menu **Riwayat** (tanda *Belum diunduh*).

**Apakah harus mengunggah dua file?**
Tidak. Muat file yang dibutuhkan fitur yang sedang dipakai. Oil Losses dan Kualitas memakai file losses; Limbah memakai file limbah.

**Riwayat hilang kalau saya menutup browser?**
Tidak. Riwayat tersimpan di perangkat sampai Anda menghapus data browser atau menghapusnya sendiri.

**Apakah bisa dipakai di HP?**
Bisa. Tampilan responsif dengan navigasi bawah. File Excel dipilih dari penyimpanan HP.

---

## 14. Batasan yang perlu diketahui

- Isi ke Excel untuk **Kernel** ("Kernel Losses") dan penulisan bagian Kernel/Moisture Bottom di Kualitas **belum tersedia**; saat ini hanya dibaca dan ditampilkan.
- Jam mulai shift Kernel (07:00 / 19:00) adalah asumsi bawaan; ubah di pengaturan Kernel jika berbeda.
- Draf, riwayat, dan pengaturan tersimpan **per browser per perangkat**.
- Format teks laporan yang benar-benar baru mungkin belum dikenali. Baris seperti itu ditampilkan sebagai "tidak dikenali" dan tidak ditulis. Kirim contohnya agar parser dapat disesuaikan.
- Template yang strukturnya berubah jauh (baris shift, kolom titik) memerlukan penyesuaian di `src/core/config.js`.
- Hanya `.xlsx`.
- Pastikan sel AC8 di sheet `01` mengikuti bulan berjalan (atau pakai opsi ubah tanggal di Excel).

---

## 15. Struktur proyek

```
tera/
├─ siap-pakai/            ← yang diunggah ke Vercel / dibuka langsung
│  ├─ index.html          (aplikasi satu file)
│  └─ vercel.json
├─ src/
│  ├─ core/               logika murni dan teruji
│  │  ├─ parser.js        membaca laporan Oil Losses
│  │  ├─ validate.js      pemeriksaan tanggal, rentang, ODM
│  │  ├─ planner.js       rencana pengisian Oil Losses
│  │  ├─ quality.js       laporan Kualitas
│  │  ├─ effluent.js      laporan Limbah (deteksi tata letak)
│  │  ├─ kernel.js        rata-rata sampling Kernel
│  │  ├─ headerdate.js    opsi ubah tanggal di Excel
│  │  ├─ draft.js         nama file unduhan
│  │  ├─ kind.js          mengenali jenis file (losses / limbah)
│  │  ├─ xlsx.js          penyunting .xlsx langsung di XML
│  │  └─ config.js        semua asumsi template Excel
│  ├─ app/                tampilan React (halaman, komponen, state)
│  ├─ lib/                penyimpanan dan akses file
│  └─ styles.css
├─ tests/                 tes otomatis + contoh template
├─ scripts/build.mjs      membuat satu file HTML
├─ vendor/jszip.min.js
├─ PRD.md                 dokumen kebutuhan produk
└─ README.md
```

Dokumen kebutuhan produk lengkap ada di [`PRD.md`](PRD.md).

---

## 16. Riwayat versi

| Versi | Perubahan utama |
|---|---|
| **v4.6** | Opsi *Ubah tanggal di Excel sesuai laporan*; nama file unduhan rapi `UPDATE LAPORAN BY TERA (…)`. |
| **v4.5** | Draf: isi banyak tanggal, unduh sekali; draf tersimpan di perangkat; tanda *Belum diunduh* di Riwayat; Batalkan sampai 10 langkah. |
| v4.4 | Auto-routing file losses/limbah; perbaikan pemilih file dan peringatan bulan. |
| v4.x | Fitur Kualitas, Limbah, dan Kernel; tampilan baru; dua slot file. |
| v3 | Oil Losses dan Kernel dasar. |
