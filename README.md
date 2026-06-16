# 🗂️ Claude File Generator

**Claude File Generator** adalah aplikasi web yang memungkinkan kamu mengeksekusi kode JavaScript dan mengunduh hasilnya sebagai file — langsung dari browser, tanpa server, tanpa instalasi.

Cukup paste kode JS yang dihasilkan Claude AI, klik Generate, dan file-mu siap diunduh. Format output bisa apa saja: `.docx`, `.drawio`, `.csv`, `.html`, `.xml`, dan lainnya — selama kode JS-mu menghasilkannya.

🌐 **Akses publik:** [claudedocx.prakpksi.my.id](http://claudedocx.prakpksi.my.id)
📦 **Repositori:** [github.com/kurokana/Claude-File-Generator](https://github.com/kurokana/Claude-File-Generator)

---

## ✨ Fitur Utama

- **Multi-format output** — generate file `.docx`, `.drawio`, `.csv`, `.html`, `.xml`, dan format lainnya
- **100% browser-based** — tidak perlu Node.js, backend, atau instalasi apapun
- **Library `docx` terintegrasi** — siap pakai tanpa setup tambahan untuk generate Word document
- **Eksekusi kode JS langsung** — paste dan jalankan kode dari Claude AI dalam satu klik
- **Console output real-time** — tampilkan log dan error langsung di halaman
- **Tidak perlu login** — gratis dan terbuka untuk umum

---

## 🚀 Cara Menggunakan

### Langkah 1 — Buka Website

```
http://claudedocx.prakpksi.my.id
```

---

### Langkah 2 — Minta Claude Membuat Kode

Buka [claude.ai](https://claude.ai) dan minta Claude untuk membuat kode JavaScript yang menghasilkan file. Sesuaikan prompt dengan format yang kamu inginkan.

**Contoh prompt untuk file `.docx`:**
> *"Buatkan kode JavaScript menggunakan library `docx` (CommonJS / `require`) untuk membuat file Word berisi laporan dengan tabel dan heading. Gunakan `Packer.toBuffer` dan `fs.writeFileSync` untuk menyimpannya."*

**Contoh prompt untuk file `.drawio`:**
> *"Buatkan kode JavaScript yang menghasilkan file `.drawio` berisi diagram flowchart sederhana, lalu simpan dengan `fs.writeFileSync`."*

**Contoh prompt untuk file `.csv`:**
> *"Buatkan kode JavaScript yang membuat file CSV berisi data penjualan bulanan, lalu simpan dengan `fs.writeFileSync('data.csv', ...)`."*

---

### Langkah 3 — Paste Kode ke Website

1. Salin seluruh kode JavaScript dari Claude
2. Paste ke kotak editor berlabel **"Paste your JS code here"**
3. Jumlah baris kode akan otomatis terhitung di bagian bawah editor

---

### Langkah 4 — Generate & Download

1. Klik tombol **"Generate"**
2. Website akan mengeksekusi kode JS di browser
3. File akan **otomatis terunduh** ke perangkatmu
4. Jika ada error, cek panel **Console** di bagian bawah untuk detailnya

---

### Langkah 5 — Cek Console (jika ada masalah)

Panel Console menampilkan log eksekusi secara real-time. Jika file gagal di-generate, pesan error di sini biasanya cukup untuk mendiagnosis masalahnya — misalnya syntax error atau library yang tidak tersedia.

Klik **"Clear"** untuk membersihkan log.

---

## 📁 Format File yang Didukung

Website ini bisa menghasilkan format file apapun yang bisa dibuat lewat kode JavaScript. Berikut status format yang sudah dikonfirmasi:

| Format | Status | Keterangan |
|---|---|---|
| `.docx` | ✅ Confirmed | Menggunakan library `docx` yang sudah terintegrasi |
| `.drawio` | ✅ Confirmed | Generate XML diagram untuk draw.io / diagrams.net |
| `.csv` | 🔲 Belum dites | Bisa dibuat via string manipulation biasa |
| `.html` | 🔲 Belum dites | Bisa dibuat via string template |
| `.xml` | 🔲 Belum dites | Bisa dibuat via string atau DOM serialization |
| `.json` | 🔲 Belum dites | Bisa dibuat via `JSON.stringify` |
| `.md` | 🔲 Belum dites | Bisa dibuat via string manipulation |

> Format lain bisa jadi juga berjalan — selama kode JS-nya bisa membentuk konten file dan menentukan nama file output-nya.

---

## 💡 Tips & Catatan Penting

### Gunakan CommonJS, bukan ES Module
Kode harus menggunakan `require(...)`, bukan `import`. Claude terkadang menghasilkan ES Module syntax — pastikan kamu memintanya menggunakan CommonJS.

```javascript
// ✅ Benar
const { Document, Packer } = require('docx');

// ❌ Akan error di website ini
import { Document, Packer } from 'docx';
```

### `fs.writeFileSync` tidak benar-benar menulis ke disk
Karena berjalan di browser, `fs.writeFileSync('nama-file.xyz', konten)` akan **otomatis diinterpretasi sebagai trigger download**. Kamu tidak perlu mengubah kode — cukup pastikan nama file dan ekstensinya benar.

### Library yang tersedia
Saat ini library yang sudah di-bundle dan siap pakai adalah **`docx`**. Untuk format lain yang tidak butuh library eksternal (seperti `.csv`, `.drawio`, `.html`, `.json`), kamu bisa membuat kontennya murni dengan JavaScript biasa.

### Contoh kode untuk format non-docx

**`.drawio` (XML diagram):**
```javascript
const xml = `<mxfile><diagram><mxGraphModel><root>
  <mxCell id="0"/><mxCell id="1" parent="0"/>
  <mxCell id="2" value="Start" style="ellipse;" vertex="1" parent="1">
    <mxGeometry x="100" y="100" width="120" height="60" as="geometry"/>
  </mxCell>
</root></mxGraphModel></diagram></mxfile>`;

const fs = require('fs');
fs.writeFileSync('diagram.drawio', xml);
```

**`.csv`:**
```javascript
const data = [
  ['Nama', 'Bulan', 'Penjualan'],
  ['Produk A', 'Januari', '1500000'],
  ['Produk B', 'Januari', '2300000'],
];

const csv = data.map(row => row.join(',')).join('\n');
const fs = require('fs');
fs.writeFileSync('data.csv', csv);
```

---

## 🗂️ Struktur Repositori

```
Claude-File-Generator/
├── index.html       # Halaman utama website
├── style.css        # Styling / tampilan UI
├── app.js           # Logika utama: eksekusi kode & trigger download
├── docx.iife.js     # Library docx (versi browser / IIFE bundle)
├── test.html        # Halaman pengujian
└── CNAME            # Konfigurasi domain kustom (GitHub Pages)
```

---

## 🛠️ Menjalankan Secara Lokal

Website ini statis murni — tidak ada build step atau dependency yang perlu diinstall.

```bash
# Clone repositori
git clone https://github.com/kurokana/Claude-File-Generator.git

cd Claude-File-Generator

# Buka index.html di browser
# Atau gunakan Live Server (VS Code) untuk menghindari masalah CORS lokal
```

> **Tips:** Beberapa browser memblokir eksekusi JS dari file lokal (`file://`). Gunakan ekstensi **Live Server** di VS Code, atau akses langsung versi publiknya.

---

## 🤝 Kontribusi

Pull request terbuka! Jika kamu ingin menambah dukungan library baru, memperbaiki bug, atau meningkatkan UI:

1. Fork repositori ini
2. Buat branch baru: `git checkout -b fitur-baru`
3. Commit: `git commit -m 'Tambah fitur X'`
4. Push: `git push origin fitur-baru`
5. Buat Pull Request

---

## 👤 Pembuat

Dibuat oleh **[Kurokana Unila](https://github.com/kurokana)** — didukung oleh **[Harutcha](https://saweria.co/harutcha)**.

---

## 📜 Lisensi

Proyek ini bersifat open source. Silakan digunakan, dimodifikasi, dan didistribusikan dengan menyertakan kredit kepada pembuat aslinya.
