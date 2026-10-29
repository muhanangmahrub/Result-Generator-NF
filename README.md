# Result Test Generator

Aplikasi desktop Windows untuk QA Engineer: membuat dokumen **Result Test UAT (Word)**
dari file **skenario test (Excel)**, lengkap dengan screenshot hasil eksekusi per test case.

> **Status: Tahap 1-4 dan fase 2 opsional selesai.**
> Parser Excel, project + GUI, screenshot + global hotkey, generate dokumen Word,
> tulis balik hasil ke salinan Excel, dan export PDF.

---

## Instalasi

Butuh **Python 3.11+** di Windows.

```powershell
cd F:\2026\Result-Generator-NF
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Untuk Tahap 1 saja cukup `pip install openpyxl pytest`.

## Menjalankan aplikasi

```powershell
python app.py

# atau langsung membuka sebuah project
python app.py "D:\UAT\AMF Segmentation - Result"
```

### Alur pemakaian

1. **File > Project baru** — pilih file skenario `.xlsx`, lalu folder project.
2. Muncul **daftar sheet dengan checkbox**. Sheet yang punya kolom wajib ditandai
   *sheet skenario* dan langsung tercentang; sheet rekap seperti `Main` atau
   `ZIP Group` tidak bisa dicentang. Sheet bertanda `?` bisa diperbaiki lewat
   **Mapping kolom manual...** (dialog ini bisa dibuka untuk sheet apa pun, dan hasilnya
   disimpan ke `project.json`).
3. Tree di kiri menampilkan **Module → Scenario → Test Case → Test Step** beserta
   **jumlah screenshot per test case**. Kotak *Cari* memfilter test case berdasarkan nama
   maupun isi langkahnya.
4. Tekan **Ctrl+Shift+S** (bisa diubah) di mana pun — walau aplikasi sedang diminimize —
   untuk mengambil screenshot. Muncul popup berisi preview, pilihan test case, dan caption.
5. Tab **Screenshot** di kanan: thumbnail test case terpilih. Urutannya bisa diubah dengan
   **drag**, dan ada tombol *Tambah dari file...*, *Caption...*, *Pindah test case...*,
   *Hapus*. Urutan inilah yang dipakai di dokumen Word.
6. **Project > Metadata & judul module** — ubah metadata dan judul `Heading 1` tiap module.
   Kolom yang dikosongkan kembali memakai nilai dari Excel (ditampilkan sebagai placeholder).
7. **F5 / File > Re-import Excel** — baca ulang skenario tanpa menghapus screenshot.
8. Tab **Hasil** di kanan: isi *Actual Results*, *Pass/Fail*, *Date Executed* dan
   *Test Executed by* per step. Ada tombol *Isi semua step...* dan
   *Tandai semua Pass hari ini*.
9. **Ctrl+E / Export > Tulis hasil ke salinan Excel** — hasil tadi ditulis ke **salinan**
   file skenario. File aslinya tidak pernah disentuh.
10. **Ctrl+G / Export > Generate dokumen Word** — pilih per module atau gabungan, lihat
   ringkasan test case yang belum ada screenshot-nya, lalu generate ke folder `output/`.

11. **Export > Export PDF** — semua `.docx` di folder `output/` diubah jadi PDF.

Tab **Peringatan** mengumpulkan peringatan parser, file screenshot yang hilang, jumlah
screenshot orphan, dan hasil eksekusi yang test case-nya sudah tidak ada di Excel.

## Tampilan

Gaya dasarnya Fusion dengan tema sendiri (`rtg/gui/theme.py`), terang atau gelap mengikuti
pengaturan Windows dan berganti otomatis tanpa restart. Semua warna berasal dari satu set
token di file itu; stylesheet-nya dibangun dari token tersebut.

- Toolbar berisi ikon saja (tooltip memuat nama lengkap dan shortcut), kecuali **Capture**
  dan **Generate** di kanan.
- Sidebar tanpa header: ikon centang hijau = sudah ada screenshot, lingkaran kosong = belum,
  segitiga oranye = skenario berubah dan perlu dicek ulang. Progres test case yang terisi
  tampil di status bar.
- Ikon memakai set Phosphor lewat `qtawesome` (`rtg/gui/icons.py`). Kalau qtawesome tidak
  terpasang, aplikasi tetap jalan tanpa ikon.
- Ukuran jendela dan posisi pemisah panel diingat antar sesi.

## Menjalankan lewat CLI

Parser dan generator juga bisa dipakai tanpa GUI:

```powershell
# Lihat sheet mana yang terdeteksi sebagai sheet skenario + mapping kolomnya
python -m rtg.cli inspect "D:\path\Scenario Test UAT.xlsx"

# Baca skenario dan tampilkan tree Module -> Scenario -> Test Case -> Test Step
python -m rtg.cli parse "D:\path\Scenario Test UAT.xlsx"

# Hanya module tertentu (boleh diulang), dan simpan hasilnya ke JSON
python -m rtg.cli parse "D:\path\Scenario Test UAT.xlsx" --sheet "AMF Segmentation" --json out.json

# Buat dokumen Result Test UAT (.docx)
python -m rtg.cli generate "D:\path\Scenario Test UAT.xlsx" --out output

# Salin file config alias kolom supaya bisa diedit
python -m rtg.cli write-config config
```

### Generate dokumen Word

```powershell
# Ringkasan per module dulu, tanpa menulis file apa pun
python -m rtg.cli generate "Scenario.xlsx" --dry-run

# Satu file .docx per module (default)
python -m rtg.cli generate "Scenario.xlsx" --out output

# Satu file gabungan untuk semua module terpilih
python -m rtg.cli generate "Scenario.xlsx" --out output --combined

# Hanya module tertentu
python -m rtg.cli generate "Scenario.xlsx" --out output --sheet "AMF Segmentation"

# Pakai .docx milik sendiri sebagai template style (font, header/footer, logo)
python -m rtg.cli generate "Scenario.xlsx" --out output --template "Result Test UAT AMF Segmentation.docx"

# Sisipkan screenshot dari file mapping JSON
python -m rtg.cli generate "Scenario.xlsx" --out output --screenshots shots.json
```

Format `shots.json` — kunci-nya `TestCase.uid` (`"<nama sheet>::<Test Case ID>"`), nilainya
daftar screenshot berurutan. File ini dihasilkan otomatis oleh project (lihat `Project.export_screenshots()`), atau
bisa ditulis manual untuk mengetes:

```json
{
  "Validation AMF Segmentation::1.1": [
    "shots/1-1_1.png",
    { "path": "shots/1-1_2.png", "caption": "Hasil setelah submit" }
  ],
  "Validation AMF Segmentation::2.1": ["shots/2-1_1.png"]
}
```

Perilaku generator:

- `Heading 1` = judul module, ditentukan berurutan: judul eksplisit dari `titles=` (apa adanya,
  tanpa diubah huruf besar-kecilnya) → metadata `module_name` sheet itu → metadata
  `project_name` sheet itu → `project_name` workbook (hanya kalau workbook cuma punya satu
  module) → nama sheet. Judul otomatis dikapitalkan untuk heading, tetapi nama file memakai
  huruf aslinya. Kalau dua module sampai menghasilkan judul otomatis yang sama, keduanya
  kembali memakai nama sheet — dokumen dengan dua `Heading 1` identik tidak ada gunanya;
- `Heading 2` = `<Test Case ID>` + **TAB** + `<Test Case>`, memakai style Word asli sehingga
  muncul di Navigation Pane;
- screenshot disisipkan berurutan tepat di bawah heading test case-nya, selebar area cetak
  (tinggi menyesuaikan sendiri), caption opsional dicetak miring di bawah gambar;
- test case tanpa screenshot tetap dibuat headingnya + teks `[Belum ada screenshot]`, dan
  dihitung di ringkasan per module sebelum export;
- file screenshot yang hilang atau rusak tidak menggagalkan export — diganti catatan di
  dokumen dan dilaporkan di ringkasan;
- nama output default `Result Test UAT <Module>.docx` (per module) atau
  `Result Test UAT <Project Name>.docx` (gabungan, dari metadata `project_name`).

Arti tanda pada `inspect`:

| Tanda | Arti |
|-------|------|
| `[x]` | Sheet skenario — kolom wajib lengkap, siap dipakai |
| `[?]` | Mirip sheet skenario tetapi ada kolom wajib yang belum dikenali — perlu mapping manual |
| `[ ]` | Bukan sheet skenario (Main, Changes Logs, ZIP Group, dll.) |

Tanda `[?]` hanya *hint*: syaratnya kolom **Test Case ID** terdeteksi plus satu kolom wajib
lainnya. Sheet rekap seperti `Main` dan `Summary Scenario` sering memakai kata "Test Case"
dan "Test Step" sebagai judul kolom hitungan, dan dengan aturan ini tidak ikut disarankan.
Di GUI, dialog mapping manual tetap bisa dibuka untuk sheet apa pun.

## Validasi terhadap file contoh

Parser diuji terhadap `Test Case - New Validation AMF Segmentation.xlsx`:

| Yang dihitung | Rekap sheet `Main` | Hasil parser |
|---------------|--------------------|--------------|
| Test Case | 32 | **32** |
| Test Step | 151 | **151** |
| Test Scenario | 17 | 16 (+1 peringatan) |

Selisih scenario memang benar: baris 87 hanya berisi teks *Test Scenario* tanpa Scenario ID,
Test Case ID, maupun Test Step, jadi tidak membentuk scenario yang punya test case. Rekap di
sheet `Main` ikut menghitung baris itu. Parser melaporkannya sebagai peringatan, bukan
membuangnya diam-diam:

```
! Sheet 'Validation AMF Segmentation': baris 87 terisi sebagian tetapi tidak punya
  Test Step dan tidak membentuk test case baru, jadi diabaikan.
```

Sheet skenario itu juga terdeteksi dengan header di baris 7, 16 kolom terpetakan (termasuk
header bergaya `Pre - Condition`, `Pass/Fail (P/F)`, `Test Executed by (Name)`), 6 metadata
terbaca, dan kolom `Use Katalon` / `Secondary Test (Y)` dilewati tanpa error.

## Menjalankan test

```powershell
python -m pytest
```

## Screenshot & global hotkey

**Screenshot > Pengaturan screenshot...** mengatur dua hal, dan keduanya disimpan di
`project.json`:

| Pengaturan | Pilihan |
|------------|---------|
| **Hotkey** | default `Ctrl+Shift+S`. Harus memakai minimal satu modifier (Ctrl/Alt/Shift/Win) plus satu tombol utama. Bisa huruf, angka, `F1`-`F24`, atau tombol khusus seperti `PrintScreen`, `Insert`, `PageDown` |
| **Mode capture** | *Seluruh layar* (semua monitor digabung), *Monitor tertentu*, atau *Pilih area (drag)* |

Penulisan hotkey bebas urutan dan huruf besar-kecil: `shift+ctrl+s`, `Ctrl-Shift-S`, dan
`CTRL + SHIFT + S` semuanya dinormalkan jadi `Ctrl+Shift+S`. Dialognya memvalidasi sambil
diketik dan menolak OK kalau hotkey-nya tidak masuk akal (misalnya `Ctrl` saja, atau `S`
tanpa modifier — itu akan menyala setiap kali Anda mengetik huruf S di aplikasi apa pun).

Hotkey didaftarkan lewat hook keyboard tingkat OS (pynput), jadi **tetap aktif walaupun
aplikasi diminimize atau tidak sedang fokus**. Kalau pendaftaran gagal (biasanya karena
kombinasi itu sudah dipakai aplikasi lain), aplikasi memberi tahu dan screenshot tetap bisa
diambil lewat **Screenshot > Ambil screenshot sekarang**.

### Alur sekali tekan

1. Jendela aplikasi disembunyikan dulu kalau sedang tampil, supaya tidak ikut terfoto.
2. Layar diambil sesuai mode. Untuk *Pilih area*, layar **dibekukan dulu** lalu overlay
   gelap muncul di atasnya — persis Snipping Tool. Seret untuk memilih (ukuran area tampil
   saat menyeret), `Esc` untuk batal. Karena yang dipotong adalah gambar beku itu,
   overlay-nya sendiri mustahil ikut terfoto.
3. Popup tujuan muncul: preview, pilihan module, **kotak cari yang mencari di semua module**,
   daftar test case dengan format `Test Case ID - Test Case` plus jumlah screenshot yang
   sudah ada, dan caption opsional.
4. Pilihannya **default ke test case terakhir** yang dipakai, jadi rentetan capture cukup
   Enter. Tombol **Next test case** maju ke test case berikutnya sesuai urutan skenario
   (menyeberang module dan berputar di akhir); kalau test case berikutnya sedang tersembunyi
   oleh filter, filternya ikut dipindahkan supaya tetap terjangkau.
5. **Simpan** menyimpan file ke
   `screenshots/<module-slug>/<TestCaseID>_<test-case-slug>/<TestCaseID>_<urutan>_<timestamp>.png`
   dan tree langsung melompat ke test case itu. **Buang** membuang gambarnya tanpa menulis
   apa pun.

Menekan hotkey lagi selagi popup terbuka diabaikan, jadi capture tidak pernah bertumpuk.

### Catatan DPI

mss bekerja dalam pixel fisik, Qt dalam pixel logis. Pada mode *Pilih area*, faktor skala
dihitung dari perbandingan lebar hasil capture dengan lebar virtual desktop menurut Qt, lalu
dipakai saat memotong — jadi di layar 150%/200% area yang Anda pilih tetap pas. Setup
multi-monitor dengan DPI berbeda-beda per monitor belum diuji.

## Tulis hasil ke salinan Excel

Tab **Hasil** menyimpan empat kolom per step di `project.json`:

| Kolom | Catatan |
|-------|---------|
| `actual_results` | teks bebas |
| `pass_fail` | bebas; dropdown menawarkan `P`, `F`, `N/A` |
| `date_executed` | `YYYY-MM-DD` |
| `test_executed_by` | nama penguji |

Nilai yang sudah ada di file Excel ikut **ditampilkan** sebagai acuan, tetapi tidak
tersimpan sebagai hasil sampai Anda mengeditnya — jadi write-back tidak pernah menulis ulang
data yang memang sudah benar di sana. Field yang dikosongkan dihapus dari project, bukan
disimpan sebagai string kosong, supaya tidak menimpa isi workbook dengan nilai kosong.

Hasil disimpan dengan kunci **Test Step ID** (`2.1.1`), bukan nomor baris, jadi tetap
menempel di step yang benar setelah Re-import walaupun barisnya bergeser. Sheet tanpa kolom
Test Step ID memakai posisi step (`#1`, `#2`).

### Kenapa bukan openpyxl

Membuka lalu menyimpan ulang file skenario asli dengan openpyxl **menghilangkan isi**:
`image2.emf` hilang sama sekali, logo JPEG dikonversi jadi PNG, dan `xl/richData`,
`customXml`, `metadata.xml` serta printer settings ikut terbuang (603 KB jadi 525 KB pada
file contoh). File hasil adalah dokumen yang diserahkan ke orang lain, jadi itu tidak bisa
diterima.

Karena itu write-back bekerja langsung di dalam zip: file disalin **byte per byte**, lalu
hanya satu bagian `xl/worksheets/sheetN.xml` yang disunting. Diuji pada file contoh: dari
seluruh isi paket, **hanya `xl/worksheets/sheet5.xml` yang berubah**, dan Excel membukanya
tanpa prompt perbaikan dengan keenam gambar utuh.

Dua detail yang membuatnya aman:

- **Teks ditulis sebagai `inlineStr`**, jadi `sharedStrings.xml` tidak perlu disentuh sama
  sekali.
- **Tanggal ditulis sebagai serial number hanya kalau sel tujuannya memang sudah berisi
  angka** (artinya sudah berformat tanggal); kalau selnya kosong, tanggal ditulis sebagai
  teks `YYYY-MM-DD`. Tanpa aturan ini, menulis serial ke sel tanpa format tanggal akan
  tampil sebagai `46322` di layar.

Semua deklarasi `xmlns` di elemen root ikut dipertahankan. Ini bukan detail kosmetik:
`mc:Ignorable="x14ac xr xr2 xr3"` menyebut prefix secara literal, dan ElementTree hanya
menuliskan namespace yang benar-benar terpakai — tanpa pemulihan ini Excel **menolak membuka
file**.

Nama file default: `<nama file skenario> - Result.xlsx` di `<project>/excel/`. Menulis ke
file skenario asli ditolak secara eksplisit.

## Export PDF

**Export > Export PDF** mengubah semua `.docx` di folder `output/` menjadi PDF. Dua
converter, dicoba berurutan:

| Converter | Syarat | Catatan |
|-----------|--------|---------|
| Microsoft Word | Word terpasang + `pip install pywin32` | hasil paling setia, karena dirender oleh program yang sama dengan pembuat dokumennya |
| LibreOffice | `soffice` di PATH atau di lokasi standar | jalan tanpa Word, tata letak bisa sedikit bergeser |

Deteksi converter memakai pembacaan registry, bukan menjalankan Word — menyalakan dan
menutup Word hanya untuk bertanya "ada tidak?" memakan beberapa detik dan membuat COM
melempar error RPC kalau dilakukan berulang.

Kalau tidak ada converter sama sekali, aplikasi menjelaskan apa yang kurang; `.docx` tetap
bisa dibuat seperti biasa.

## Project

Satu project = satu folder:

```
<folder project>/
├─ project.json        pengaturan + mapping screenshot
├─ screenshots/
│  └─ <module-slug>/<TestCaseID>_<test-case-slug>/<TestCaseID>_<urutan>_<timestamp>.png
├─ excel/              salinan .xlsx berisi hasil eksekusi
└─ output/             dokumen .docx dan .pdf hasil generate
```

Isi `project.json`:

| Key | Isi |
|-----|-----|
| `excel_path` | path file skenario (relatif kalau file-nya ada di dalam folder project) |
| `selected_sheets` | module yang dipilih |
| `column_overrides` | hasil mapping kolom manual per sheet (`header_row` + `columns`) |
| `metadata` | metadata hasil edit manual (menimpa nilai dari Excel) |
| `module_titles` | judul `Heading 1` per module |
| `screenshots` | `{"<module>::<Test Case ID>": [{file, caption, captured_at}]}` |
| `orphan_screenshots` | screenshot yang test case-nya sudah tidak ada |
| `results` | `{"<module>::<Test Case ID>": {"<Test Step ID>": {actual_results, pass_fail, date_executed, test_executed_by}}}` |
| `snapshots` | isi test case saat bukti pertama diambil, dasar deteksi perubahan skenario |
| `needs_review` | test case yang berubah sejak bukti diambil, beserta alasannya |
| `orphan_step_results` | hasil eksekusi milik step yang sudah tidak ada di Excel |
| `excel_sha256` | hash file skenario saat terakhir dibaca |
| `hotkey` | kombinasi global hotkey, mis. `"Ctrl+Shift+S"` |
| `capture_mode`, `capture_monitor` | `fullscreen` / `monitor` / `region`, dan nomor monitor |
| `output` | mode output (`per_module` / `combined`) + path template |

Path screenshot disimpan **relatif terhadap folder project**, jadi foldernya bisa dipindah,
disalin, atau di-zip tanpa merusak mapping. `project.json` ditulis secara atomik (file
sementara lalu di-*replace*) supaya tidak pernah separuh tertulis kalau aplikasi mati.

### Re-import dan screenshot orphan

**Re-import Excel** membaca ulang skenario tanpa menyentuh screenshot yang sudah diambil:

- test case yang masih ada → screenshot-nya tetap menempel;
- test case yang **hilang** dari Excel → screenshot-nya jadi **orphan**; file-nya *tidak*
  dihapus, hanya dipindahkan ke `orphan_screenshots`;
- test case yang **muncul lagi** → screenshot orphan-nya dipulihkan otomatis;
- judul module yang modulenya hilang ikut dibersihkan.

Lewat **Project > Screenshot orphan...** screenshot orphan bisa dipindahkan ke test case
lain atau dibuang permanen. Karena identitas test case adalah `module + Test Case ID`,
Test Case ID yang sama di dua sheet tidak pernah tertukar.

### Skenario diedit setelah capture dimulai

Pembuat skenario bisa menambah/menghapus test case, mengganti nama, mengubah step, atau
me-renumber ID. Supaya bukti tidak diam-diam menempel ke test case yang salah:

- Saat screenshot/hasil pertama sebuah test case disimpan, project menyimpan **snapshot**
  isinya (nama, step, expected result) di `project.json`. Project lama dibuatkan snapshot
  baseline saat dibuka.
- Setiap **Re-import (F5)** dan setiap **buka project** yang file Excel-nya berubah (dicek
  lewat hash) membandingkan snapshot dengan isi Excel terbaru:

  | Status | Arti | Aksi yang disarankan |
  |--------|------|----------------------|
  | Berubah isinya | uid sama, isi beda | bukti tetap, test case ditandai ⚠ sampai direview |
  | Bergeser / ter-renumber | uid sama tapi isinya mirip test case lain | saran pindah ke test case yang mirip |
  | Hilang, ada pengganti | uid hilang, ada test case mirip (>= 80%) | saran pindah |
  | Hilang | uid hilang tanpa pengganti | jadi orphan seperti biasa |

- Dialog **Review perubahan skenario** (otomatis muncul, atau **Project > Review perubahan
  skenario...**) menampilkan saran. **Tidak ada yang dipindahkan sebelum Terapkan**; baris
  dengan kandidat ambigu tidak dicentang otomatis. Pemindahan dilakukan sekaligus, jadi
  rantai (2.1 → 2.2, 2.2 → 2.3) dan tukar posisi aman. File screenshot di disk tidak
  dipindah atau dihapus.
- Hasil eksekusi dipindahkan bersama screenshot; kunci step diganti mengikuti pasangan
  step terdekat. Hasil milik step yang dihapus disimpan di `orphan_step_results` (tidak
  ditulis ke Excel, dan kembali sendiri kalau step-nya muncul lagi).
- Sebelum re-import, `project.json` disalin ke `backups/` (10 terakhir disimpan). Tombol
  **Undo re-import** di dialog mengembalikan backup terakhir.
- Test case yang masih ⚠ memunculkan konfirmasi saat **Generate dokumen Word**, dan
  muncul di tab Peringatan. Klik kanan test case > *Tandai sudah direview* untuk menghapus
  tanda.

## Build .exe (PyInstaller)

Cara tercepat: `.\build.ps1` (jalankan test, build satu folder, smoke test
`--selftest`, lalu zip ke `dist\ResultTestGenerator-<versi>-win64.zip`). Penerima cukup
mengekstrak zip dan menjalankan `Result Test Generator.exe`, tanpa Python.

**Release otomatis:** naikkan versi di `rtg/__init__.py` dan `pyproject.toml`, lalu

```powershell
git tag v0.2.0
git push origin v0.2.0
```

Workflow `.github/workflows/release.yml` membangun zip di GitHub dan melampirkannya ke
halaman Releases. Tag harus sama dengan versi di `rtg/__init__.py`, kalau tidak build
dihentikan. Windows SmartScreen bisa memperingatkan karena `.exe` belum ditandatangani
(klik *More info > Run anyway*); export PDF tetap butuh Word atau LibreOffice di komputer
penerima.

Perintah yang dipakai skrip (satu folder, bukan satu file, supaya start cepat dan lebih
jarang ditandai antivirus). Yang harus ikut di-bundle: alias kolom bawaan
(`rtg/data/column_aliases.json`) dan font ikon qtawesome.

```powershell
pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onedir --windowed `
  --name "Result Test Generator" `
  --add-data "rtg/data/column_aliases.json;rtg/data" `
  --collect-all qtawesome `
  --hidden-import pynput.keyboard._win32 --hidden-import pynput.mouse._win32 `
  app.py
```

Hasilnya folder `dist\Result Test Generator\` (sekitar 75 MB setelah di-zip). Cek hasil build
dengan `"dist\Result Test Generator\Result Test Generator.exe" --selftest` (kode keluar 0 =
aman).

---

## Cara kerja parser

### Deteksi sheet skenario

Nama sheet **tidak pernah** dipakai sebagai patokan. Sebuah sheet dianggap sheet skenario
jika pada 30 baris pertama ada baris header yang memuat ketiga kolom wajib:
**Test Case ID**, **Test Case**, dan **Test Step**. Nama module = nama sheet.

### Pemetaan kolom

Posisi/huruf kolom (A, B, C...) tidak pernah dipakai. Teks header dinormalisasi lalu
dicocokkan dengan daftar alias:

1. huruf kecil semua;
2. semua tanda baca dan tanda kurung diganti spasi;
3. spasi berlebih dan newline dirapikan;
4. deretan huruf tunggal disatukan kembali.

Contoh: `Pass/Fail (P/F)` → `pass fail pf`, `Pre-Condition` → `pre condition`.

Pencocokan memakai skor: kecocokan **persis** selalu menang atas kecocokan sebagian, jadi
sheet yang punya kolom `Test Case` dan `Test Case ID` sekaligus tetap terpetakan benar
berapa pun urutannya. Setiap kolom dan setiap field hanya dipakai sekali.

Kolom wajib: `test_case_id`, `test_case`, `test_step`. Kolom lain opsional — jika tidak ada,
dilewati tanpa error.

### Alias kolom (bisa diedit)

Default ada di `rtg/data/column_aliases.json`. Untuk menambah alias sendiri, edit
`config/column_aliases.json` — isinya **digabung** di atas default, jadi alias bawaan tidak
hilang. Yang bisa diatur:

| Key | Fungsi |
|-----|--------|
| `aliases` | nama header per field |
| `metadata_aliases` | label metadata di atas header |
| `required_fields` | kolom wajib |
| `fill_down_groups` | kolom yang di-fill-down dan kunci reset-nya |
| `header_scan_rows` | berapa baris pertama yang dipindai untuk cari header (default 30) |
| `max_blank_rows` | berhenti baca setelah sekian baris kosong berturut-turut (default 10) |
| `metadata_scan_columns` | sejauh apa nilai metadata dicari ke kanan (default 10) |

### Fill-down

Kolom yang hanya terisi di baris pertama tiap kelompok (`Scenario ID`, `Test Scenario`,
`Test Case ID`, `Test Case`) diisi turun ke baris berikutnya. Fill-down bertingkat:
begitu **Test Case ID** baru muncul, nama test case yang dibawa dari atas dibuang — jadi
test case baru tidak pernah mewarisi nama test case sebelumnya. Sel yang di-*merge* di
Excel diperlakukan sama seperti sel terisi biasa.

Pembacaan berhenti setelah 10 baris kosong berturut-turut; jika ternyata masih ada data
di bawahnya, hal itu dilaporkan sebagai peringatan (bukan error).

### Metadata

Metadata (`Project Name`, `Version`, `Environment Testing`, `Type Testing`, `Created By`,
`Created Date`, ...) dicari di baris-baris **di atas** header berdasarkan teks labelnya,
lalu nilainya diambil dari sel pertama di kanannya yang tidak kosong (kalau tidak ada,
sel tepat di bawahnya). Metadata dibaca per sheet; `ScenarioFile.merged_metadata()`
memakai nilai dari sheet lain sebagai cadangan bila sebuah sheet tidak punya metadata.

### Identitas test case

Test Case ID yang sama boleh muncul di module berbeda. Karena itu identitas unik sebuah
test case adalah **module + Test Case ID**, dipakai lewat `TestCase.uid`
(`"New Validation AMF Segmentation::1.1"`). Nilai inilah yang nanti dipakai sebagai kunci
mapping screenshot di `project.json`.

**Test Case ID ganda dalam satu sheet tidak digabung.** ID kembar biasanya salah ketik,
jadi tiap kemunculan tetap menjadi test case sendiri, dan parser memberi peringatan:

```
Test Case ID '4.2' muncul 2 kali di sheet 'Parameter Decryption Key JKT' (baris 12, 30);
tetap dibuat sebagai test case terpisah. Cek apakah ID-nya salah ketik.
```

- Kemunculan pertama memakai uid biasa (`Sheet::4.2`, tidak berubah, jadi project lama tetap
  cocok). Kemunculan kedua dan seterusnya: `Sheet::4.2~2`, `Sheet::4.2~3`.
- Baris dianggap test case baru kalau ID-nya pernah muncul tetapi baris sebelumnya punya ID
  lain, atau kalau ID-nya sama dengan baris sebelumnya tetapi nama Test Case-nya berbeda.
  Baris lanjutan (Test Case kosong atau sama) tetap masuk test case yang sedang dibaca.
- Di GUI/CLI labelnya diberi `[ID ganda #2]` dan folder screenshotnya `4.2-2_<slug>`.
  Judul `Heading 2` di Word tetap memakai ID persis seperti di Excel.
- Project lama yang dulu menyimpan hasil step dari test case yang tergabung: hasil itu
  dipindahkan otomatis ke test case yang memiliki step tersebut saat re-import/buka project.
  Screenshot lama tetap di kemunculan pertama; pindahkan lewat *Pindah test case...*.

### Keamanan file sumber

File Excel asli **tidak pernah** ditulis atau di-overwrite; workbook hanya dibuka untuk
dibaca lalu ditutup. Ada unit test yang memverifikasi hash dan mtime file tidak berubah
setelah parsing.

### Pesan error

| Situasi | Error |
|---------|-------|
| File tidak ada | `ScenarioFileNotFoundError` |
| Bukan .xlsx/.xlsm (mis. .xls lama) | `UnsupportedFileFormatError` |
| File terkunci proses lain | `FileLockedError` |
| File rusak / bukan workbook | `CorruptWorkbookError` |
| Tidak ada sheet skenario sama sekali | `NoScenarioSheetError` |
| Sheet dipilih tapi kolom wajib kurang | `MissingRequiredColumnsError` (membawa daftar header untuk dialog mapping manual) |
| Nama sheet di project.json sudah tidak ada | `SheetNotFoundError` |

Semua pesan error berbahasa Indonesia dan siap ditampilkan apa adanya di GUI.
Jika file sedang terbuka di Excel (`~$nama.xlsx` ada), file tetap dibaca tetapi
ditambahkan peringatan.

---

## Struktur project

```
Result-Generator-NF/
├─ app.py                  # entry point GUI
├─ rtg/
│  ├─ __init__.py          # API publik
│  ├─ config.py            # load & merge config alias
│  ├─ errors.py            # exception + pesan user
│  ├─ excel_parser.py      # inspect_workbook / parse_workbook
│  ├─ docx_builder.py      # summarize / build_document / export
│  ├─ capture.py           # grab / crop / save_png (mss + Pillow, tanpa Qt)
│  ├─ hotkey.py            # parse_hotkey / HotkeyListener (pynput, tanpa Qt)
│  ├─ excel_writer.py      # tulis hasil ke salinan .xlsx (bedah XML di dalam zip)
│  ├─ pdf_export.py        # .docx -> .pdf lewat Word COM atau LibreOffice
│  ├─ project.py           # Project: project.json, screenshot, re-import, orphan
│  ├─ scenario_diff.py     # snapshot test case + deteksi perubahan skenario
│  ├─ models.py            # ScenarioFile > Module > ScenarioGroup > TestCase > TestStep
│  ├─ normalize.py         # normalisasi header, teks sel, slug
│  ├─ cli.py               # CLI untuk verifikasi parser & generate Word
│  ├─ gui/
│  │  ├─ main_window.py    # jendela utama + tree + hotkey/capture
│  │  ├─ dialogs.py        # sheet picker, mapping kolom, metadata, export, orphan
│  │  ├─ capture_dialogs.py # overlay pilih area, popup tujuan, pengaturan capture
│  │  └─ results_panel.py  # tabel hasil per step + dialog isi massal
│  └─ data/column_aliases.json   # alias bawaan
├─ config/column_aliases.json    # alias milik user (menimpa/menambah default)
├─ tests/
│  ├─ fixtures.py          # generator workbook sintetis + PNG
│  ├─ test_normalize.py
│  ├─ test_config.py
│  ├─ test_excel_parser.py
│  ├─ test_docx_builder.py
│  ├─ test_project.py
│  ├─ test_capture.py      # dilewati kalau tidak ada display
│  ├─ test_hotkey.py       # pakai pynput palsu, tidak memasang hook asli
│  ├─ test_excel_writer.py
│  ├─ test_project_results.py
│  ├─ test_pdf_export.py   # konversi asli dilewati kalau tidak ada converter
│  ├─ test_gui.py          # offscreen, dilewati kalau PySide6 belum ada
│  ├─ test_gui_capture.py
│  └─ test_gui_results.py
├─ requirements.txt
└─ pyproject.toml
```

### API singkat

```python
from rtg import inspect_workbook, parse_workbook

# 1) untuk daftar checkbox sheet di GUI
inspection = inspect_workbook("Scenario.xlsx")
inspection.scenario_sheet_names      # ['AMF Segmentation', 'Billing Adjustment']
inspection.mapping_candidates        # sheet yang perlu mapping manual

# 2) untuk tree + generate Word
result = parse_workbook("Scenario.xlsx", ["AMF Segmentation"])
for module in result.modules:
    module.title                     # nama sheet kapital: 'AMF SEGMENTATION'
                                     # (judul final lihat resolve_naming di bawah)
    for case in module.test_cases:
        case.heading                 # Heading 2: '1.1\tMove Partner Onboarding File ...'
        case.display_label           # untuk tree GUI: '1.1  Move Partner Onboarding ...'
        case.uid                     # kunci screenshot: 'AMF Segmentation::1.1'
        case.folder_name             # '1.1_move-partner-onboarding-file-via-sharefolder'
```

Lalu untuk membuat dokumennya:

```python
from rtg import resolve_naming, summarize, export

# judul & nama file yang akan dipakai tiap module
resolve_naming(result)
# {'Validation AMF Segmentation': ModuleNaming(
#     title='NEW VALIDATION AMF SEGMENTATION',
#     file_label='New Validation AMF Segmentation',
#     source='project_name')}

# ringkasan dulu (test case yang belum ada screenshot-nya)
for report in summarize(result, screenshots):
    print(report.title, report.test_case_count, report.without_screenshot)

export(
    result, screenshots,
    output_dir="output",
    combined=False,                      # True = satu file gabungan
    titles={"AMF Segmentation": "NEW VALIDATION AMF SEGMENTATION"},
    template="punya-saya.docx",          # opsional
)
```

`case.heading` memakai **tab** antara Test Case ID dan nama test case, mengikuti file
`Result Test UAT AMF Segmentation.docx` yang dijadikan acuan (`"2.1\tSet Single Disburse to
0 by Query"`). Acuan lain dari file itu untuk Tahap 4: Heading 1 = nama module huruf kapital,
screenshot disisipkan sebagai gambar inline di paragraf `Normal` tepat di bawah heading
test case-nya, dan lebar area cetak 6,27 inci (A4, margin 1 inci).

Mapping manual disimpan sebagai `overrides` dan nanti ikut masuk ke `project.json`:

```python
parse_workbook(
    "Scenario.xlsx", ["Scenario"],
    overrides={"Scenario": {"header_row": 9, "columns": {"test_case": 4}}},
)
```

### Project lewat Python

```python
from rtg.project import Project

project = Project.create("D:/UAT/AMF - Result", "Scenario.xlsx",
                         sheets=["Validation AMF Segmentation"])
project.add_screenshot("Validation AMF Segmentation::1.1", "bukti.png",
                       caption="File tersalin")
project.generate()                 # -> output/Result Test UAT ....docx

report = project.reimport()        # setelah file Excel diperbarui
print(report.summary_lines())
```

### Screenshot lewat Python

```python
from rtg.capture import MODE_REGION, capture_to_file, grab, list_monitors
from rtg.hotkey import HotkeyListener

list_monitors()                                   # [Monitor(0, ...), Monitor(1, ...)]
capture_to_file("shot.png")                       # seluruh layar
capture_to_file("area.png", mode=MODE_REGION, box=(100, 100, 800, 600))

with HotkeyListener("Ctrl+Shift+S", lambda: print("jepret")) as listener:
    ...                                           # aktif selama blok ini
```

### Hasil & PDF lewat Python

```python
from rtg.pdf_export import docx_to_pdf
from rtg.project import Project

project = Project.open("D:/UAT/AMF - Result")

project.set_results_for_test_case("Validation AMF Segmentation::2.1", {
    "pass_fail": "P",
    "date_executed": "2026-10-27",
    "test_executed_by": "Mahrub",
})
report = project.write_result_workbook()     # -> excel/<nama> - Result.xlsx
print(report.summary_lines())

for document in project.output_dir.glob("*.docx"):
    docx_to_pdf(document)
```
