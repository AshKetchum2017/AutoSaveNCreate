# AutoSaveNCreate

AutoSaveNCreate adalah macro VBA untuk CorelDRAW yang dirancang untuk membuat folder tujuan dan menyimpan dokumen CDR aktif ke lokasi produksi sesuai mode pekerjaan dan pola penamaan operator.

Macro ini dibuat untuk mengurangi pekerjaan manual saat menyiapkan hasil produksi, terutama ketika file harus disimpan ke root directory berbeda, mengikuti struktur tanggal sumber, dan menggunakan nama folder yang konsisten.

AutoSaveNCreate mendukung empat mode kerja, yaitu **DieA**, **HiDie**, **KissA**, dan **HiKiss**, serta menyediakan pengaturan tambahan seperti **DisposableSave**, versi output CDR, dan **EmbedColorProfiles**.

## Main Features

### Save and Create Folder

AutoSaveNCreate membaca dokumen aktif, menentukan directory tujuan, membuat folder yang belum tersedia, lalu menjalankan **Save As** dalam format CDR.

Pada penyimpanan normal, lokasi hasil dibentuk dari:

- root directory untuk mode yang dipilih;
- tahun, bulan, dan tanggal dari struktur folder sumber;
- nama folder hasil parsing token operator;
- nama file CDR asal.

Seluruh page dokumen ikut disimpan. Proses ini bukan export objek terpilih dan bukan penyimpanan berkala di latar belakang.

Dokumen sumber harus sudah tersimpan sebagai file `.cdr` sebelum diproses.

---

## Processing Modes

AutoSaveNCreate menyediakan empat pilihan mode untuk memisahkan tujuan penyimpanan:

| Mode | Caption | Kontrol |
|---|---|---|
| `DieA` | Die Cut A3 | `optDieA` |
| `HiDie` | Hires Die Cut | `optHiDie` |
| `KissA` | Kiss Cut A3 | `optKissA` |
| `HiKiss` | Hires Kiss Cut | `optHiKiss` |

Setiap mode mempunyai root directory permanen sendiri. Mengubah directory untuk satu mode tidak mengganti directory mode lainnya.

Mode di sini menentukan lokasi penyimpanan. Macro tidak mendistribusikan objek, mengubah Cut Line, atau menjalankan proses cutting.

### Initial Mode Detection

Saat membaca dokumen yang baru aktif, macro mencoba menentukan mode awal dari nama file CDR.

Deteksi menggunakan keyword jenis cutting dan format pekerjaan, seperti `KISS CUT`, `DIE CUT`, `A3`, `LASER`, `ECO SOLVENT`, dan `METERAN`, dengan toleransi typo terbatas.

Hasil deteksi membantu pemilihan awal. User tetap perlu memeriksa mode sebelum menjalankan **Process**.

Pilihan manual dipertahankan selama dokumen yang sama masih aktif. Jika mode tidak terdeteksi, pilihan dikembalikan ke nilai default dari UserForm.

---

## User Settings

Pengaturan operator dikelola melalui `UserSettingsMenu`. Setiap entry berisi nama operator dan pola token untuk membentuk nama folder hasil.

Daftar tersebut ditampilkan pada `lbxUserLists` dan dapat dikelola melalui:

- `cmdAdd`: menambahkan operator.
- `cmdModify`: mengubah entry terpilih.
- `cmdRemove`: menghapus entry terpilih.
- `cmdSave`: menyimpan daftar operator.

Form `AddUserSettings` menyediakan `txbUserName` untuk nama operator dan `txbUserTokens` untuk pola token. Nama operator harus unik dan pola token harus valid.

Operator yang akan digunakan dipilih melalui `cmbUserSelection` pada Main UserForm.

Pengaturan operator disimpan ke registry agar dapat dimuat kembali pada sesi berikutnya.

---

## Folder Tokens

Pola token menentukan nama folder hasil pada penyimpanan normal.

| Token | Perilaku |
|---|---|
| `/` | Mengambil kode kategori dari struktur folder sumber. |
| `~` | Mengambil nama folder tempat file CDR berada. |
| `"teks"` | Menambahkan teks literal tanpa tanda kutip pembungkus. |
| Spasi | Menambahkan pemisah spasi. |
| `-` | Menambahkan separator tanda hubung. |

Contoh pola:

`"r" - / ~`

Jika kategori sumber adalah `RESELLER` dan nama folder CDR adalah `Contoh Printing`, nama folder hasil menjadi:

`r - rc contoh printing`

Nama yang dibaca melalui `~` diubah menjadi huruf kecil. Parser juga menerapkan koreksi tertentu, misalnya `prnting` menjadi `printing` dan `rklame` menjadi `reklame`.

Koreksi ini berlaku pada nama folder hasil, bukan mengganti nama folder sumber.

### Category Codes

Token `/` membaca kategori pada susunan folder bulan, tanggal, dan kategori sumber.

| Contoh Nama Kategori | Kode |
|---|---|
| `RESELLER`, `RESELLER CENTER` | `rc` |
| `BLACKPAINT`, `BLACK PAINT PRINTSHOP` | `bp` |
| `FOLDER PRINTEX`, `FOLDERPRINT` | `fp` |
| `CORPORATE` | `co` |
| `CETAK ULANG` | `cu` |

Beberapa variasi ejaan kategori juga dikenali oleh parser. Kategori yang tidak dikenali akan menghentikan proses ketika pola menggunakan `/`.

Teks biasa harus dibungkus tanda kutip. Token `__` belum didukung.

---

## Output Folder Structure

Penyimpanan normal menggunakan struktur:

`Root Directory\Tahun\Bulan\Tanggal\Folder Hasil Token\Nama File Asal.cdr`

Contoh ilustratif sumber:

`D:\Sumber\2026\09. SEPTEMBER\12\RESELLER\Contoh Printing\Desain.cdr`

Dengan root directory `D:\Output` dan pola `"r" - / ~`, hasilnya menjadi:

`D:\Output\2026\09. SEPTEMBER\12\r - rc contoh printing\Desain.cdr`

Tahun, bulan, dan tanggal berasal dari folder sumber, bukan tanggal komputer saat **Process** dijalankan.

Folder bulan menggunakan format seperti `09. SEPTEMBER`. Tahun dicari dari folder empat digit terlebih dahulu, kemudian dari tahun yang terdapat dalam nama folder induk.

Jika struktur tanggal sumber tidak ditemukan atau tahun ambigu, proses normal dihentikan. Directory khusus dapat dipilih melalui **DisposableSave** untuk melewati pembentukan struktur tersebut.

---

## Directory Settings

Form `DirectorySettings` digunakan untuk mengatur root directory permanen dan directory sementara.

Pengaturan tiap mode dibuka melalui `cmdDieASetting`, `cmdHiDieSetting`, `cmdKissASetting`, atau `cmdHiKissSetting`.

Untuk menyimpan root directory permanen:

1. Buka pengaturan directory mode yang diinginkan.
2. Isi `txbDirectory` atau pilih folder melalui `cmdBrowse`.
3. Gunakan `cmdSave` untuk menyimpan root directory mode tersebut.

Pada penyimpanan normal, directory ini menjadi folder dasar. Macro masih menambahkan struktur tanggal dan folder hasil token di bawahnya.

---

## DisposableSave

**DisposableSave** digunakan ketika file perlu disimpan langsung ke folder tertentu tanpa mengikuti struktur tanggal dan pola token operator.

Directory yang dipilih menjadi folder akhir penyimpanan.

Contoh ilustratif:

| Input | Nilai |
|---|---|
| Directory sementara | `D:\Khusus\Order Revisi` |
| Nama file asal | `Desain.cdr` |
| Hasil | `D:\Khusus\Order Revisi\Desain.cdr` |

### Temporary Directory List

Directory sementara ditampilkan pada `lbxDisposableLists` di `DirectorySettings`.

- `cmdAdd` menambahkan directory yang sudah ada ke daftar sementara.
- `cmdSelect` memilih entry untuk dokumen aktif.
- `cmdRemove` menghapus entry dari daftar, bukan folder di disk.
- `cmdClear` mengosongkan daftar sementara, bukan menghapus folder di disk.

Saat directory sementara dipilih, pilihan mode pada Main UserForm dikunci. Proses tidak memerlukan pola token operator dan tidak menambahkan subfolder tahun, bulan, atau tanggal.

### Session Behavior

Daftar directory disimpan dalam memori melalui `SNCDirectorySession`, bukan registry.

Beberapa aturan penting:

- Daftar bertahan selama sesi VBA belum di-reset.
- Pilihan directory aktif terikat ke identitas dokumen.
- Menutup Main UserForm melepas pilihan aktif, tetapi tidak mengosongkan daftar.
- Setelah penyimpanan melalui DisposableSave berhasil, pilihan aktif dilepas.

Karena daftar bersifat sementara, isinya tidak menjadi konfigurasi permanen untuk sesi CorelDRAW berikutnya.

---

## CDR Version

Versi format hasil dipilih melalui `cmbCorelVersion`.

Pilihan yang tersedia mencakup `25.1 (2024.1)`, `25.0 (2024)`, `24.3 (2023)`, `24.0 (2022)`, serta versi utama `23.0` sampai `15.0`.

Pilihan terakhir disimpan ke registry. Default yang digunakan adalah `25.0 (2024)` jika pengaturan tersimpan tidak valid.

Daftar ini merupakan pilihan versi **output CDR**, bukan pernyataan bahwa macro telah diuji berjalan pada seluruh versi CorelDRAW tersebut.

---

## Embedding Options

### EmbedColorProfiles

`chkEmbedColorProfiles` menentukan apakah profil warna disertakan saat menyimpan CDR. Nilainya diterapkan ke `EmbedICCProfile` pada opsi penyimpanan.

Preferensinya disimpan dan dimuat kembali saat UserForm dibuka.

### EmbedFonts

`chkEmbedFonts` saat ini hanya menyimpan preferensi.

Penerapan **EmbedFonts** ke hasil Save As CDR belum tersedia dan masih menunggu pemetaan API. Mengaktifkan kontrol ini belum berarti font akan ditanamkan ke file hasil.

---

## Existing File Handling

Jika file dengan nama yang sama sudah ada di tujuan, macro menampilkan pilihan:

| Pilihan | Perilaku |
|---|---|
| **Yes** | Membuat file duplikat dengan suffix nomor, seperti `Desain (1).cdr`. |
| **No** | Menimpa file tujuan yang sudah ada. |
| **Cancel** | Membatalkan penyimpanan. |

Tanpa benturan nama, nama file asal tetap digunakan.

---

## Basic Workflow

1. Buka dokumen CorelDRAW yang sudah tersimpan sebagai `.cdr`.
2. Jalankan UserForm `AutoSaveNCreate`.
3. Pada penggunaan awal, siapkan operator dan pola token melalui `UserSettingsMenu`.
4. Atur root directory untuk mode yang akan digunakan melalui `DirectorySettings`.
5. Periksa pilihan **DieA**, **HiDie**, **KissA**, atau **HiKiss**.
6. Pilih operator melalui `cmbUserSelection`.
7. Pilih versi hasil melalui `cmbCorelVersion` dan atur `chkEmbedColorProfiles`.
8. Tekan **Process** melalui `cmdProcess`.
9. Jika nama file bentrok, pilih duplikat, overwrite, atau cancel.

Untuk **DisposableSave**, tambahkan dan pilih directory sementara melalui `DirectorySettings` sebelum menjalankan **Process**. Root directory mode dan pola token tidak digunakan pada jalur ini.

Setelah berhasil, macro menampilkan lokasi file hasil.

---

## Processing Safety

Sebelum menyimpan, AutoSaveNCreate memeriksa dokumen aktif, lokasi file sumber, versi output, dan directory tujuan.

Pada penyimpanan normal, macro juga memvalidasi pilihan operator, mode, pola token, struktur tanggal sumber, serta nama folder hasil.

Folder hasil diperiksa agar tetap berada di bawah root directory yang ditentukan. Nama folder yang mengandung karakter terlarang atau nama perangkat Windows yang dicadangkan akan ditolak.

Selama proses berlangsung, `cmdProcess` dinonaktifkan untuk mencegah pemrosesan berulang. Jika dokumen aktif berubah saat **Process** ditekan, macro meminta user memeriksa pilihan mode dan setting user, lalu menekan **Process** kembali.

Jika terjadi error, macro menampilkan informasi kegagalan. Validasi ini tidak menggantikan backup, terutama ketika memilih overwrite.

---

## Project Structure

### `AutoSaveNCreate.vba`

Code-behind untuk Main UserForm. Mengatur pilihan mode, operator, versi CDR, embedding preferences, sinkronisasi dokumen aktif, dan pemanggilan proses penyimpanan.

### `DirectorySettings.vba`

Code-behind pengaturan root directory per mode dan pengelolaan daftar DisposableSave.

### `UserSettingsMenu.vba` dan `AddUserSettings.vba`

Mengelola daftar operator serta pengisian nama dan pola token. Bantuan token melalui `cmdTokenHint` memiliki pesan fallback jika form `UserTokenHints` tidak tersedia.

### `SNCFolderParser.cls`

Menangani parsing token, pembacaan kategori dan tanggal sumber, koreksi nama folder tertentu, serta validasi nama folder hasil.

### `SNCModeDetector.cls`

Menangani deteksi mode berdasarkan keyword nama file atau directory, termasuk toleransi typo terbatas dan pemeriksaan sinyal yang bertentangan.

### `SNCSaveRunner.cls`

Menjalankan Save As CDR, membentuk lokasi hasil, membuat folder, dan menangani file tujuan yang sudah ada.

### `SNCSettingsStore.cls`

Mengelola pengaturan registry untuk operator, root directory, versi CDR, dan embedding preferences di bawah aplikasi `RinCorelMacros`, section `AutoSaveNCreate`.

### `SNCDirectorySession.cls` dan `SNCSession.bas`

Menyediakan penyimpanan directory sementara dan instance session yang digunakan bersama oleh UserForm.

### `Changelog.log`

Mencatat penambahan fitur, perubahan behavior, perbaikan, dan perkembangan AutoSaveNCreate.

---

## Notes

AutoSaveNCreate dibuat untuk workflow produksi tertentu di CorelDRAW. Struktur folder sumber, kategori, keyword deteksi mode, dan koreksi nama mengikuti aturan yang terdapat dalam source code.

Repository menyediakan source VBA dan code-behind UserForm. File code-behind tidak dengan sendirinya menyediakan layout visual UserForm; form dan kontrol dengan nama yang sesuai tetap diperlukan pada project VBA.

Penyesuaian workflow di luar aturan tersebut dapat memerlukan perubahan source code. Pilihan versi output tidak menjamin kompatibilitas seluruh fitur dokumen pada format CDR yang lebih lama.

---

## Current Scope

AutoSaveNCreate saat ini mencakup:

- penyimpanan dokumen aktif melalui Save As CDR;
- root directory terpisah untuk DieA, HiDie, KissA, dan HiKiss;
- pengaturan operator dan pola token folder;
- struktur hasil berbasis tanggal sumber;
- automatic folder creation;
- DisposableSave berbasis sesi;
- initial mode detection;
- pilihan versi output CDR;
- penerapan EmbedColorProfiles;
- preferensi EmbedFonts yang belum diterapkan ke file hasil;
- penanganan duplicate, overwrite, dan cancel;
- validasi input dan pelaporan error.

Fokus utamanya adalah mengurangi pekerjaan manual saat menentukan lokasi dan menyimpan file produksi, bukan mengubah susunan objek di dalam dokumen.

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
