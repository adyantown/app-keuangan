# 🇮🇩 Sistem Monitoring Kas Bendahara & GUP

**Satuan Kerja: KPU Kabupaten Tulang Bawang Barat**

Aplikasi desktop sederhana berbasis Python untuk membantu Bendahara Pengeluaran dalam mencatat transaksi harian, memonitor sisa Uang Persediaan (UP), dan mengelola mekanisme Ganti Uang Persediaan (GUP/Revolving Fund) secara transparan dan akurat.

---

## 🌟 Fitur Utama

1. **Dashboard Real-time:**

- **Saldo Kas Tunai:** Menampilkan uang fisik yang seharusnya ada di brankas (Pagu - Belanja + GUP).
- **Realisasi Bulan Ini:** Menampilkan total kuitansi belanja khusus bulan berjalan (tidak terreset meski ada GUP masuk).

2. **Pencatatan Belanja:** Form input lengkap (Tanggal, Uraian, Nominal, No. Rek, Penerima, Penyedia, Kode Akun).
3. **Sistem GUP Otomatis:** Deteksi otomatis jika penggunaan dana sudah mencapai 50% dan tombol instan untuk _top-up_ saldo (Revolving).
4. **Edit Data:** Kemudahan mengoreksi kesalahan input tanpa perlu bongkar file Excel.
5. **Database Excel:** Data tersimpan dalam file `.xlsx` yang bisa langsung dibuka, diprint, atau diaudit.
6. **Shortcut Excel:** Tombol cepat untuk membuka file laporan langsung dari aplikasi.

---

## 🚀 Cara Menjalankan Aplikasi

### Untuk Pengguna (User/Bendahara)

1. Pastikan file aplikasi (`app_keuangan.exe`) berada di dalam folder yang aman.
2. **Klik 2x** pada `app_keuangan.exe`.
3. Tunggu sebentar hingga jendela aplikasi dengan kop **KPU TULANG BAWANG BARAT** terbuka.
4. Aplikasi akan otomatis membuat file database `pembukuan_up_kpu_tubaba.xlsx` jika belum ada.

---

## 📖 Panduan Penggunaan Singkat

### 1. Mencatat Pengeluaran Baru

- Isi semua kolom pada kotak **"Input Transaksi Belanja"**.
- Pastikan **Nominal** diisi angka saja (tanpa titik/koma).
- Klik tombol **"SIMPAN PENGELUARAN"**.
- Data akan muncul di tabel bawah dan saldo akan berkurang otomatis.

### 2. Mengedit Kesalahan Input

- Klik **salah satu baris** pada tabel yang ingin diubah.
- Klik tombol biru **"EDIT DATA TERPILIH"**.
- Data akan naik kembali ke form input.
- Lakukan perbaikan, lalu klik tombol oranye **"UPDATE DATA"**.

### 3. Melakukan GUP (Revolving Fund)

- Perhatikan tombol di **Pojok Kanan Atas**.
- Jika tombol berubah warna menjadi **ORANYE**, artinya pemakaian dana sudah >50%.
- Klik tombol tersebut jika dana pengganti (SP2D GUP) sudah cair/diterima tunai.
- Klik **YES** pada pesan konfirmasi. Saldo Kas akan otomatis terisi kembali ke Pagu Awal (Rp 11.400.000).

### 4. Membuka Laporan Excel

- Klik tombol hijau **"📂 BUKA FILE EXCEL"** di sebelah kanan atas tabel.
- File Excel akan terbuka otomatis untuk keperluan cetak laporan atau copy-paste data.

---

## ⚠️ PENTING: Hal yang Harus Diperhatikan

1. **DILARANG Membuka Excel Saat Menyimpan Data:** Pastikan file Excel **TERTUTUP** saat kamu menekan tombol _Simpan_ atau _Update_ di aplikasi. Jika Excel sedang terbuka, aplikasi akan menolak menyimpan (muncul error "Permission Error") untuk mencegah data korup.
2. **Jangan Ubah Nama File Excel:** Biarkan nama file `pembukuan_up_kpu_tubaba.xlsx` apa adanya. Jangan di-rename.
3. **Jangan Hapus Baris Header di Excel:** Jika mengedit manual via Excel, jangan menghapus baris pertama (Judul Kolom) agar aplikasi tetap bisa membaca data.

---

**Dibuat oleh:** [Adyantown x GeminiAI] <br> **Versi:** 1.0 (Stable Excel Release)
