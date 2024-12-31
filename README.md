# Deskripsi Program VBA untuk Pengolahan Data di Excel
Program VBA ini dirancang untuk memproses dan mengelola data dalam lembar kerja Excel dengan cara yang efisien. Program ini terdiri dari beberapa subrutin yang masing-masing memiliki fungsi spesifik, mulai dari membersihkan data, menyalin data antar lembar kerja, hingga membandingkan dan menyoroti perbedaan. Tujuan utama dari program ini adalah untuk memudahkan pengguna dalam mengelola data yang besar dan kompleks, serta memastikan bahwa data yang dihasilkan bersih dan terstruktur dengan baik.

### Sub RUN_COMPARE
- Fungsi: Subrutin ini berfungsi sebagai pengendali utama yang memanggil semua subrutin lain yang diperlukan untuk memproses data.
- Proses: Dengan memanggil subrutin lain secara berurutan, RUN_COMPARE memastikan bahwa semua langkah pemrosesan data dilakukan dalam urutan yang benar. Ini mencakup pembulatan nilai, penghapusan spasi, penyalinan data, pembersihan format, pemindahan data, pengisian detail, pewarnaan baris, perbandingan data, dan penambahan atau penghapusan bagian.

# Cara Menjalankan Program di Excel
1. Buka Excel: Pastikan Anda membuka aplikasi Microsoft Excel.

2. Buka Editor VBA:
    - Tekan ALT + F11 untuk membuka jendela Editor VBA.

3. Masukkan Kode:
    - Di jendela Editor VBA, klik Insert > Module untuk membuat modul baru.
    - Salin dan tempel kode VBA yang Anda berikan ke dalam modul tersebut.

4. Tutup Editor VBA: Setelah menempelkan kode, tutup jendela Editor VBA.

5. Kembali ke Excel: Kembali ke lembar kerja Excel.

6. Menjalankan Makro:
    - Tekan ALT + F8 untuk membuka dialog "Macro".
    - Pilih RUN_COMPARE dari daftar makro yang tersedia.
