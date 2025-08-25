# javareport: Pemroses Laporan Pass Masuk GAC

Proyek Java sederhana untuk memproses data pass masuk dari file Excel (.xlsx) menggunakan library Apache POI secara manual, tanpa build tool seperti Maven atau Gradle. Program ini membaca file Excel input, mengisi data ke template, dan menghasilkan file laporan baru.

## Struktur Proyek
javareport/
├── src/
│   ├── Main.java             # Titik masuk utama program
│   └── UploadProcessor.java  # Logika inti untuk memproses Excel
├── lib/                      # Folder untuk semua library Apache POI (.jar)
├── out/                      # Folder untuk hasil kompilasi (.class files)
├── template/
│   └── Report Pass Masuk GAC.xlsx # File template Excel yang akan diisi
├── data.xlsx                 # File input Excel yang akan diproses
└── README.md
## Persyaratan (Requirements)

* **Java Development Kit (JDK)** 8 atau versi yang lebih baru.
* **Library Apache POI (.jar files)**: Semua file yang relevan (seperti `poi-*.jar`, `poi-ooxml-*.jar`, `xmlbeans-*.jar`, `commons-compress-*.jar`, dan lain-lain) harus ditempatkan di folder `lib/`. Untuk menghilangkan pesan `ERROR StatusLogger`, pastikan juga `log4j-core-*.jar` ada di folder ini.

## Cara Menjalankan

1.  **Siapkan File Input dan Template:**
    * Pastikan file Excel yang akan diproses (`data.xlsx`) sudah ada di folder root proyek (`javareport/`).
    * Pastikan file template (`Report Pass Masuk GAC.xlsx`) sudah ada di folder `javareport/template/`.

2.  **Kompilasi Proyek** dari terminal atau Command Prompt, pastikan Anda berada di dalam folder `javareport/`.

    * **Untuk Windows:**
        ```bash
        javac -d out -cp "lib/*" src/Main.java src/UploadProcessor.java
        ```
    * **Untuk macOS/Linux:**
        ```bash
        javac -d out -cp "lib/*" src/Main.java src/UploadProcessor.java
        ```

3.  **Jalankan Program** dari terminal.

    * **Untuk Windows:**
        ```bash
        java -cp "out;lib/*" Main
        ```
    * **Untuk macOS/Linux:**
        ```bash
        java -cp "out:lib/*" Main
        ```

Setelah program selesai berjalan, Anda akan menemukan file laporan Excel baru bernama `Report_Pass_Masuk_GAC_Output.xlsx` di folder root proyek Anda.