Author: Muhammad Fadil

Date : 2026-08-10

# Formula Excel

Materi Pembahasan:

1. Matematika & Statistik Dasar
2. Logika
3. Lookup & Referensi
4. Teks
5. Tanggal & Waktu
6. Fungsi Array Dinamis (Excel terbaru)
7. Informasi

---

Link data: [download data](https://docs.google.com/spreadsheets/d/1z_ZlMxsfP1OmaNKJuu2aQqxC6k5a3il51VElDeZVsv0)

---

![1. Matematika & Statistik Dasar](img/01.png)

| Fungsi        | Contoh Formula                                       | Penjelasan                                                |
| ------------- | ---------------------------------------------------- | --------------------------------------------------------- |
| `SUM`         | `'=SUM(A1:A10)`                                      | Menjumlahkan semua angka di A1 sampai A10                 |
| `SUMIF`       | `'=SUMIF(B2:B10,"Jakarta",C2:C10)`                   | Jumlahkan C jika B = "Jakarta"                            |
| `SUMIFS`      | `'=SUMIFS(C2:C10,B2:B10,"Jakarta",D2:D10,">100000")` | Jumlahkan C jika B="Jakarta" DAN D>100000                 |
| `AVERAGE`     | `'=AVERAGE(A1:A10)`                                  | Rata-rata dari A1:A10                                     |
| `AVERAGEIF`   | `'=AVERAGEIF(B2:B10,"Sales",C2:C10)`                 | Rata-rata C jika B = "Sales"                              |
| `AVERAGEIFS`  | `'=AVERAGEIFS(C2:C10,B2:B10,"Sales",D2:D10,">0")`    | Rata-rata dengan banyak syarat                            |
| `COUNT`       | `'=COUNT(A1:A10)`                                    | Hitung jumlah sel berisi angka                            |
| `COUNTA`      | `'=COUNTA(A1:A10)`                                   | Hitung sel yang tidak kosong (angka & teks)               |
| `COUNTBLANK`  | `'=COUNTBLANK(A1:A10)`                               | Hitung jumlah sel kosong                                  |
| `COUNTIF`     | `'=COUNTIF(B2:B10,"Lulus")`                          | Hitung sel B yang isinya "Lulus"                          |
| `COUNTIFS`    | `'=COUNTIFS(B2:B10,"Lulus",C2:C10,">=80")`           | Hitung dengan banyak syarat                               |
| `MAX`         | `'=MAX(A1:A10)`                                      | Nilai tertinggi                                           |
| `MIN`         | `'=MIN(A1:A10)`                                      | Nilai terendah                                            |
| `MAXIFS`      | `'=MAXIFS(C2:C10,B2:B10,"Jakarta")`                  | Nilai tertinggi C dengan syarat B                         |
| `MINIFS`      | `'=MINIFS(C2:C10,B2:B10,"Jakarta")`                  | Nilai terendah C dengan syarat B                          |
| `MEDIAN`      | `'=MEDIAN(A1:A10)`                                   | Nilai tengah dari kumpulan data                           |
| `MODE.SNGL`   | `'=MODE.SNGL(A1:A10)`                                | Nilai yang paling sering muncul                           |
| `ROUND`       | `'=ROUND(12345.678,2)` → `12345.68`                  | Bulatkan ke 2 desimal                                     |
| `ROUNDUP`     | `'=ROUNDUP(12.1,0)` → `13`                           | Bulatkan ke atas                                          |
| `ROUNDDOWN`   | `'=ROUNDDOWN(12.9,0)` → `12`                         | Bulatkan ke bawah                                         |
| `ABS`         | `'=ABS(-50)` → `50`                                  | Nilai absolut (positif)                                   |
| `POWER`       | `'=POWER(2,3)` → `8`                                 | 2 pangkat 3                                               |
| `SQRT`        | `'=SQRT(64)` → `8`                                   | Akar kuadrat                                              |
| `PRODUCT`     | `'=PRODUCT(A1:A3)`                                   | Kalikan semua nilai A1 sampai A3                          |
| `MOD`         | `'=MOD(10,3)` → `1`                                  | Sisa hasil bagi 10 dibagi 3                               |
| `RAND`        | `'=RAND()`                                           | Angka acak antara 0 dan 1                                 |
| `RANDBETWEEN` | `'=RANDBETWEEN(1,100)`                               | Angka acak bulat antara 1-100                             |
| `SUMPRODUCT`  | `'=SUMPRODUCT(A1:A5,B1:B5)`                          | Kalikan tiap pasangan lalu jumlahkan (mis. Qty × Harga)   |
| `SUBTOTAL`    | `'=SUBTOTAL(9,A1:A10)`                               | Total (kode 9 = SUM), abaikan baris yang di-filter/hidden |

**Contoh:**

```
=SUM(A1:A10)
```

---

![2. Logika](img/02.png)

| Fungsi         | Contoh Formula                                                 | Penjelasan                                      |
| -------------- | -------------------------------------------------------------- | ----------------------------------------------- |
| `IF`           | `'=IF(A1>=75,"Lulus","Tidak Lulus")`                           | Jika nilai A1 ≥75 maka "Lulus"                  |
| `IFS`          | `'=IFS(A1>=90,"A",A1>=75,"B",A1>=60,"C",TRUE,"D")`             | Banyak kondisi sekaligus                        |
| `AND`          | `'=AND(A1>60,B1="Aktif")` → `TRUE`/`FALSE`                     | Benar jika SEMUA kondisi terpenuhi              |
| `OR`           | `'=OR(A1="Ya",B1="Ya")` → `TRUE`/`FALSE`                       | Benar jika SALAH SATU kondisi terpenuhi         |
| `NOT`          | `'=NOT(A1="Selesai")`                                          | Kebalikan dari kondisi                          |
| `IFERROR`      | `'=IFERROR(A1/B1,"Error")`                                     | Tampilkan "Error" jika rumus menghasilkan error |
| `IFNA`         | `'=IFNA(VLOOKUP(A1,C:D,2,0),"Tidak Ditemukan")`                | Tangani khusus error #N/A                       |
| `SWITCH`       | `'=SWITCH(A1,1,"Senin",2,"Selasa",3,"Rabu","Tidak Diketahui")` | Cocokkan nilai A1 ke daftar kasus               |
| `TRUE`/`FALSE` | `'=IF(A1=TRUE,"Ya","Tidak")`                                   | Nilai boolean langsung                          |

**Contoh gabungan (IF bertingkat + AND):**

```
=IF(AND(A1>=80,B1="Hadir"),"Lolos","Tidak Lolos")
```

---

![3. Lookup & Referensi](img/03.png)

| Fungsi        | Contoh Formula                                      | Penjelasan                                                                  |
| ------------- | --------------------------------------------------- | --------------------------------------------------------------------------- |
| `VLOOKUP`     | `'=VLOOKUP(A1,D:F,3,FALSE)`                         | Cari A1 di kolom D, ambil hasil dari kolom ke-3 (F), exact match            |
| `HLOOKUP`     | `'=HLOOKUP(A1,D1:H5,3,FALSE)`                       | Sama seperti VLOOKUP tapi mencari secara horizontal                         |
| `XLOOKUP`     | `'=XLOOKUP(A1,D:D,F:F,"Tidak ada")`                 | Cari A1 di kolom D, ambil hasil dari kolom F (lebih fleksibel dari VLOOKUP) |
| `INDEX`       | `'=INDEX(F2:F10,3)`                                 | Ambil nilai baris ke-3 dari range F2:F10                                    |
| `MATCH`       | `'=MATCH("Budi",A2:A10,0)`                          | Cari posisi baris "Budi" di A2:A10                                          |
| `INDEX+MATCH` | `'=INDEX(F2:F10,MATCH(A1,D2:D10,0))`                | Cari nilai A1 di D, ambil hasil sebaris di F (fleksibel ke kiri/kanan)      |
| `LOOKUP`      | `'=LOOKUP(A1,D2:D10,F2:F10)`                        | Versi sederhana, data harus terurut                                         |
| `CHOOSE`      | `'=CHOOSE(2,"Merah","Kuning","Hijau")` → `"Kuning"` | Pilih nilai ke-2 dari daftar                                                |
| `OFFSET`      | `'=OFFSET(A1,2,1)`                                  | Referensi sel 2 baris ke bawah, 1 kolom ke kanan dari A1                    |
| `INDIRECT`    | `'=INDIRECT("A1")`                                  | Ubah teks "A1" menjadi referensi sel A1                                     |
| `ROW`         | `'=ROW(A5)` → `5`                                   | Nomor baris dari sel A5                                                     |
| `COLUMN`      | `'=COLUMN(C1)` → `3`                                | Nomor kolom dari sel C1                                                     |
| `ROWS`        | `'=ROWS(A1:A10)` → `10`                             | Jumlah baris dalam range                                                    |
| `COLUMNS`     | `'=COLUMNS(A1:D1)` → `4`                            | Jumlah kolom dalam range                                                    |
| `TRANSPOSE`   | `'=TRANSPOSE(A1:C1)`                                | Ubah data baris menjadi kolom (atau sebaliknya)                             |

---

![4. Teks](img/04.png)

| Fungsi        | Contoh Formula                                        | Penjelasan                                                       |
| ------------- | ----------------------------------------------------- | ---------------------------------------------------------------- |
| `CONCATENATE` | `'=CONCATENATE(A1," ",B1)`                            | Gabungkan A1 dan B1 dengan spasi                                 |
| `CONCAT`      | `'=CONCAT(A1:A5)`                                     | Versi modern CONCATENATE, bisa range                             |
| `TEXTJOIN`    | `'=TEXTJOIN(", ",TRUE,A1:A5)`                         | Gabungkan A1:A5 dipisah koma, abaikan sel kosong                 |
| `LEFT`        | `'=LEFT("Excel Indonesia",5)` → `"Excel"`             | Ambil 5 karakter dari kiri                                       |
| `RIGHT`       | `'=RIGHT("Excel Indonesia",9)` → `"Indonesia"`        | Ambil 9 karakter dari kanan                                      |
| `MID`         | `'=MID("Excel Indonesia",7,3)` → `"Ind"`              | Ambil 3 karakter mulai posisi ke-7                               |
| `LEN`         | `'=LEN("Excel")` → `5`                                | Jumlah karakter dalam teks                                       |
| `TRIM`        | `'=TRIM("  Excel   ")` → `"Excel"`                    | Hapus spasi berlebih                                             |
| `UPPER`       | `'=UPPER("excel")` → `"EXCEL"`                        | Ubah jadi huruf besar                                            |
| `LOWER`       | `'=LOWER("EXCEL")` → `"excel"`                        | Ubah jadi huruf kecil                                            |
| `PROPER`      | `'=PROPER("budi santoso")` → `"Budi Santoso"`         | Kapital di tiap awal kata                                        |
| `SUBSTITUTE`  | `'=SUBSTITUTE("2024-01-01","-","/")` → `"2024/01/01"` | Ganti karakter tertentu                                          |
| `REPLACE`     | `'=REPLACE("Excel2024",6,4,"2025")` → `"Excel2025"`   | Ganti berdasarkan posisi karakter                                |
| `FIND`        | `'=FIND("l","Excel")` → `4`                           | Posisi huruf "l" (case-sensitive)                                |
| `SEARCH`      | `'=SEARCH("l","EXCEL")` → `4`                         | Sama seperti FIND tapi tidak case-sensitive                      |
| `TEXT`        | `'=TEXT(1234567,"#,##0")` → `"1,234,567"`             | Ubah angka jadi teks berformat                                   |
| `VALUE`       | `'=VALUE("123")` → `123`                              | Ubah teks angka jadi angka                                       |
| `REPT`        | `'=REPT("-",10)` → `"----------"`                     | Ulangi teks sebanyak n kali                                      |
| `CLEAN`       | `'=CLEAN(A1)`                                         | Hapus karakter non-printable (sering dari data hasil copy-paste) |
| `EXACT`       | `'=EXACT("Excel","excel")` → `FALSE`                  | Bandingkan teks persis sama (case-sensitive)                     |

---

![5. Tanggal & Waktu](img/05.png)

| Fungsi        | Contoh Formula                     | Penjelasan                                             |
| ------------- | ---------------------------------- | ------------------------------------------------------ |
| `TODAY`       | `'=TODAY()`                        | Tanggal hari ini                                       |
| `NOW`         | `'=NOW()`                          | Tanggal & waktu saat ini                               |
| `DATE`        | `'=DATE(2026,8,25)` → `25/08/2026` | Buat tanggal dari tahun, bulan, hari                   |
| `DATEDIF`     | `'=DATEDIF(A1,B1,"Y")`             | Selisih tahun antara dua tanggal ("M"=bulan, "D"=hari) |
| `DAYS`        | `'=DAYS(B1,A1)`                    | Jumlah hari antara dua tanggal                         |
| `YEAR`        | `'=YEAR(A1)`                       | Ambil tahun dari tanggal A1                            |
| `MONTH`       | `'=MONTH(A1)`                      | Ambil bulan dari tanggal A1                            |
| `DAY`         | `'=DAY(A1)`                        | Ambil hari dari tanggal A1                             |
| `WEEKDAY`     | `'=WEEKDAY(A1)` → `1-7`            | Hari ke berapa dalam seminggu                          |
| `WORKDAY`     | `'=WORKDAY(A1,10)`                 | Tanggal 10 hari kerja setelah A1 (skip Sabtu/Minggu)   |
| `NETWORKDAYS` | `'=NETWORKDAYS(A1,B1)`             | Jumlah hari kerja antara dua tanggal                   |
| `EDATE`       | `'=EDATE(A1,3)`                    | Tanggal 3 bulan setelah A1                             |
| `EOMONTH`     | `'=EOMONTH(A1,0)`                  | Tanggal akhir bulan dari A1                            |

---

![6. Fungsi Array Dinamis](img/06.png)

| Fungsi       | Contoh Formula                      | Penjelasan                                                       |
| ------------ | ----------------------------------- | ---------------------------------------------------------------- |
| `FILTER`     | `'=FILTER(A2:C10,B2:B10="Jakarta")` | Tampilkan baris dimana kolom B = "Jakarta"                       |
| `SORT`       | `'=SORT(A2:B10,2,-1)`               | Urutkan berdasarkan kolom 2, dari besar ke kecil                 |
| `SORTBY`     | `'=SORTBY(A2:A10,B2:B10,1)`         | Urutkan A berdasarkan urutan nilai B (ascending)                 |
| `UNIQUE`     | `'=UNIQUE(A2:A100)`                 | Ambil daftar nilai unik (tanpa duplikat)                         |
| `SEQUENCE`   | `'=SEQUENCE(5,1,1,1)` → `1,2,3,4,5` | Buat urutan angka otomatis                                       |
| `LET`        | `'=LET(x,A1*2,y,B1*3,x+y)`          | Buat variabel sementara dalam satu formula                       |
| `LAMBDA`     | `'=LAMBDA(a,b,a+b)(5,10)` → `15`    | Buat fungsi custom sendiri                                       |
| `TEXTSPLIT`  | `'=TEXTSPLIT(A1,",")`               | Pecah teks menjadi beberapa sel berdasarkan pemisah              |
| `TEXTBEFORE` | `'=TEXTBEFORE(A1,"@")`              | Ambil teks sebelum karakter tertentu (mis. sebelum "@" di email) |
| `TEXTAFTER`  | `'=TEXTAFTER(A1,"@")`               | Ambil teks setelah karakter tertentu                             |

---

![7. Informasi](img/07.png)

| Fungsi     | Contoh Formula                              | Penjelasan                             |
| ---------- | ------------------------------------------- | -------------------------------------- |
| `ISBLANK`  | `'=ISBLANK(A1)` → `TRUE`/`FALSE`            | Cek apakah sel kosong                  |
| `ISERROR`  | `'=ISERROR(A1/B1)` → `TRUE`/`FALSE`         | Cek apakah rumus menghasilkan error    |
| `ISNA`     | `'=ISNA(VLOOKUP(A1,C:D,2,0))`               | Cek apakah hasilnya error #N/A         |
| `ISNUMBER` | `'=ISNUMBER(A1)` → `TRUE`/`FALSE`           | Cek apakah isi sel berupa angka        |
| `ISTEXT`   | `'=ISTEXT(A1)` → `TRUE`/`FALSE`             | Cek apakah isi sel berupa teks         |
| `CELL`     | `'=CELL("address",A1)` → `"$A$1"`           | Info tentang sel (alamat, format, dll) |
| `TYPE`     | `'=TYPE(A1)` → `1` (angka), `2` (teks), dst | Tipe data dari nilai sel               |

---
