Author: Muhammad Fadil

Date : 2026-08-10

# Formula Excel

Materi Pembahasan:

1. Matematika & Statistik Dasar
2. Logika
3. Lookup & Referensi
4. Teks
5. Tanggal & Waktu
6. Fungsi Array Dinamis (Excel 365 / modern)
7. Informasi

---

## 1. Matematika & Statistik Dasar

| Fungsi        | Contoh Formula                                      | Penjelasan                                                |
| ------------- | --------------------------------------------------- | --------------------------------------------------------- |
| `SUM`         | `=SUM(A1:A10)`                                      | Menjumlahkan semua angka di A1 sampai A10                 |
| `SUMIF`       | `=SUMIF(B2:B10,"Jakarta",C2:C10)`                   | Jumlahkan C jika B = "Jakarta"                            |
| `SUMIFS`      | `=SUMIFS(C2:C10,B2:B10,"Jakarta",D2:D10,">100000")` | Jumlahkan C jika B="Jakarta" DAN D>100000                 |
| `AVERAGE`     | `=AVERAGE(A1:A10)`                                  | Rata-rata dari A1:A10                                     |
| `AVERAGEIF`   | `=AVERAGEIF(B2:B10,"Sales",C2:C10)`                 | Rata-rata C jika B = "Sales"                              |
| `AVERAGEIFS`  | `=AVERAGEIFS(C2:C10,B2:B10,"Sales",D2:D10,">0")`    | Rata-rata dengan banyak syarat                            |
| `COUNT`       | `=COUNT(A1:A10)`                                    | Hitung jumlah sel berisi angka                            |
| `COUNTA`      | `=COUNTA(A1:A10)`                                   | Hitung sel yang tidak kosong (angka & teks)               |
| `COUNTBLANK`  | `=COUNTBLANK(A1:A10)`                               | Hitung jumlah sel kosong                                  |
| `COUNTIF`     | `=COUNTIF(B2:B10,"Lulus")`                          | Hitung sel B yang isinya "Lulus"                          |
| `COUNTIFS`    | `=COUNTIFS(B2:B10,"Lulus",C2:C10,">=80")`           | Hitung dengan banyak syarat                               |
| `MAX`         | `=MAX(A1:A10)`                                      | Nilai tertinggi                                           |
| `MIN`         | `=MIN(A1:A10)`                                      | Nilai terendah                                            |
| `MAXIFS`      | `=MAXIFS(C2:C10,B2:B10,"Jakarta")`                  | Nilai tertinggi C dengan syarat B                         |
| `MINIFS`      | `=MINIFS(C2:C10,B2:B10,"Jakarta")`                  | Nilai terendah C dengan syarat B                          |
| `MEDIAN`      | `=MEDIAN(A1:A10)`                                   | Nilai tengah dari kumpulan data                           |
| `MODE.SNGL`   | `=MODE.SNGL(A1:A10)`                                | Nilai yang paling sering muncul                           |
| `ROUND`       | `=ROUND(12345.678,2)` → `12345.68`                  | Bulatkan ke 2 desimal                                     |
| `ROUNDUP`     | `=ROUNDUP(12.1,0)` → `13`                           | Bulatkan ke atas                                          |
| `ROUNDDOWN`   | `=ROUNDDOWN(12.9,0)` → `12`                         | Bulatkan ke bawah                                         |
| `ABS`         | `=ABS(-50)` → `50`                                  | Nilai absolut (positif)                                   |
| `POWER`       | `=POWER(2,3)` → `8`                                 | 2 pangkat 3                                               |
| `SQRT`        | `=SQRT(64)` → `8`                                   | Akar kuadrat                                              |
| `PRODUCT`     | `=PRODUCT(A1:A3)`                                   | Kalikan semua nilai A1 sampai A3                          |
| `MOD`         | `=MOD(10,3)` → `1`                                  | Sisa hasil bagi 10 dibagi 3                               |
| `RAND`        | `=RAND()`                                           | Angka acak antara 0 dan 1                                 |
| `RANDBETWEEN` | `=RANDBETWEEN(1,100)`                               | Angka acak bulat antara 1-100                             |
| `SUMPRODUCT`  | `=SUMPRODUCT(A1:A5,B1:B5)`                          | Kalikan tiap pasangan lalu jumlahkan (mis. Qty × Harga)   |
| `SUBTOTAL`    | `=SUBTOTAL(9,A1:A10)`                               | Total (kode 9 = SUM), abaikan baris yang di-filter/hidden |

---
