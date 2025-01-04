# Not Hesaplama
## [İndirme Linki](https://raw.githubusercontent.com/omerdynasty/not-hesaplama/refs/heads/main/Excel%20Dosyalar%C4%B1/not%20hesaplama.xlsm)

### [VirusTotal](https://www.virustotal.com/gui/file/34067ca54d87b47033d05c55667617567276b66f6e1576abf29ea06279a77014)

**Excel[^1] dosyasının içindeki yönergeleri okumayı unutmayın**
---


### **1. Sub Hesapla()**
Bu alt prosedür, sütun B'deki sayısal değerleri alır ve her birini bir formüle göre hesaplayarak sütun C'ye yazar.

#### Kodun Parçalarına Ayrılması:
```vba
Dim i As Long
Dim sonSatir As Long
```
- **`Dim i As Long`**: `i` değişkeni, bir sayaç olarak kullanılır ve satırları döngüyle dolaşır. Uzun tamsayı (`Long`) veri tipindedir.
- **`Dim sonSatir As Long`**: Sütun B'deki son dolu satırın numarasını saklamak için kullanılan değişkendir.

---

```vba
sonSatir = Cells(Rows.Count, 2).End(xlUp).Row
```
- **Amacı:** Sütun B'nin son dolu satırını bulur.
- **Nasıl Çalışır?**
  1. `Rows.Count`: Çalışma sayfasındaki toplam satır sayısını alır (örneğin, Excel 2019[^1]'da bu 1,048,576'dır).
  2. `Cells(Rows.Count, 2)`: Sütun B'nin en alt hücresini seçer.
  3. `.End(xlUp)`: Bu hücreden yukarı doğru dolu olan ilk hücreye gider.
  4. `.Row`: Bulunan hücrenin satır numarasını alır ve `sonSatir` değişkenine atar.

Örneğin:
- Eğer sütun B'nin son dolu hücresi B10 ise, `sonSatir = 10` olur.

---

```vba
For i = 1 To sonSatir
```
- Bu döngü, 1'den `sonSatir` değişkenine kadar her bir satır için işlem yapar. Yani sütun B'deki tüm dolu hücreler üzerinde işlem gerçekleştirilir.

---

```vba
If IsNumeric(Cells(i, 2).Value) Then
```
- **Amacı:** Sütun B'deki hücrenin içeriğinin sayısal bir değer olup olmadığını kontrol eder.
- **`IsNumeric` Fonksiyonu:** Eğer hücre sayısal bir değer içeriyorsa `True` döner, aksi halde `False`.

---

```vba
Cells(i, 3).Value = Application.WorksheetFunction.Round((Cells(i, 2).Value * 100) / 60, 0)
```
- **Hesaplama:**
  1. **`Cells(i, 2).Value`**: Sütun B'deki ilgili hücrenin değerini alır.
  2. **`(Cells(i, 2).Value * 100) / 60`**: Bu değeri 100 ile çarpıp 60'a böler.
  3. **`Application.WorksheetFunction.Round(..., 0)`**: Sonucu en yakın tam sayıya yuvarlar.
  4. **`Cells(i, 3).Value`**: Hesaplanan değeri sütun C'nin aynı satırına yazar.

---

```vba
Next i
```
- Döngüyü bir sonraki satıra geçirir. İşlem `sonSatir` değerine ulaşana kadar devam eder.

---

### **2. Sub Temizle()**
Bu alt prosedür, sütun B ve C'deki tüm verileri temizler.

#### Kodun Parçalarına Ayrılması:
```vba
Columns("B:C").ClearContents
```
- **`Columns("B:C")`**: Sütun B ve C'yi seçer.
- **`.ClearContents`**: Bu sütunların içeriklerini temizler. Ancak hücre biçimlendirmelerine (örneğin, renk veya kenarlık) dokunmaz.

---

### **Kodun İşleyişi (Genel Akış):**
1. **Hesapla Makrosu:**
   - Sütun B'deki tüm dolu satırları kontrol eder.
   - Eğer bir hücre sayısal bir değer içeriyorsa, o değeri `(B sütunundaki değer * 100) / 60` formülüyle işler.
   - Sonucu yuvarlayarak aynı satırdaki sütun C'ye yazar.

2. **Temizle Makrosu:**
   - Sütun B ve C'deki tüm verileri temizler.

---

### **Örnek:**
| B (Girdi) | C (Sonuç) |
|-----------|-----------|
| 30        | 50        |
| 45        | 75        |
| abc       |           |
| 60        | 100       |

Eğer **"Temizle"** makrosunu çalıştırırsanız, B ve C sütunları tamamen boş olur.

[^1]: Excel, Microsoft Corporation'ın tescilli bir markasıdır.
