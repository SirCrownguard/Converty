# Converty

Converty is an open-source tool that converts **PDF to PPTX** and **PPTX to PDF**. It is currently available only for **Windows**, but in the future, support for **Linux and macOS may be added**.

![image](https://github.com/user-attachments/assets/899e0bb2-0d03-4eb1-b915-8783b74f35b5)


## How It Works
- Converts **PDF pages to PowerPoint slides** by extracting images from each page and inserting them into slides.
- Converts **PPTX slides to PDF** using PowerPoint automation or LibreOffice.
- Provides a **graphical user interface (GUI)** for easy file selection and conversion.

## Supported Platforms
- **Windows**: Installer available.
- **Linux/macOS**: Currently not supported, but the Python script can be executed manually.

---

## Installation & Usage

### Windows (Installer)
1. **[Download the latest version](https://github.com/SirCrownguard/Converty/releases/latest)**
2. Run `Converty_Setup.exe` and follow the installation steps.
3. Launch Converty from the Start Menu or Desktop shortcut.

### Linux/macOS (Manual Execution)
While official support is not available yet, you can still run the script manually:

1. Install Python 3.x and dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Run the script:
   ```
   python pdf_to_pptx.py
   ```

---

## How to Use
1. Open Converty.
2. Choose the conversion type.
3. Select your PDF or PPTX file.
4. Save the converted file.

---

## License
This project is licensed under **GNU General Public License v3.0 (GPLv3)**.  
[View the license](https://www.gnu.org/licenses/gpl-3.0.txt).

---

## Disclaimer
Converty is a **hobby project**. There is **no guarantee for updates, bug fixes, or new features**. I take no responsibility for any loss of files or anything else that happens. The program is open source and can be reviewed by anyone.

---

## Contributing
- Report bugs via **[Issues](https://github.com/SirCrownguard/Converty/issues)**.
- Submit new features via **Pull Requests**.
- Contributions are welcome, but updates are not guaranteed.

---

# Converty (Türkçe)

Converty, **PDF'i PPTX'e** ve **PPTX'i PDF'e** dönüştüren açık kaynaklı bir araçtır. Şu anda **yalnızca Windows için mevcuttur**, ancak gelecekte **Linux ve macOS desteği eklenebilir**.

## Nasıl Çalışır?
- **PDF sayfalarını PowerPoint slaytlarına dönüştürür**, her sayfanın görüntüsünü alarak PowerPoint'e ekler.
-  PowerPoint veya LibreOffice kullanarak **PPTX dosyalarını PDF'ye dönüştürür**.
- **Kullanıcı dostu bir arayüz** sağlar.

## Desteklenen Platformlar
- **Windows**: Yükleyici mevcuttur.
- **Linux/macOS**: Henüz desteklenmiyor, ancak Python betiği manuel olarak çalıştırılabilir.

---

## Kurulum ve Kullanım

### Windows (Yükleyici)
1. **[Son sürümü indir](https://github.com/SirCrownguard/Converty/releases/latest)**
2. `Converty_Setup.exe` dosyasını çalıştır ve yüklemeyi tamamla.
3. Başlat Menüsü veya Masaüstü kısayolundan aç.

### Linux/macOS (Manuel Çalıştırma)
Resmi destek henüz mevcut değil, ancak betik manuel olarak çalıştırılabilir:

1. Python 3.x ve bağımlılıkları yükle:
   ```
   pip install -r requirements.txt
   ```
2. Betiği çalıştır:
   ```
   python pdf_to_pptx.py
   ```

---

## Nasıl Kullanılır?
1. Converty'yi aç.
2. Dönüştürme türünü belirle.
3. PDF veya PPTX dosyanı seç.
4. Dönüştürülen dosyayı kaydet.

---

## Geliştiriciler İçin
### Gereksinimler
- Python 3.x
- Bağımlılıklar (`pip install -r requirements.txt` ile yüklenebilir)

### Yerelde Çalıştırma
```
python pdf_to_pptx.py
```

---

## Lisans
Bu proje **GNU General Public License v3.0 (GPLv3)** ile lisanslanmıştır.  
[Lisansı görüntüle](https://www.gnu.org/licenses/gpl-3.0.txt).

---

## Sorumluluk Reddi
Converty bir **hobi projesidir**. Güncellemeler, hata düzeltmeleri veya yeni özellikler **garanti edilmez**. Hiçbir dosya kaybından veya yaşanan herhangi bir şeyden sorumluluk kabul etmiyorum. Program açık kaynak kodludur ve herkes tarafından incelenebilir.

---

## Katkıda Bulun
- Hata bildirmek için **[Issues](https://github.com/SirCrownguard/Converty/issues)** kısmını kullanabilirsiniz.
- Yeni özellikler için **Pull Request** gönderebilirsin.
- Katkılar memnuniyetle karşılanır, ancak güncellemeler garanti edilmez.
