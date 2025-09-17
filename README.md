# PDKS İşleme Sistemi

Bu sistem, PDKS Excel dosyalarını okuyup günlük çalışma verilerini işleyerek JSON ve CSV çıktıları üretir. Modern Electron arayüzü ile kullanımı kolaydır.

## Özellikler

- 🖥️ **Modern Electron Arayüzü**: Kullanıcı dostu grafik arayüz
- 📁 **Dosya Seçme**: Excel dosyalarını kolayca seçme
- 📊 **Otomatik İşleme**: Excel dosyalarını otomatik okuma ve işleme
- 👥 **Personel Yönetimi**: Personel bloklarını otomatik ayırma
- ⏰ **Vardiya Sistemi**: V1, V2, V3 vardiyalarını otomatik belirleme
- 💰 **Fazla Mesai**: 20dk tolerans ile fazla mesai hesaplama
- ⏱️ **Çalışma Süresi**: 30dk mola düş ile çalışma süresi hesaplama
- 📄 **Çoklu Çıktı**: JSON ve CSV formatında rapor üretme
- 🎨 **Modern Tasarım**: Responsive ve kullanıcı dostu arayüz

## Kurulum

```bash
npm install
```

## Kullanım

### Electron Arayüzü (Önerilen)
```bash
npm start
```

### Komut Satırı
```bash
npm run process
```

## Çıktılar

- `out/daily.json` - Detaylı JSON verisi
- `out/daily.csv` - CSV formatında tablo verisi

## Vardiya Tanımları

- **V1**: 08:30-16:30 (Giriş penceresi: 06:00-12:00)
- **V2**: 16:30-00:30 (Giriş penceresi: 14:00-20:00)
- **V3**: 00:30-08:30 (Giriş penceresi: 22:00-04:00)

## Hesaplama Kuralları

- Planlı çalışma süresi: 7,5 saat (8 saat - 30dk mola)
- Fazla mesai toleransı: 20 dakika
- Gece taşmaları otomatik hesaplanır
- Eksik giriş/çıkış kayıtları "eksik" olarak işaretlenir

## Veri Yapısı

### JSON Çıktısı
```json
{
  "personel": "ABDULLAH BİNİCİ",
  "tarih": "2025-09-01",
  "ic_dis": "10:05",
  "vardiya": {
    "kod": "V1",
    "plan_bas": "08:30",
    "plan_bit": "16:30"
  },
  "gercek": {
    "gir": "08:28",
    "cik": "18:33"
  },
  "calisma_dk": 575,
  "fm_dk": 103,
  "fm_hhmm": "01:43",
  "durum": "ok",
  "not": ""
}
```

## Gereksinimler

- Node.js
- xlsx paketi
