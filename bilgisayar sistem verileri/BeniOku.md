# Windows Sistem Bilgi Toplayıcı

Bu program Windows işletim sisteminde çalışan bilgisayarların donanım ve sistem bilgilerini toplar ve Belgelerim klasörüne Excel dosyası olarak kaydeder.

## Kurulum ve Çalıştırma

1. Python'u bilgisayarınıza yükleyin (https://www.python.org/downloads/)
2. Gerekli kütüphaneleri yükleyin:
   ```
   npm run setup
   ```
3. Programı çalıştırın:
   ```
   npm start
   ```

## Program Çıktısı

Program çalıştığında aşağıdaki bilgileri toplayıp Belgelerim klasörüne Excel dosyası olarak kaydeder:

- Cihaz adı
- Bilgisayarın markası
- Bilgisayarın modeli
- Bilgisayarın seri numarası
- Wifi MAC adresi
- Ethernet MAC adresi
- İşletim sistemi adı
- İşletim sistemi sürümü
- Sistem türü (32/64 bit)
- CPU markası
- CPU modeli
- RAM miktarı
- Disk modeli
- Disk kapasitesi
- GPU markası

Excel dosyası "sistem_bilgileri_YYYYMMDD_HHMMSS.xlsx" formatında kaydedilir.