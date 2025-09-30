# Payzgate Veri Çekme Botu

Bu Python scripti Payzgate sitesinden spesifik sipariş verilerini çekmek için tasarlanmıştır.

## Çekilen Veriler

Bot aşağıdaki spesifik verileri çeker:
- **Status**: Sipariş durumu (Pending, Processing, Complete, Cancelled)
- **ID**: Sipariş numarası
- **Customer**: Müşteri adı
- **Order Subtotal**: Sipariş alt toplamı
- **Order Total**: Sipariş toplamı
- **Reference Code**: Referans kodu
- **IBAN**: IBAN bilgisi

## Çoklu Sipariş Desteği

Bot birden fazla sipariş ID'sini aynı anda işleyebilir:

### ID Giriş Formatları:
- **Tek ID**: `6172936`
- **Çoklu ID**: `6172936,6172937,6172938`
- **Aralık**: `6172936-6172940` (6172936'dan 6172940'a kadar)
- **Karışık**: `6172936,6172938-6172940,6172945`

### Özellikler:
- ✅ Tek Excel dosyasında tüm siparişler
- ✅ Her sipariş için ayrı satır
- ✅ Başarılı/başarısız özet raporu
- ✅ Otomatik sıralama ve duplikat temizleme

## Özellikler

- Payzgate sitesinden spesifik sipariş verilerini çeker
- Login desteği (email/şifre ile)
- Verileri JSON ve CSV formatında kaydeder
- Selenium ile dinamik içerik desteği
- BeautifulSoup ile HTML parsing
- Spesifik veri çıkarma algoritması

## Kurulum

1. Gerekli kütüphaneleri yükleyin:
```bash
pip install -r requirements.txt
```

2. Scripti çalıştırın:
```bash
python3 main.py
```

## Kullanım

### Seçenek 1: Gerçek Login ile
```bash
python3 main.py
```
1. Email adresinizi girin
2. Şifrenizi girin
3. Script otomatik olarak sipariş verilerini çekecek

### Seçenek 2: Test Login ile
```bash
python3 test_with_login.py
```
Gerçek Payzgate hesap bilgilerinizi girin

### Seçenek 3: Login Olmadan (Veri Çekilemez)
```bash
python3 simple_test.py
```
Login olmadan test eder, Excel'de "VERİ ÇEKİLEMEDİ" yazacak

### Seçenek 4: Sabit Bilgilerle Test
```bash
python3 test_with_credentials.py
```
Script içindeki email/şifre bilgilerini düzenleyip kullanın

### Seçenek 5: Gerçek Login Testi
```bash
python3 test_real_login.py
```
Gerçek bilgilerinizi script içine yazıp test edin

### Seçenek 6: Çoklu Sipariş (Önerilen)
```bash
python3 multiple_orders.py
```
Birden fazla sipariş ID'si girip tek Excel'de toplayın

### Seçenek 7: Çoklu Sipariş Test
```bash
python3 test_multiple.py
```
Script içindeki ID listesini düzenleyip test edin

## Test

Mock data ile test etmek için:
```bash
python3 test_mock.py
```

## Çıktı Dosyaları

- `order_[ID]_[timestamp].json`: Ham veri (JSON formatında)
- `order_[ID]_[timestamp].xlsx`: Spesifik veriler (Excel formatında - tek formda)
  - **Başlık satırı**: Sipariş ID, URL, Çekilme Tarihi, Status, ID, Customer, Order Subtotal, Order Total, Reference Code, IBAN
  - **Veri satırı**: İlgili bilgiler

## Örnek Çıktı

```
=== SİPARİŞ VERİLERİ ===
Status: Pending
ID: 6173023
Customer: Mehmet emin Koç
Order Subtotal: 490,00 ₺ excl tax
Order Total: 490,00 ₺
Reference Code: 5nsrpip3-e4ud-if9o-nqly-q9m6ogka1otq
IBAN: Garanti TR700005902580130258013225 MYPAYZ ÖDEME KURULUSU A.S.
```

## Notlar

- Site login gerektiriyorsa email ve şifre girmeniz gerekebilir
- Chrome WebDriver otomatik olarak indirilir
- Script headless modda çalışır (tarayıcı penceresi açılmaz)
- Mock test ile scraper'ın doğru çalıştığı doğrulanmıştır

## Gereksinimler

- Python 3.6+
- Chrome tarayıcısı
- İnternet bağlantısı
- Payzgate hesap bilgileri (login için)
