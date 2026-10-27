# Influencer Engagement Rate

Shortlist Excel'indeki her influencer için Instagram ve TikTok'taki **son 5 paylaşıma** bakarak engagement rate (ER) hesaplar. Veriler Apify'daki actor'lerle çekilir.

## Kurulum

```bash
pip install -r requirements.txt
```

Apify API token'ı: https://console.apify.com/settings/integrations

## Kullanım

**1. Önce hesapları kontrol et** (Apify çalışmaz, ücret çıkmaz):

```bash
python engagement.py "Samsung - Shortlist.xlsx" --sadece-linkler
```

Excel'de hangi satırda hangi Instagram / TikTok hesabının bulunduğunu listeler. Eksik ya da yanlış bulunan hesap varsa önce Excel'i düzelt.

**2. ER hesapla:**

```bash
export APIFY_TOKEN=apify_api_xxx
python engagement.py "Samsung - Shortlist.xlsx"
```

Çıktı: aynı klasörde `Samsung - Shortlist - ER.xlsx`
- Orijinal sayfaların sağına şu kolonlar eklenir: `IG Takipçi`, `IG Post`, `IG Toplam Etkileşim`, `IG ER`, aynıları TikTok için (`TT ...`) ve `Genel ER`.
- `Post Detay` sayfası: hesaplamaya giren her post (tarih, link, like, share, yorum). Rakamları buradan kontrol edebilirsin.

## Hesapları nasıl buluyor?

Her satırdaki tüm hücrelere bakar; hem hücre metnindeki hem de hücreye bağlı hyperlink'teki linkleri okur:
- Instagram: `instagram.com/<kullanıcı>`
- TikTok: `tiktok.com/@<kullanıcı>`

Kolon adı önemli değil. Sadece `@kullanıcı` yazılıp link verilmemişse veya kısa link (`vm.tiktok.com/...`) varsa hesap bulunamaz; bunları tam profil linkiyle değiştir.

## Hesaplama

Kullanılan Apify actor'leri (iki platform için toplam 2 çalıştırma, tüm hesaplar tek seferde):

| Platform | Actor | Takipçi | Like | Share | Etiketleme |
|---|---|---|---|---|---|
| Instagram | `apify/instagram-profile-scraper` | `followersCount` | `likesCount` | yok (Instagram public vermiyor) → 0 | `commentsCount` |
| TikTok | `clockworks/tiktok-scraper` | `authorMeta.fans` | `diggCount` | `shareCount` | `commentCount` |

**Post başına etkileşim**

```
Etkileşim = Like + Share + Etiketleme
```

- "Etiketleme" için platformların public bir sayacı yok. Etiketlemeler yorumlarda yapıldığı için en yakın karşılık olarak **yorum sayısı** kullanılıyor.
- Instagram'da like sayısı gizlenmiş postlarda like 0 sayılır.

**Platform ER** (son 5 post)

```
ER = Toplam Etkileşim / (Takipçi × Post Sayısı)
```

Toplam Etkileşim 5 postun toplamıdır. Takipçi, post sayısıyla çarpılıyor çünkü her post aynı takipçi kitlesine gösteriliyor; böylece sonuç "post başına ortalama ER" olur ve sektördeki %1–5 gibi benchmark'larla kıyaslanabilir. Hesapta 5'ten az post varsa olan postlarla hesaplanır.

Örnek: 10.000 takipçi, son 5 postta toplam 5.050 etkileşim → `5050 / (10000 × 5) = %10,1`

**Son 5 post nasıl seçiliyor?** Instagram'dan profilin son ~12 postu, TikTok'tan son 10 videosu çekilir, yayın tarihine göre sıralanıp en yeni 5'i alınır. Böylece en üste sabitlenmiş (pinned) eski postlar hesaba girmez.

**Genel ER** (influencer başına tek rakam)

```
Genel ER = (IG Toplam Etkileşim + TT Toplam Etkileşim) / (IG Takipçi × IG Post + TT Takipçi × TT Video)
```

Tek platformu olan influencer'da Genel ER o platformun ER'ine eşittir.

## Ayar değiştirmek

`engagement.py` dosyasının başında:
- `POST_SAYISI = 5` → kaç post alınacağı
- `IG_ALANLAR` / `TT_ALANLAR` → etkileşime hangi alanların gireceği (örn. TikTok'ta "Etiketleme" yerine kaydetme sayısını kullanmak için `"etiketleme": "collectCount"`)
