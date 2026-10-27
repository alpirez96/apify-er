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

Her influencer için bulunan Instagram / TikTok kullanıcı adını listeler. Aynı hesap birden fazla satırda geçiyorsa uyarı verir. Yanlış link varsa önce Excel'i düzelt.

**2. ER hesapla:**

```bash
export APIFY_TOKEN=apify_api_xxx
python engagement.py "Samsung - Shortlist.xlsx"
```

Çıktı: aynı klasörde `Samsung - Shortlist - ER.xlsx` (orijinal dosyaya dokunulmaz)
- Shortlist'teki `ER%` kolonları doldurulur: **C = Instagram ER**, **E = TikTok ER**. Apify'dan verisi gelmeyen hesapta hücredeki mevcut değer olduğu gibi kalır; script bu hesapları sonda listeler.
- `Post Detay` sayfası: hesaplamaya giren her post (tarih, link, like, share, yorum). Rakamları buradan kontrol edebilirsin.

## Excel düzeni

Script ilk sayfayı okur. 2. satır başlık, sonraki satırlar influencer:

| A | B | C | D | E |
|---|---|---|---|---|
| Influencer | IG (hücreye profil linki bağlı) | ER% | TT (hücreye profil linki bağlı) | ER% |

Kullanıcı adı, hücreye bağlı hyperlink'ten (`instagram.com/<kullanıcı>`, `tiktok.com/@<kullanıcı>`) okunur. Link olmayan (`-`) hücreler atlanır. Kolonlar değişirse `engagement.py` başındaki `IG_KOLON`, `IG_ER_KOLON`, `TT_KOLON`, `TT_ER_KOLON` sabitlerini güncelle.

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

Toplam Etkileşim, 5 postun etkileşimlerinin toplamıdır. Bu, shortlist'teki notla aynı şeyi söyler: "Toplam Etkileşim / Toplam Takipçi, ortalama son 4-5 içerik bakılarak". Takipçi sayısı post sayısıyla çarpılır, böylece sonuç **post başına ortalama ER** olur. Hesapta 5'ten az post varsa mevcut postlarla hesaplanır.

Örnek: 837.000 takipçi, son 5 postta toplam 52.500 etkileşim → `52500 / (837000 × 5) = %1,25`

Takipçi sayısı Apify'dan anlık çekilir; Excel'deki B/D kolonlarındaki rakamlar kullanılmaz.

**Son 5 post nasıl seçiliyor?** Instagram'dan profilin son ~12 postu, TikTok'tan son 10 videosu çekilir. Bunlar yayın tarihine göre sıralanır ve en yeni 5'i alınır. Böylece en üste sabitlenmiş (pinned) eski postlar hesaba girmez.

## Ayar değiştirmek

`engagement.py` dosyasının başında:
- `POST_SAYISI = 5` → kaç post alınacağı
- `IG_ALANLAR` / `TT_ALANLAR` → etkileşime hangi alanların gireceği (örn. TikTok'ta "Etiketleme" yerine kaydetme sayısını kullanmak için `"etiketleme": "collectCount"`)
