---
name: influencer-engagement-rate
description: Influencer shortlist Excel'inden Instagram ve TikTok engagement rate (ER) raporu üretir. Profil linklerini okur, yanlış girilmiş linkleri bulup düzeltir, Apify ile güncel takipçi sayısını ve son 5 içeriği çeker, ER'yi hesaplar (Instagram = (Beğeni + Yorum) / Takipçi, TikTok = (Beğeni + Yorum + Paylaşım) / Takipçi) ve formüllerin yazdığı tasarımlı one-pager Excel raporu çıkarır. Kullanıcı bir marka için influencer/creator shortlist'i paylaştığında; engagement rate, ER%, etkileşim oranı, takipçi güncelleme, influencer link kontrolü, kampanya için influencer karşılaştırması veya "bu listedeki influencer'ların ER'lerini çıkar" gibi bir şey istediğinde, "engagement" kelimesi geçmese bile bu skill'i kullan.
compatibility: Python 3.10+, apify-client ve openpyxl (requirements.txt). APIFY_TOKEN ortam değişkeni ve api.apify.com'a ağ erişimi gerekir.
---

# Influencer Engagement Rate Raporu

Markalar için hazırlanan influencer shortlist'lerinde (ör. "Samsung - Shortlist.xlsx") her influencer'ın Instagram ve TikTok ER'sini aynı yöntemle, güncel veriyle hesaplayıp okunaklı bir rapora dönüştürür. Amaç, influencer'ları kıyaslanabilir tek bir ölçüyle yan yana koymak. Yanlış girilmiş bir link (ör. iki kişiye aynı TikTok'un verilmesi) bu kıyası sessizce bozduğu için akışta link kontrolü de var.

İki script var (yollar bu skill klasörüne göre):
- `scripts/engagement.py`: linkleri okur, kontrol eder, Apify'dan veriyi çeker, raporu yazar.
- `scripts/link_duzelt.py`: Excel'de tek bir hücrenin linkini ve görünen değerini düzeltir.

Kurulum: `pip install -r requirements.txt`

## Akış

### 1. Shortlist'i incele

openpyxl ile ilk ~10 satıra bak; hücre değerlerini ve **hyperlink hedeflerini** birlikte yazdır. Linkler genelde takipçi sayısının yazdığı hücreye bağlıdır (hücrede "837K" görünür, link profile gider). Şunları belirle:
- Sayfa adı, kolon başlıklarının olduğu satır
- Influencer adı kolonu, Instagram link kolonu, TikTok link kolonu

Varsayılan düzen Samsung shortlist'inki: 2. satır başlık, A = Influencer, B = IG, D = TT. Farklıysa `--sayfa`, `--baslik-satiri`, `--isim-kolon`, `--ig-kolon`, `--tt-kolon` ile ver. Platformlardan biri yoksa o kolonu boş geç (`--tt-kolon ""`).

Excel'de metodoloji notu varsa (ör. "Engagement Rate = Toplam Etkileşim / Toplam Takipçi") oku. Aşağıdaki hesaplama kurallarıyla çelişiyorsa kullanıcıya sor.

### 2. Linkleri çıkar ve kontrol et

```bash
python scripts/engagement.py "<shortlist.xlsx>" --sadece-linkler [kolon seçenekleri]
```

Bu adım Apify'a gitmez ve ücret çıkarmaz. Kullanıcıya kısaca özetle: kaç influencer bulundu, kimin TikTok'u veya Instagram'ı yok, hangi uyarılar çıktı.

Uyarıları şöyle değerlendir:
- **"aynı hesap başka bir influencer'da da var"**: neredeyse her zaman kopyala-yapıştır hatasıdır. Hangi satırın yanlış olduğunu isimden anla.
- **"hesap adı influencer adıyla uyuşmuyor"**: çoğu zaman marka veya takma ad hesabıdır (ör. "Hype" → @aynisinemalar, "Emre Gökgöz" → @fityemek). Bunları kendi muhakemenle ele. Gerçekten şüpheli olanları (ör. çok genel bir kullanıcı adı veya Instagram ile TikTok'ta bambaşka kullanıcı adları) web'de doğrula.

Doğrulama için web araması kullan: `"tiktok.com/@kullaniciadi"`, `"@kullaniciadi" tiktok`, `<isim> TikTok hesabı`. TikTok discover sayfaları, profil bio'ları ve Instagram ↔ TikTok çapraz linkleri en güvenilir kaynaklar. Doğrulayamadığın bir hesabı doğru ya da yanlış diye sunma; "doğrulanamadı" de. Yanlış link raporu sessizce bozar, uydurma bir "düzeltme" ise daha kötüsüdür.

### 3. Yanlış linkleri düzelt

Kesin olarak yanlış olan ve doğrusu bulunan linkleri düzelt. Kullanıcının orijinal dosyasının **bir kopyası** üzerinde çalış:

```bash
python scripts/link_duzelt.py "<kopya.xlsx>" D54 https://www.tiktok.com/@obenalkan --deger 197K
```

`--deger`, hücrede görünen metni (takipçi sayısını) da günceller. Eski değer başka hesabınki ise bunu kullan. Excel'i openpyxl ile açıp kaydetme. Kurumsal dosyalarda Microsoft hassasiyet etiketi (`docMetadata/LabelInfo.xml`) ve eklenti panelleri olur, openpyxl bunları siler. `link_duzelt.py` sadece ilgili XML'i değiştirir. Düzelttikten sonra 2. adımı tekrar çalıştırıp uyarının kalktığını gör.

### 4. Raporu üret

```bash
APIFY_TOKEN=... python scripts/engagement.py "<shortlist.xlsx>" [kolon seçenekleri]
```

- Token ortamda yoksa kullanıcıdan token'ı **sohbete yapıştırmasını isteme**. Ortam değişkeni ya da ortamın secret ayarı olarak tanımlamasını söyle (token: https://console.apify.com/settings/integrations).
- api.apify.com'a erişim yoksa (ör. bulut oturumunda ağ politikası engelliyorsa) bunu kullanıcıya söyle ve hostun izin listesine eklenmesi gerektiğini belirt. **Sahte veya tahmini veriyle rapor üretme.** Rakamlar müşteriye gidebilir.
- Her platform için tek bir actor çalıştırması yapılır, tüm hesaplar birlikte gönderilir. Ücret kullanıcının Apify hesabından çıkar.

Çıktı, girdinin yanına `<ad> - ER Raporu.xlsx` olarak yazılır, orijinal dosya değişmez:

| Sayfa | İçerik |
|---|---|
| Özet | One-pager: en üstte formüller ve açıklamalar; tabloda her influencer için tıklanabilir hesaplar, güncel takipçi, ER (renk skalalı) ve kontrol notu. Yatay A4'e sığar. |
| Link Kontrol | Her hesabın Apify'daki profil adı, Excel'deki ve güncel takipçisi, uyarılar (uyarılılar üstte). |
| Post Detay | ER'e giren her içerik: tarih, link, beğeni, yorum, paylaşım, etkileşim, içerik ER'i. |

### 5. Sonucu özetle

Kullanıcıya şunları söyle:
- En yüksek ve en düşük ER'ler
- Apify sonrası çıkan yeni uyarılar ("takipçi Excel'dekinden çok farklı" ve "Apify'da bulunamadı" sadece bu aşamada çıkar)
- Verisi gelmeyen hesaplar

Bir hesabın ER'si olağandışı görünüyorsa (çok yüksek ya da çok düşük) Post Detay'dan o hesabın içeriklerine bak; sebebi genelde viral bir içerik ya da gizlenmiş beğenilerdir.

## Hesaplama kuralları

```
Instagram ER = (Beğeni + Yorum) / Takipçi
TikTok ER    = (Beğeni + Yorum + Paylaşım) / Takipçi
ER = son 5 içeriğin toplam etkileşimi / (Takipçi × 5)    → içerik başına ortalama
```

- **Neden son 5 içeriğin ortalaması:** Tek bir içerik viral olabilir ya da sönük kalabilir. Ortalama bunu dengeler ve içerik başına ER, sektördeki benchmark'larla kıyaslanabilir. Toplamı takipçiye bölmek ER'yi 5 kat şişirir.
- **Neden Instagram'da paylaşım yok:** Instagram paylaşım sayısını herkese açık vermiyor.
- **Son 5 içerik nasıl seçilir:** Yayın tarihine göre en yeni 5 içerik alınır. Profilin en üstüne sabitlenmiş (pinned) eski içerikler hesabı bozmasın diye fazladan çekilip tarihe göre sıralanır. 5'ten az içerik varsa mevcut içeriklerle hesaplanır.
- **Takipçi:** Apify'dan rapor anında çekilir. Excel'deki rakamlar eski olabileceği için ER'de kullanılmaz, sadece link kontrolünde karşılaştırılır.
- **Gizli beğeni:** Beğenisi gizlenmiş Instagram içeriklerinde beğeni 0 sayılır, ER düşük çıkar. Raporda notu var.

Kullanıcı formülü değiştirmek isterse `scripts/engagement.py` başındaki `IG_ALANLAR` / `TT_ALANLAR` sözlüklerini düzenle; rapordaki formül metinleri bunlardan üretilir. İçerik sayısı için `POST_SAYISI` değişkenini değiştir. Diğer Apify alanları: TikTok'ta kaydetme `collectCount`, izlenme `playCount`; Instagram'da video izlenme `videoViewCount`.

## Bilinen sınırlar

- Sadece tam profil linkleri okunur (`instagram.com/<kullanıcı>`, `tiktok.com/@<kullanıcı>`). Kısa linkler (`vm.tiktok.com/...`) ve linksiz `@kullanıcı` metinleri okunmaz. Bunları web'de bulup `link_duzelt.py` ile düzelt ya da kullanıcıya sor.
- Gizli (private) veya silinmiş hesaplar "Apify'da bulunamadı" uyarısı alır.
- Kullanılan actor'ler: `apify/instagram-profile-scraper` (takipçi `followersCount`, içerikler `latestPosts`) ve `clockworks/tiktok-scraper` (takipçi `authorMeta.fans`). Apify actor çıktısını değiştirirse ER'ler 0 ya da boş gelir. Böyle bir durumda Post Detay'ı kontrol et ve alan adlarını güncelle.
