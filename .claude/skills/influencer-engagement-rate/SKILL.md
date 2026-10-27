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

### 0. Apify key'ini kullanıcıyla birlikte kontrol et

Skill Apify key'i olmadan çalışmaz; script de key yoksa hiçbir modda çalışmaz. Shortlist'e bakmadan önce, **her çalıştırmada** kullanılacak key'i kullanıcıyla birlikte netleştir. Ekipte birden fazla Apify hesabı olabilir ve ücret seçilen hesaptan çıkar; yanlış hesapla çalışmak sonradan fark edilmesi zor bir hatadır.

1. Ortamda `APIFY_TOKEN` var mı bak, değerini yazdırma: `test -n "$APIFY_TOKEN" && echo var || echo yok`
2. **Yoksa** kullanıcıdan key'i iste ve key gelene kadar başka hiçbir adıma geçme. Key'i Apify Console → Settings → Integrations'tan (https://console.apify.com/settings/integrations) alabilir. İki yolu var:
   - Ortam değişkeni olarak tanımlaması (Mac'te `~/.zshrc`'ye `export APIFY_TOKEN=...`, bulut oturumunda ortamın secret ayarı). Kalıcıdır, önerilen yol budur.
   - Sohbete yapıştırması. Bu durumda key'i sadece komutların başında ortam değişkeni olarak ver (`APIFY_TOKEN='...' python ...`). Hiçbir dosyaya yazma, cevaplarında tamamını tekrar etme. Sohbet geçmişinde kalacağını bir kez hatırlat.
3. Key'i doğrula ve hangi hesaba ait olduğunu göster:
   ```bash
   python scripts/engagement.py --key-kontrol
   ```
   Çıktı: `Apify key: apify_api_Ab…1234 | hesap: <kullanıcı adı> <e-posta>`. Bu satırı kullanıcıya göster ve **"Bu key / hesapla devam edeyim mi?"** diye sor. Onay gelmeden devam etme. Key ortamda tanımlı olsa bile bu onayı her seferinde al. Kullanıcı başka bir key isterse 2. adıma dön.

Script her çalıştırmada aynı satırı en başta tekrar yazdırır. Sonraki adımlarda hesabın değişmediğini buradan görebilirsin.

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

Bu adım sadece key'i doğrular, actor çalıştırmaz, ücret çıkmaz. Kullanıcıya kısaca özetle: kaç influencer bulundu, kimin TikTok'u veya Instagram'ı yok, hangi uyarılar çıktı.

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

- Bu adımda 0. adımda onaylanan key kullanılır. Çıktının ilk satırındaki hesabın onaylananla aynı olduğuna bak.
- api.apify.com'a erişim yoksa (ör. bulut oturumunda ağ politikası engelliyorsa) bunu kullanıcıya söyle ve hostun izin listesine eklenmesi gerektiğini belirt. **Sahte veya tahmini veriyle rapor üretme.** Rakamlar müşteriye gidebilir.
- Her platform için tek bir actor çalıştırması yapılır, tüm hesaplar birlikte gönderilir. Ücret kullanıcının Apify hesabından çıkar.

Çıktı, girdinin yanına `<ad> - ER Raporu.xlsx` olarak yazılır, orijinal dosya değişmez:

| Sayfa | İçerik |
|---|---|
| Özet | One-pager: en üstte formüller ve açıklamalar; tabloda her influencer için tıklanabilir hesaplar, güncel takipçi, ER (renk skalalı) ve kontrol notu. Yatay A4'e sığar. |
| Link Kontrol | Her hesabın Apify'daki profil adı, Excel'deki ve güncel takipçisi, uyarılar (uyarılılar üstte). |
| Post Detay | ER'e giren her içerik: tarih, link, beğeni, yorum, paylaşım, etkileşim, içerik ER'i. |

### 5. Dosyaları teslim et ve sonucu özetle

Kullanıcıya iki dosyayı ver: **ER raporu** ve linkler düzeltildiyse **düzeltilmiş shortlist**. Kullanıcının bilgisayarında çalışıyorsan dosyalar orijinalin yanındadır, tam yollarını söyle. Bulut oturumu veya claude.ai gibi kullanıcının dosya sistemine erişmediği bir ortamdaysan dosyayı mevcut dosya gönderme aracıyla (ya da ortamın çıktı klasörüne koyarak) indirilebilir şekilde ver; sadece yolunu söylemek yetmez.

Ardından şunları söyle:
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
