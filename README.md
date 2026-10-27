# Influencer Engagement Rate

Shortlist Excel'indeki her influencer için Instagram ve TikTok'taki **son 5 içeriğe** bakarak engagement rate (ER) hesaplar ve tek sayfalık, tasarımlı bir Excel raporu üretir. Veriler (takipçi sayıları dahil) Apify'dan güncel olarak çekilir.

## Claude skill'i olarak kullanım

Bu repo bir Claude skill'i içerir: [`.claude/skills/influencer-engagement-rate`](.claude/skills/influencer-engagement-rate/SKILL.md). Skill; linkleri kontrol etme, yanlış linkleri web'de doğrulayıp düzeltme, raporu üretme ve sonucu özetleme adımlarını Claude'a anlatır.

- **Claude Code:** Bu repoyu açman yeterli. Skill otomatik yüklenir. Shortlist'i verip "bu listenin ER'lerini çıkar" demen yeterli.
- **claude.ai:** `influencer-engagement-rate.skill` dosyasını sohbette açıp **Save skill** ile profiline ekle. Kod çalıştırma ortamının api.apify.com'a erişimi olmalı.

Skill'i güncellemek için `.claude/skills/influencer-engagement-rate` klasörünü düzenle. Hesaplama kuralları `SKILL.md` içinde, kod `scripts/` altında.

## Kurulum (elle çalıştırmak için)

```bash
pip install -r .claude/skills/influencer-engagement-rate/requirements.txt
```

Apify API token'ı: https://console.apify.com/settings/integrations

## Kullanım

**1. Önce linkleri kontrol et** (Apify çalışmaz, ücret çıkmaz):

```bash
python .claude/skills/influencer-engagement-rate/scripts/engagement.py "Samsung - Shortlist.xlsx" --sadece-linkler
```

Her influencer için bulunan Instagram / TikTok hesabını ve olası link hatalarını listeler (aşağıda "Link kontrolü").

**2. Raporu oluştur:**

```bash
export APIFY_TOKEN=apify_api_xxx
python .claude/skills/influencer-engagement-rate/scripts/engagement.py "Samsung - Shortlist.xlsx"
```

Çıktı aynı klasöre `Samsung - Shortlist - ER Raporu.xlsx` adıyla kaydedilir. Orijinal dosya değişmez.

| Sayfa | İçerik |
|---|---|
| **Özet** | One-pager. Formüller ve açıklamalar en üstte; altında her influencer için Instagram ve TikTok hesabı (tıklanabilir), güncel takipçi, ER ve kontrol notu. ER kolonları renk skalalı. Yatay A4'e sığacak şekilde ayarlı. |
| **Link Kontrol** | Her hesap için Apify'daki profil adı, Excel'deki takipçi ve güncel takipçi, uyarılar. Uyarılı satırlar en üstte. |
| **Post Detay** | Hesaplamaya giren her içerik: tarih, link, beğeni, yorum, paylaşım, etkileşim, içerik ER'i. |

## Excel düzeni

Script ilk sayfayı okur. 2. satır başlık, sonraki satırlar influencer:

| A | B | D |
|---|---|---|
| Influencer | IG takipçi (hücreye profil linki bağlı) | TT takipçi (hücreye profil linki bağlı) |

Kullanıcı adı, hücreye bağlı hyperlink'ten okunur (`instagram.com/<kullanıcı>`, `tiktok.com/@<kullanıcı>`). Link olmayan (`-`) hücreler atlanır. Farklı düzendeki bir shortlist için: `--sayfa`, `--baslik-satiri`, `--isim-kolon`, `--ig-kolon`, `--tt-kolon` (platform yoksa `--tt-kolon ""`).

## Hesaplama

```
Instagram ER = (Beğeni + Yorum) / Takipçi
TikTok ER    = (Beğeni + Yorum + Paylaşım) / Takipçi
```

ER, hesabın **son 5 içeriğinin ortalamasıdır**:

```
ER = Son 5 içeriğin toplam etkileşimi / (Takipçi × 5)
```

Örnek: 100.000 takipçi, son 5 içerikte toplam 25.000 etkileşim → `25.000 / (100.000 × 5) = %5,00`

- **Takipçi:** Apify'dan rapor anında çekilir; Excel'deki rakamlar ER için kullanılmaz, sadece link kontrolünde karşılaştırılır.
- **Son 5 içerik:** Instagram'dan profilin son ~12 içeriği, TikTok'tan son 10 videosu çekilir. Yayın tarihine göre sıralanır ve en yeni 5'i alınır. Böylece en üste sabitlenmiş (pinned) eski içerikler hesaba girmez. 5'ten az içerik varsa mevcut içeriklerle hesaplanır.
- **Gizli beğeni:** Beğenileri gizlenmiş Instagram içeriklerinde beğeni 0 sayılır.

Kullanılan Apify actor'leri (her platform için tek çalıştırma, tüm hesaplar birlikte):

| Platform | Actor | Takipçi | Beğeni | Yorum | Paylaşım |
|---|---|---|---|---|---|
| Instagram | `apify/instagram-profile-scraper` | `followersCount` | `likesCount` | `commentsCount` | – |
| TikTok | `clockworks/tiktok-scraper` | `authorMeta.fans` | `diggCount` | `commentCount` | `shareCount` |

## Link kontrolü

Yanlış girilmiş linkleri yakalamak için her hesap şu kurallarla kontrol edilir:

- **Aynı hesap birden fazla influencer'da** geçiyor (ör. iki kişiye aynı TikTok linki verilmiş).
- **Hesap adı influencer adıyla uyuşmuyor:** influencer adındaki kelimelerin hiçbiri ne kullanıcı adında ne de Apify'dan gelen profil adında geçiyor. Marka/takma adlı hesaplarda (ör. "Hype" → @aynisinemalar) normal olarak uyarı verebilir; göz ile kontrol yeterli.
- **Takipçi Excel'dekinden çok farklı:** güncel takipçi, Excel'deki rakamın yarısından az ya da iki katından fazla.
- **Apify'da bulunamadı:** hesap kapalı, silinmiş veya kullanıcı adı yanlış.

`--sadece-linkler` modunda Apify çalışmadığı için sadece ilk iki kontrol (kullanıcı adına göre) yapılır.

Yanlış bir linki düzeltmek için (hücrede görünen takipçi metniyle birlikte):

```bash
python .claude/skills/influencer-engagement-rate/scripts/link_duzelt.py "Samsung - Shortlist.xlsx" D54 https://www.tiktok.com/@obenalkan --deger 197K
```

Dosya openpyxl ile yeniden kaydedilmez, sadece o hücrenin XML'i değişir. Böylece Microsoft hassasiyet etiketi gibi parçalar korunur.

## Ayar değiştirmek

`.claude/skills/influencer-engagement-rate/scripts/engagement.py` dosyasının başında:
- `POST_SAYISI = 5`: kaç içerik alınacağı
- `IG_ALANLAR` / `TT_ALANLAR`: etkileşime hangi alanların gireceği. Rapordaki formül metinleri de buradan üretilir.
