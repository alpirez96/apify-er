"""Shortlist Excel'indeki influencer'lar için Instagram ve TikTok engagement rate hesaplar.

Kullanım:
    python engagement.py "Samsung - Shortlist.xlsx" --sadece-linkler   # sadece bulunan hesapları listeler
    APIFY_TOKEN=... python engagement.py "Samsung - Shortlist.xlsx"    # Apify'ı çalıştırıp ER yazar
"""

import argparse
import os
import re
import sys
from collections import Counter
from pathlib import Path

from apify_client import ApifyClient
from openpyxl import load_workbook

POST_SAYISI = 5

# Shortlist düzeni (ilk sayfa): 2. satır başlık, A=Influencer, B=IG (link), C=ER%, D=TT (link), E=ER%
BASLIK_SATIRI = 2
ISIM_KOLON = "A"
IG_KOLON, IG_ER_KOLON = "B", "C"
TT_KOLON, TT_ER_KOLON = "D", "E"

IG_ACTOR = "apify/instagram-profile-scraper"
TT_ACTOR = "clockworks/tiktok-scraper"

# Toplam Etkileşim = Like + Send(Share) + Etiketleme
# Her bileşenin Apify çıktısındaki alan adı. None = platform bu veriyi public vermiyor, 0 sayılır.
# "Etiketleme" için public bir sayaç yok; en yakın karşılık olan yorum sayısı kullanılıyor.
IG_ALANLAR = {"like": "likesCount", "share": None, "etiketleme": "commentsCount"}
TT_ALANLAR = {"like": "diggCount", "share": "shareCount", "etiketleme": "commentCount"}

IG_REGEX = re.compile(r"instagram\.com/([A-Za-z0-9._]+)", re.I)
TT_REGEX = re.compile(r"tiktok\.com/@([A-Za-z0-9._]+)", re.I)


def kullanici_adi(hucre, regex):
    """Hücreye bağlı hyperlink'ten (yoksa hücre metninden) kullanıcı adını çıkarır."""
    metin = f"{hucre.hyperlink.target if hucre.hyperlink else ''} {hucre.value or ''}"
    eslesme = regex.search(metin)
    return eslesme.group(1).lower() if eslesme else None


def hesaplari_bul(ws):
    satirlar = []  # (satır no, isim, ig, tt)
    for no in range(BASLIK_SATIRI + 1, ws.max_row + 1):
        ig = kullanici_adi(ws[f"{IG_KOLON}{no}"], IG_REGEX)
        tt = kullanici_adi(ws[f"{TT_KOLON}{no}"], TT_REGEX)
        if ig or tt:
            satirlar.append((no, str(ws[f"{ISIM_KOLON}{no}"].value or "").strip(), ig, tt))
    return satirlar


def actor_calistir(client, actor, run_input):
    print(f"-> {actor} çalışıyor...")
    run = client.actor(actor).call(run_input=run_input)
    if run is None or run.status != "SUCCEEDED":
        sys.exit(f"{actor} başarısız oldu: {run and run.status}")
    return list(client.dataset(run.default_dataset_id).iterate_items())


def etkilesim(post, alanlar):
    # Gizli like'lar Apify'da -1 / None gelebilir, 0 sayılır.
    return {k: max(post.get(alan) or 0, 0) if alan else 0 for k, alan in alanlar.items()}


def instagram_verisi(client, kullanicilar):
    items = actor_calistir(client, IG_ACTOR, {"usernames": kullanicilar})
    sonuc = {}
    for p in items:
        if not p.get("username"):
            continue
        postlar = sorted(p.get("latestPosts") or [], key=lambda x: x.get("timestamp") or "", reverse=True)
        sonuc[p["username"].lower()] = {
            "takipci": p.get("followersCount") or 0,
            "postlar": [
                {"tarih": x.get("timestamp"), "url": x.get("url"), **etkilesim(x, IG_ALANLAR)}
                for x in postlar[:POST_SAYISI]
            ],
        }
    return sonuc


def tiktok_verisi(client, kullanicilar):
    items = actor_calistir(client, TT_ACTOR, {
        "profiles": kullanicilar,
        "profileScrapeSections": ["videos"],
        "profileSorting": "latest",
        # Sabitlenmiş (pinned) eski videolar da gelebildiği için fazladan çekip tarihe göre sıralıyoruz.
        "resultsPerPage": POST_SAYISI + 5,
        "shouldDownloadVideos": False,
        "shouldDownloadCovers": False,
        "shouldDownloadSubtitles": False,
        "shouldDownloadSlideshowImages": False,
    })
    sonuc = {}
    for v in items:
        yazar = v.get("authorMeta") or {}
        if not yazar.get("name"):
            continue
        hesap = sonuc.setdefault(yazar["name"].lower(), {"takipci": yazar.get("fans") or 0, "postlar": []})
        hesap["postlar"].append({"tarih": v.get("createTimeISO"), "url": v.get("webVideoUrl"), **etkilesim(v, TT_ALANLAR)})
    for hesap in sonuc.values():
        hesap["postlar"] = sorted(hesap["postlar"], key=lambda x: x["tarih"] or "", reverse=True)[:POST_SAYISI]
    return sonuc


def engagement_rate(hesap):
    """ER = Toplam Etkileşim / (Takipçi x Post Sayısı), yani post başına ortalama etkileşim / takipçi."""
    if not hesap or not hesap["postlar"] or not hesap["takipci"]:
        return None
    toplam = sum(p["like"] + p["share"] + p["etiketleme"] for p in hesap["postlar"])
    return toplam / (hesap["takipci"] * len(hesap["postlar"]))


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("excel", type=Path)
    parser.add_argument("--sadece-linkler", action="store_true", help="Apify çalıştırmadan bulunan hesapları listele")
    args = parser.parse_args()

    wb = load_workbook(args.excel)
    ws = wb.worksheets[0]
    satirlar = hesaplari_bul(ws)
    print(f"{len(satirlar)} influencer bulundu:")
    for no, isim, ig, tt in satirlar:
        print(f"  #{no:<3} {isim:<28} instagram={ig or '-':<26} tiktok={tt or '-'}")
    for platform, i in (("Instagram", 2), ("TikTok", 3)):
        for kullanici, adet in Counter(s[i] for s in satirlar if s[i]).items():
            if adet > 1:
                isimler = ", ".join(s[1] for s in satirlar if s[i] == kullanici)
                print(f"  UYARI: {platform} hesabı '{kullanici}' birden fazla satırda: {isimler}")
    if args.sadece_linkler:
        return

    token = os.environ.get("APIFY_TOKEN")
    if not token:
        sys.exit("APIFY_TOKEN ortam değişkeni tanımlı değil.")
    client = ApifyClient(token)

    ig_kullanicilar = sorted({s[2] for s in satirlar if s[2]})
    tt_kullanicilar = sorted({s[3] for s in satirlar if s[3]})
    ig = instagram_verisi(client, ig_kullanicilar) if ig_kullanicilar else {}
    tt = tiktok_verisi(client, tt_kullanicilar) if tt_kullanicilar else {}

    # ER'leri shortlist'teki ER% kolonlarına yaz. Veri gelmeyen hesapta mevcut değer olduğu gibi kalır.
    for no, _, ig_ad, tt_ad in satirlar:
        for veri, ad, kolon in ((ig, ig_ad, IG_ER_KOLON), (tt, tt_ad, TT_ER_KOLON)):
            er = engagement_rate(veri.get(ad))
            if er is not None:
                ws[f"{kolon}{no}"] = er
                ws[f"{kolon}{no}"].number_format = "0.00%"

    # Kontrol için post bazında detay sayfası.
    detay = wb.create_sheet("Post Detay")
    detay.append(["Influencer", "Platform", "Kullanıcı", "Takipçi", "Tarih", "URL",
                  "Like", "Share", "Etiketleme (Yorum)", "Etkileşim"])
    for no, isim, ig_ad, tt_ad in satirlar:
        for platform, veri, ad in (("Instagram", ig, ig_ad), ("TikTok", tt, tt_ad)):
            hesap = veri.get(ad)
            for p in (hesap or {}).get("postlar", []):
                detay.append([isim, platform, ad, hesap["takipci"], p["tarih"], p["url"],
                              p["like"], p["share"], p["etiketleme"], p["like"] + p["share"] + p["etiketleme"]])

    bulunamayan = [f"IG:{u}" for u in ig_kullanicilar if u not in ig] + [f"TT:{u}" for u in tt_kullanicilar if u not in tt]
    if bulunamayan:
        print("Veri gelmeyen hesaplar:", ", ".join(bulunamayan))

    cikti = args.excel.with_name(f"{args.excel.stem} - ER.xlsx")
    wb.save(cikti)
    print(f"Kaydedildi: {cikti}")


if __name__ == "__main__":
    main()
