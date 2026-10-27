"""Shortlist Excel'indeki influencer'lar için Instagram ve TikTok engagement rate hesaplar.

Kullanım:
    python engagement.py "Samsung - Shortlist.xlsx" --sadece-linkler   # sadece bulunan hesapları listeler
    APIFY_TOKEN=... python engagement.py "Samsung - Shortlist.xlsx"    # Apify'ı çalıştırıp ER yazar
"""

import argparse
import os
import re
import sys
from pathlib import Path

from apify_client import ApifyClient
from openpyxl import load_workbook

POST_SAYISI = 5

IG_ACTOR = "apify/instagram-profile-scraper"
TT_ACTOR = "clockworks/tiktok-scraper"

# Toplam Etkileşim = Like + Send(Share) + Etiketleme
# Her bileşenin Apify çıktısındaki alan adı. None = platform bu veriyi public vermiyor, 0 sayılır.
# "Etiketleme" için public bir sayaç yok; en yakın karşılık olan yorum sayısı kullanılıyor.
IG_ALANLAR = {"like": "likesCount", "share": None, "etiketleme": "commentsCount"}
TT_ALANLAR = {"like": "diggCount", "share": "shareCount", "etiketleme": "commentCount"}

IG_REGEX = re.compile(r"instagram\.com/([A-Za-z0-9._]+)", re.I)
IG_HARIC = {"p", "reel", "reels", "tv", "stories", "explore", "accounts"}
TT_REGEX = re.compile(r"tiktok\.com/@([A-Za-z0-9._]+)", re.I)


def hesaplari_bul(path):
    """Her satırdaki hücre metinlerinden ve hyperlink'lerden IG / TikTok kullanıcı adlarını çıkarır."""
    wb = load_workbook(path)
    satirlar = []  # (sayfa adı, satır no, ig, tt)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            metin = " ".join(
                f"{c.value or ''} {c.hyperlink.target if c.hyperlink else ''}" for c in row
            )
            ig = next((u for u in IG_REGEX.findall(metin) if u.lower() not in IG_HARIC), None)
            tt = next(iter(TT_REGEX.findall(metin)), None)
            if ig or tt:
                satirlar.append((ws.title, row[0].row, ig and ig.lower(), tt and tt.lower()))
    return wb, satirlar


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


def ozet(hesap):
    """(takipçi, post sayısı, toplam etkileşim, ER) döner. ER = Toplam Etkileşim / (Takipçi x Post Sayısı)."""
    if not hesap or not hesap["postlar"] or not hesap["takipci"]:
        return None
    n = len(hesap["postlar"])
    toplam = sum(p["like"] + p["share"] + p["etiketleme"] for p in hesap["postlar"])
    return hesap["takipci"], n, toplam, toplam / (hesap["takipci"] * n)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("excel", type=Path)
    parser.add_argument("--sadece-linkler", action="store_true", help="Apify çalıştırmadan bulunan hesapları listele")
    args = parser.parse_args()

    wb, satirlar = hesaplari_bul(args.excel)
    print(f"{len(satirlar)} satırda hesap bulundu:")
    for sayfa, no, ig, tt in satirlar:
        print(f"  [{sayfa} #{no}] instagram={ig or '-'}  tiktok={tt or '-'}")
    if args.sadece_linkler:
        return

    token = os.environ.get("APIFY_TOKEN")
    if not token:
        sys.exit("APIFY_TOKEN ortam değişkeni tanımlı değil.")
    client = ApifyClient(token)

    ig_kullanicilar = sorted({ig for _, _, ig, _ in satirlar if ig})
    tt_kullanicilar = sorted({tt for _, _, _, tt in satirlar if tt})
    ig = instagram_verisi(client, ig_kullanicilar) if ig_kullanicilar else {}
    tt = tiktok_verisi(client, tt_kullanicilar) if tt_kullanicilar else {}

    # Sonuçları her sayfanın sağına yeni kolonlar olarak yaz.
    basliklar = ["IG Kullanıcı", "IG Takipçi", "IG Post", "IG Toplam Etkileşim", "IG ER",
                 "TT Kullanıcı", "TT Takipçi", "TT Video", "TT Toplam Etkileşim", "TT ER", "Genel ER"]
    ilk_kolon = {}
    for sayfa, no, ig_ad, tt_ad in satirlar:
        ws = wb[sayfa]
        if sayfa not in ilk_kolon:
            ilk_kolon[sayfa] = ws.max_column + 1
            baslik_satiri = max(no - 1, 1)
            for i, b in enumerate(basliklar):
                ws.cell(baslik_satiri, ilk_kolon[sayfa] + i, b)
        ig_ozet, tt_ozet = ozet(ig.get(ig_ad)), ozet(tt.get(tt_ad))
        degerler = [ig_ad, *(ig_ozet or [None] * 4), tt_ad, *(tt_ozet or [None] * 4)]
        # Genel ER = (IG + TT toplam etkileşim) / (IG takipçi x post + TT takipçi x video)
        ozetler = [o for o in (ig_ozet, tt_ozet) if o]
        degerler.append(sum(o[2] for o in ozetler) / sum(o[0] * o[1] for o in ozetler) if ozetler else None)
        for i, d in enumerate(degerler):
            c = ws.cell(no, ilk_kolon[sayfa] + i, d)
            if basliklar[i].endswith("ER"):
                c.number_format = "0.00%"

    # Kontrol için post bazında detay sayfası.
    detay = wb.create_sheet("Post Detay")
    detay.append(["Platform", "Kullanıcı", "Takipçi", "Tarih", "URL", "Like", "Share", "Etiketleme (Yorum)", "Etkileşim"])
    for platform, veri in (("Instagram", ig), ("TikTok", tt)):
        for kullanici, hesap in veri.items():
            for p in hesap["postlar"]:
                detay.append([platform, kullanici, hesap["takipci"], p["tarih"], p["url"],
                              p["like"], p["share"], p["etiketleme"], p["like"] + p["share"] + p["etiketleme"]])

    bulunamayan = [f"IG:{u}" for u in ig_kullanicilar if u not in ig] + [f"TT:{u}" for u in tt_kullanicilar if u not in tt]
    if bulunamayan:
        print("Veri gelmeyen hesaplar:", ", ".join(bulunamayan))

    cikti = args.excel.with_name(f"{args.excel.stem} - ER.xlsx")
    wb.save(cikti)
    print(f"Kaydedildi: {cikti}")


if __name__ == "__main__":
    main()
