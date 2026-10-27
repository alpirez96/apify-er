"""Shortlist Excel'inde bir hücrenin profil linkini (ve istenirse görünen değerini) yerinde düzeltir.

Dosya openpyxl ile açılıp kaydedilmez: bu, Microsoft hassasiyet etiketi (docMetadata/LabelInfo.xml),
eklenti paneli gibi parçaları siler. Sadece ilgili hücre ve link XML'i değiştirilir, geri kalan her şey aynen kopyalanır.

Kullanım:
    python link_duzelt.py "Samsung - Shortlist.xlsx" D54 https://www.tiktok.com/@obenalkan --deger 197K
"""

import argparse
import html
import os
import re
import sys
import zipfile


def sayfa_yolu(z, sayfa_adi):
    """Sayfa adına (verilmezse ilk sayfaya) karşılık gelen worksheet XML yolunu bulur."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for etiket in re.findall(r"<sheet\b[^>]*>", wb):
        ad = html.unescape(re.search(r'\bname="([^"]*)"', etiket).group(1))
        if sayfa_adi is None or ad == sayfa_adi:
            rid = re.search(r'\br:id="([^"]*)"', etiket).group(1)
            rel = re.search(rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*>', rels).group(0)
            hedef = re.search(r'\bTarget="([^"]*)"', rel).group(1)
            return hedef.lstrip("/") if hedef.startswith("/") else f"xl/{hedef}"
    sys.exit(f"Sayfa bulunamadı: {sayfa_adi}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("excel")
    parser.add_argument("hucre", help="Örn. D54")
    parser.add_argument("url", help="Yeni profil linki")
    parser.add_argument("--deger", help="Hücrede görünecek yeni metin (örn. takipçi sayısı '197K')")
    parser.add_argument("--sayfa", help="Sayfa adı (varsayılan: ilk sayfa)")
    args = parser.parse_args()
    hucre = args.hucre.upper()

    with zipfile.ZipFile(args.excel) as z:
        sayfa = sayfa_yolu(z, args.sayfa)
        klasor, dosya = sayfa.rsplit("/", 1)
        sayfa_rels = f"{klasor}/_rels/{dosya}.rels"
        xml = z.read(sayfa).decode("utf-8")
        rels = z.read(sayfa_rels).decode("utf-8")

        # Hücrenin hyperlink'i -> ilişki (rels) kaydındaki Target
        link = re.search(rf'<hyperlink\b[^>]*\bref="{hucre}"[^>]*>', xml)
        if not link:
            sys.exit(f"{hucre} hücresinde link yok.")
        rid = re.search(r'\br:id="([^"]*)"', link.group(0)).group(1)
        rel = re.search(rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*>', rels).group(0)
        eski_url = html.unescape(re.search(r'\bTarget="([^"]*)"', rel).group(1))
        rels = rels.replace(rel, re.sub(r'\bTarget="[^"]*"', f'Target="{html.escape(args.url)}"', rel))

        if args.deger is not None:
            # Paylaşılan metin başka hücrelerde de kullanılıyor olabilir; bu hücreye satır içi metin yazılır.
            c = re.search(rf'<c\b(?=[^>]*\br="{hucre}")([^>]*?)(/>|>.*?</c>)', xml, re.S)
            if not c:
                sys.exit(f"{hucre} hücresi bulunamadı.")
            nitelikler = re.sub(r'\s+t="[^"]*"', "", c.group(1))
            xml = xml.replace(c.group(0), f'<c{nitelikler} t="inlineStr"><is><t>{html.escape(args.deger)}</t></is></c>')

        degisen = {sayfa: xml, sayfa_rels: rels}
        gecici = f"{args.excel}.tmp"
        with zipfile.ZipFile(gecici, "w") as zout:
            for info in z.infolist():
                zout.writestr(info, degisen[info.filename].encode("utf-8") if info.filename in degisen else z.read(info))
    os.replace(gecici, args.excel)
    print(f"{hucre}: {eski_url} -> {args.url}" + (f" (değer: {args.deger})" if args.deger is not None else ""))


if __name__ == "__main__":
    main()
