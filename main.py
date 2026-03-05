import json
import pandas as pd
import glob
import zipfile
import os

def arxivdan_excelga_otkazish(papka_yoli='.', excel_fayl_nomi="Fakturalar_hisoboti.xlsx"):
    barcha_qatorlar = []

    # Papkadagi barcha .zip fayllarni topish
    zip_fayllar = glob.glob(os.path.join(papka_yoli, "*.zip"))

    if not zip_fayllar:
        print("Ushbu papkada ZIP fayllar topilmadi.")
        return

    json_topildi = 0

    # Har bir ZIP faylni ochib ko'rish
    for zip_nomi in zip_fayllar:
        with zipfile.ZipFile(zip_nomi, 'r') as z:
            fayllar_royxati = z.namelist()

            for fayl in fayllar_royxati:
                # Agar arxiv ichidagi fayl .json bo'lsa
                if fayl.endswith('.json'):
                    json_topildi += 1

                    # JSON faylni arxivdan to'g'ridan-to'g'ri xotiraga o'qish
                    with z.open(fayl) as f:
                        data = json.load(f)

                    # Ma'lumotlarni ajratib olish
                    faktura_raqami = data.get("facturadoc", {}).get("facturano", "")
                    faktura_sanasi = data.get("facturadoc", {}).get("facturadate", "")

                    sotuvchi_nomi = data.get("seller", {}).get("name", "")
                    sotuvchi_stir = data.get("seller", {}).get("vatregcode", "")

                    xaridor_nomi = data.get("buyer", {}).get("name", "")
                    xaridor_stir = data.get("buyer", {}).get("vatregcode", "")

                    mahsulotlar = data.get("productlist", {}).get("products", [])

                    for item in mahsulotlar:
                        qator = {
                            "Hujjat Raqami": faktura_raqami,
                            "Hujjat Sanasi": faktura_sanasi,
                            "Sotuvchi Tashkilot": sotuvchi_nomi,
                            "Sotuvchi STIR": sotuvchi_stir,
                            "Xaridor Tashkilot": xaridor_nomi,
                            "Xaridor STIR": xaridor_stir,
                            "Xizmat / Mahsulot Nomi": item.get("name", ""),
                            "Soni": item.get("count", 0),
                            "Yetkazib Berish Narxi (QQSsiz)": item.get("deliverysum", 0),
                            "QQS Summasi": item.get("vatsum", 0),
                            "Jami Summa (QQS bilan)": item.get("deliverysumwithvat", 0)
                        }
                        barcha_qatorlar.append(qator)

    if json_topildi == 0:
        print("ZIP arxivlar topildi, lekin ularning ichida JSON fayllar yo'q.")
        return

    # DataFrame'ga o'girish va Excelga saqlash
    df = pd.DataFrame(barcha_qatorlar)
    df.to_excel(excel_fayl_nomi, index=False)

    print(f"Muvaffaqiyatli yakunlandi! {len(zip_fayllar)} ta ZIP arxivdan {json_topildi} ta JSON fayl o'qildi va '{excel_fayl_nomi}' fayliga saqlandi.")

# Dasturni ishga tushirish
arxivdan_excelga_otkazish()