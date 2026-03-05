import json
import pandas as pd
import glob

def fakturalarni_excelga_otkazish(json_papka_yoli, excel_fayl_nomi="Fakturalar_hisoboti.xlsx"):
    barcha_qatorlar = []

    # Papkadagi barcha .json fayllarni topish
    json_fayllar = glob.glob(f"{json_papka_yoli}/*.json")
    
    if not json_fayllar:
        print("Ushbu papkada JSON fayllar topilmadi.")
        return

    for fayl in json_fayllar:
        with open(fayl, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Asosiy hujjat ma'lumotlarini ajratib olish
        faktura_raqami = data.get("facturadoc", {}).get("facturano", "")
        faktura_sanasi = data.get("facturadoc", {}).get("facturadate", "")

        sotuvchi_nomi = data.get("seller", {}).get("name", "")
        sotuvchi_stir = data.get("seller", {}).get("vatregcode", "")

        xaridor_nomi = data.get("buyer", {}).get("name", "")
        xaridor_stir = data.get("buyer", {}).get("vatregcode", "")

        mahsulotlar = data.get("productlist", {}).get("products", [])

        # Har bir xizmat/mahsulot uchun alohida qator yaratish
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

    # Yig'ilgan ma'lumotlarni Pandas DataFrame'ga aylantirish
    df = pd.DataFrame(barcha_qatorlar)

    # DataFrame'ni Excel fayl sifatida saqlash
    df.to_excel(excel_fayl_nomi, index=False)
    
    print(f"Muvaffaqiyatli yakunlandi! {len(json_fayllar)} ta fayl '{excel_fayl_nomi}' ga saqlandi.")

# Dasturni ishga tushirish qismi
# Diqqat: Bu yerda JSON fayllaringiz joylashgan papka yo'lini ko'rsating.
# Agar Python fayli bilan bir papkada bo'lsa, shunchaki '.' qoldiring.
fakturalarni_excelga_otkazish(json_papka_yoli='.')