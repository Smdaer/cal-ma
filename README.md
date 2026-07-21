# O bir defineci — TikTok video listesi

Paylaşılan TikTok ses/video meta verileri Excel listesine dönüştürülmüştür (**293 video**).

## Dosyalar (açmak için)

| Dosya | Nasıl açılır |
|---|---|
| `O_bir_defineci_videolar.csv` | Excel’de çift tıklayın (Türkçe Excel için `;` ayracı, UTF-8) |
| `O_bir_defineci_videolar_virgullu.csv` | Virgüllü CSV alternatif |
| `O_bir_defineci_liste.xlsx` | Excel / Google Sheets (xlsxwriter) |
| `O_bir_defineci_videolar.xlsx` | Excel / Google Sheets (openpyxl) |

> XLSX açılmazsa **CSV** dosyasını kullanın: Excel → Veri → Metinden/CSV’den veya doğrudan çift tık.

## Sütunlar

1. No  
2. Sanatçı / Ses Sahibi  
3. Parça / Ses Adı  
4. Açıklama / Hashtagler  
5. Görüntülenme  
6. Görüntülenme (Sayı)  
7. Oluşturan  

## Yeniden oluşturma

```bash
pip install -r requirements.txt
python3 parse_tiktok_list.py
```

Not: Kaynak metinde gerçek TikTok URL’leri yoktu; sanatçı, parça, açıklama ve görüntülenme alanları çıkarılmıştır.
