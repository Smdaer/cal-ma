# O bir defineci — TikTok video listesi

`O_bir_defineci_videolar.xlsx` dosyası, paylaşılan TikTok ses/video meta verilerini Excel listesine dönüştürür.

## Dosyalar

- `O_bir_defineci_videolar.xlsx` — Ana liste (293 video)
  - **O bir defineci** sayfası: No, Sanatçı, Parça, Açıklama/Hashtagler, Görüntülenme
  - **Özet** sayfası: toplam ve en çok görüntülenen 20 video
- `parse_tiktok_list.py` — Listeyi yeniden üretmek için betik

## Yeniden oluşturma

```bash
pip install openpyxl
python3 parse_tiktok_list.py
```

Not: Kaynak metinde gerçek TikTok URL’leri yoktu; sanatçı, parça, açıklama ve görüntülenme alanları çıkarılmıştır.
