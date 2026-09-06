
# Dars jadvali V2

Bu versiya foydalanuvchining real Excel fayliga moslashtirilgan.

## Ishga tushirish
```bash
cd dars_jadvali_app_v2
npm install
npm start
```
Brauzer:
`http://localhost:3000`

## Funksiyalar
- Excel varag‘ini spreadsheet kabi ko‘rsatadi.
- Gorizontal va vertikal scroll.
- Excel satr/ustun raqamlari.
- Asosiy katak shrift, fon, tekislash va border formatlarini ko‘rsatish.
- O‘qituvchi paralelligi.
- Sinf paralelligi.
- 1–4-sinflar: bir kunda maksimal 5 dars; 6+ bo‘lsa ogohlantirish.
- 5-sinf va yuqori: bir kunda maksimal 6 dars; 7+ bo‘lsa ogohlantirish.
- Konfliktdan keyin bo‘sh vaqt variantlari.
- “Ko‘chirish” orqali darsni yangi kun/dars qatoriga ko‘chirish.
- Original Excel fayli asosida natijaviy `.xlsx` eksport.
- Original workbook saqlanib, faqat ko‘chirilgan katak qiymatlari o‘zgartiriladi.

Eslatma: Excel’ning barcha murakkab vizual imkoniyatlari (masalan, shartli formatlash, grafikalar, ba’zi murakkab merge/render xususiyatlari) brauzerda 100% nusxalanmasligi mumkin. Lekin eksport original Excel fayli asosida amalga oshiriladi.
