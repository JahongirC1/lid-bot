"""
GURUHLAR, USTOZLAR VA BO'SH XONALAR
-----------------------------------
O'zgartirish kerak bo'lsa - faqat shu faylni tahrirlang va GitHub'ga yuklang.

GURUHLAR - har bir guruh:
  ustoz   - ustoz ismi (USTOZLAR dagi bilan bir xil yozilsin)
  nomi    - guruh nomi (kartochkadagi sarlavha)
  xona    - xona
  kunlar  - "toq" (Du/Chor/Ju) yoki "juft" (Se/Pay/Shan)
  soat    - boshlanish soati, "HH:MM"
  davom   - dars davomiyligi, daqiqada (yozilmasa 120)
  tur     - "umumiy", "cefr", "tanal", "grammatika"
  izoh    - kartochkadagi daraja/nom (O'rta, Boshlang'ich, B1 ...)
  daraja  - shu guruhga mos lid darajalari (a0, a1, a2, b1, b2, c1)
  talaba  - hozirgi talabalar soni (kam bo'lsa - birinchi taklif qilinadi)
"""

GURUHLAR = [
    # ---------------- Akbarali Ustoz
    {"ustoz": "Akbarali", "nomi": "Toq 11:00", "xona": "4-xona", "kunlar": "toq", "soat": "11:00",
     "tur": "umumiy", "izoh": "O'rta", "daraja": ["a2", "b1"], "talaba": 15},
    {"ustoz": "Akbarali", "nomi": "Juft 11:00", "xona": "4-xona", "kunlar": "juft", "soat": "11:00",
     "tur": "umumiy", "izoh": "Boshlang'ich", "daraja": ["a0", "a1", "a2"], "talaba": 16},
    {"ustoz": "Akbarali", "nomi": "Juft 14:00", "xona": "4-xona", "kunlar": "juft", "soat": "14:00",
     "tur": "umumiy", "izoh": "B1", "daraja": ["b1"], "talaba": 17},
    {"ustoz": "Akbarali", "nomi": "Toq 14:00", "xona": "4-xona", "kunlar": "toq", "soat": "14:00",
     "tur": "umumiy", "izoh": "O'rta", "daraja": ["a2", "b1"], "talaba": 8},
    {"ustoz": "Akbarali", "nomi": "Toq 16:00", "xona": "4-xona", "kunlar": "toq", "soat": "16:00",
     "tur": "umumiy", "izoh": "O'rta", "daraja": ["a2", "b1"], "talaba": 19},
    {"ustoz": "Akbarali", "nomi": "Juft 16:00", "xona": "4-xona", "kunlar": "juft", "soat": "16:00",
     "tur": "umumiy", "izoh": "Boshlang'ich", "daraja": ["a0", "a1", "a2"], "talaba": 8},

    # ---------------- Jahongir Ustoz
    {"ustoz": "Jahongir", "nomi": "CEFR", "xona": "9-xona", "kunlar": "juft", "soat": "16:00",
     "tur": "cefr", "izoh": "CEFR", "daraja": ["a2", "b1", "b2", "c1"], "talaba": 5},
    {"ustoz": "Jahongir", "nomi": "Toq 18:00", "xona": "4-xona", "kunlar": "toq", "soat": "18:00",
     "tur": "umumiy", "izoh": "Boshlang'ich", "daraja": ["a0", "a1", "a2"], "talaba": 6},

    # ---------------- Dildora Ustoz
    {"ustoz": "Dildora", "nomi": "Toq 08:00", "xona": "9-xona", "kunlar": "toq", "soat": "08:00",
     "tur": "tanal", "izoh": "TANAL B2-C1", "daraja": ["b2", "c1"], "talaba": 3},
    {"ustoz": "Dildora", "nomi": "Toq 11:00", "xona": "9-xona", "kunlar": "toq", "soat": "11:00",
     "tur": "umumiy", "izoh": "Boshlang'ich Izofa", "daraja": ["a1", "a2"], "talaba": 4},
    {"ustoz": "Dildora", "nomi": "Juft 11:00", "xona": "9-xona", "kunlar": "juft", "soat": "11:00",
     "tur": "tanal", "izoh": "TANAL", "daraja": ["a2", "b1", "b2"], "talaba": 5},
    {"ustoz": "Dildora", "nomi": "Toq 16:00", "xona": "9-xona", "kunlar": "toq", "soat": "16:00",
     "tur": "grammatika", "izoh": "Ibrohimov fe'lda", "daraja": ["a2", "b1"], "talaba": 5},
    {"ustoz": "Dildora", "nomi": "Toq 17:30", "xona": "9-xona", "kunlar": "toq", "soat": "17:30",
     "tur": "tanal", "izoh": "TANAL A2-B1", "daraja": ["a2", "b1"], "talaba": 0},

    # ---------------- Abu Bakr Ustoz
    {"ustoz": "Abu Bakr", "nomi": "Juft 14:00", "xona": "9-xona", "kunlar": "juft", "soat": "14:00",
     "tur": "umumiy", "izoh": "O'rta", "daraja": ["a2", "b1"], "talaba": 5},
    {"ustoz": "Abu Bakr", "nomi": "Toq 14:00", "xona": "9-xona", "kunlar": "toq", "soat": "14:00",
     "tur": "umumiy", "izoh": "Boshlang'ich", "daraja": ["a0", "a1", "a2"], "talaba": 6},
]

# USTOZLAR - yangi guruh ochishda kim qaysi kurs va vaqtga dars bera oladi
#   turlar  - qaysi kurslardan dars beradi
#   vaqtlar - "ertalab", "kunduzi", "kechki"
USTOZLAR = {
    "Mubina":   {"turlar": ["tanal", "cefr"], "vaqtlar": ["ertalab"]},
    "Robiya":   {"turlar": ["tanal", "cefr"], "vaqtlar": ["kunduzi", "kechki"]},
    "Akbarali": {"turlar": ["umumiy"], "vaqtlar": ["ertalab", "kunduzi", "kechki"]},
    "Jahongir": {"turlar": ["cefr", "umumiy"], "vaqtlar": ["ertalab", "kunduzi", "kechki"]},
    "Dildora":  {"turlar": ["tanal", "umumiy", "grammatika"], "vaqtlar": ["ertalab", "kunduzi", "kechki"]},
    "Abu Bakr": {"turlar": ["umumiy"], "vaqtlar": ["ertalab", "kunduzi", "kechki"]},
}

# BO'SH XONALAR
#   ustoz - qavs ichida ustoz yozilgan bo'lsa, o'sha vaqt shu ustozga ajratilgan
#           (faqat o'sha ustoz bilan taklif qilinadi); "" - hamma uchun bo'sh
# Agar shu vaqtda xonada guruh dars o'tayotgan bo'lsa, bot buni o'zi aniqlab, bo'sh deb hisoblamaydi.
XONA_SIGIMI = {"4-xona": 30, "9-xona": 25, "5-xona": 15, "12-xona": 9}

XONALAR = [
    # 4-xona (30 kishilik)
    {"xona": "4-xona", "kunlar": "toq", "vaqt": "09:00-11:00", "ustoz": "Akbarali"},
    {"xona": "4-xona", "kunlar": "toq", "vaqt": "14:00-16:00", "ustoz": "Akbarali", "izoh": "0 dan"},
    {"xona": "4-xona", "kunlar": "toq", "vaqt": "18:00-20:00", "ustoz": "Jahongir", "izoh": "izofa"},
    {"xona": "4-xona", "kunlar": "juft", "vaqt": "09:00-11:00", "ustoz": "Akbarali"},
    {"xona": "4-xona", "kunlar": "juft", "vaqt": "18:00-20:00", "ustoz": ""},

    # 9-xona (25 kishilik)
    {"xona": "9-xona", "kunlar": "toq", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "9-xona", "kunlar": "toq", "vaqt": "18:00-20:00", "ustoz": ""},
    {"xona": "9-xona", "kunlar": "juft", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "9-xona", "kunlar": "juft", "vaqt": "16:00-18:00", "ustoz": "Jahongir", "izoh": "CEFR"},
    {"xona": "9-xona", "kunlar": "juft", "vaqt": "18:00-20:00", "ustoz": ""},

    # 5-xona (15 kishilik)
    {"xona": "5-xona", "kunlar": "toq", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "toq", "vaqt": "11:00-13:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "toq", "vaqt": "14:00-16:00", "ustoz": "Robiya"},
    {"xona": "5-xona", "kunlar": "toq", "vaqt": "16:00-18:00", "ustoz": "Robiya"},
    {"xona": "5-xona", "kunlar": "toq", "vaqt": "18:00-20:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "juft", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "juft", "vaqt": "11:00-13:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "juft", "vaqt": "14:00-16:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "juft", "vaqt": "16:00-18:00", "ustoz": ""},
    {"xona": "5-xona", "kunlar": "juft", "vaqt": "18:00-20:00", "ustoz": ""},

    # 12-xona (9 kishilik)
    {"xona": "12-xona", "kunlar": "toq", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "toq", "vaqt": "11:00-13:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "toq", "vaqt": "14:00-16:00", "ustoz": "Robiya"},
    {"xona": "12-xona", "kunlar": "toq", "vaqt": "16:00-18:00", "ustoz": "Robiya"},
    {"xona": "12-xona", "kunlar": "toq", "vaqt": "18:00-20:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "juft", "vaqt": "09:00-11:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "juft", "vaqt": "11:00-13:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "juft", "vaqt": "14:00-16:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "juft", "vaqt": "16:00-18:00", "ustoz": ""},
    {"xona": "12-xona", "kunlar": "juft", "vaqt": "18:00-20:00", "ustoz": ""},
]

# Lid formasidagi "Qaysi vaqtga" javobi -> boshlanish soati qaysi oraliqda bo'lsa mos
VAQT_ORALIQ = {
    "ertalab": ("06:00", "12:00"),
    "kunduzi": ("12:00", "16:00"),
    "kechki": ("16:00", "22:00"),
}

# Lid formasidagi "Qaysi kursga" javobi -> guruh turi
KURS_TURI = {
    "cefr": "cefr",
    "tanal": "tanal",
    "grammatika": "grammatika",
    "ibrohimov": "grammatika",
}

# Nechta guruh / yangi guruh varianti taklif qilinsin
TAKLIF_SONI = 3
