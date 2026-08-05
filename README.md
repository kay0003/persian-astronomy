# Persian Astronomy Presentation

A polished, editable PowerPoint presentation about the history of Persian astronomy, designed for Polish high-school students.

## Files

| File | Description |
|------|-------------|
| `Persian_Astronomy_Presentation.pptx` | The ready-to-open PowerPoint file (22 slides, dark night-sky theme) |
| `generate_presentation.py` | Python script that generates the `.pptx` file |
| `requirements.txt` | Python dependencies |

## How to open

Simply open `Persian_Astronomy_Presentation.pptx` with Microsoft PowerPoint, LibreOffice Impress, or Google Slides (File → Import).

All text (including Persian/Farsi characters) is fully editable native text — not images.

## How to regenerate / modify the presentation

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the script:
   ```bash
   python generate_presentation.py
   ```

This will download the latest Wikimedia images and overwrite `Persian_Astronomy_Presentation.pptx`.

> **Note:** Image downloads require internet access. If images cannot be downloaded, the script will skip them gracefully and the presentation will still be generated without photos. Re-run with internet access to include the images.

## Slides overview

1. Title slide
2. Why Persia?
3. Zoroastrianism & the Sky
4. Sacred Symbol: Faravahar
5. The Persian Calendar
6. Omar Khayyam (1048–1131)
7. Nasir al-Din al-Tusi (1201–1274)
8. The Tusi Couple
9. The Maragheh Observatory
10. Al-Biruni (973–1048)
11. The Astrolabe
12. Sky-Aligned Architecture: Persepolis
13. Chogha Zanbil Ziggurat
14. Legacy Today
15. Section divider — Let's Write in Persian!
16. Letters We Need (table)
17–20. Vocabulary words (ستاره, آسمان, ماه, خورشید)
21. Practice slide
22. Thank you / closing

