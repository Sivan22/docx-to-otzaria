import docx
from docx.enum.style import WD_STYLE_TYPE
from bs4 import BeautifulSoup
import re
import json
import os
import threading
from pathlib import Path
import flet as ft
from docx2python import docx2python

def add_html_tags(docx_file, output_file):
    """
    מעבד את הקובץ הראשי וממיר את הטקסט ל-HTML.
    שימו לב: טיפול בהערות שוליים מתבצע בנפרד (באמצעות docx2python).
    """
    doc = docx.Document(docx_file)
    # ניצור מסמך HTML בסיסי
    soup = BeautifulSoup('<html><body></body></html>', 'html.parser')
    body = soup.body

    for paragraph in doc.paragraphs:
        # טיפול בכותרות (אם הסגנון מתחיל ב-"Heading")
        if paragraph.style is not None and paragraph.style.type == WD_STYLE_TYPE.PARAGRAPH and paragraph.style.name.startswith('Heading'):
            try:
                level = int(paragraph.style.name.split()[-1])
            except:
                level = 1
            header_tag = soup.new_tag(f'h{level}')
            current_text = ""
            for run in paragraph.runs:
                # אין טיפול בהערות שוליים כאן – נעבד אותן בנפרד
                if run.font.bold:
                    if current_text:
                        header_tag.append(current_text)
                        current_text = ""
                    b_tag = soup.new_tag('b')
                    b_tag.string = run.text
                    header_tag.append(b_tag)
                else:
                    current_text += run.text
            if current_text:
                header_tag.append(current_text)
            body.append(header_tag)
            body.append('\n')
        else:
            # טיפול בפסקאות רגילות
            current_text = ""
            for run in paragraph.runs:
                if run.font.bold:
                    if current_text:
                        body.append(current_text)
                        current_text = ""
                    b_tag = soup.new_tag('b')
                    b_tag.string = run.text
                    body.append(b_tag)
                else:
                    current_text += run.text
            if current_text:
                body.append(current_text)
            body.append('\n')

    with open(output_file, 'w', encoding='utf-8') as f:
        html_output = str(soup).replace('<html><body>', '').replace('</body></html>', '')
        f.write(html_output)
    # הפונקציה מחזירה False – אין טיפול בהערות שוליים כאן
    return False

def extract_footnotes(docx_file, output_file):
    """
    מפיק את הערות השוליים מהקובץ באמצעות docx2python וכותב אותן לקובץ.
    """
    with docx2python(docx_file) as docx_content:
        footnotes = docx_content.footnotes
        with open(output_file, 'w', encoding='utf-8') as f:
            for i in range(2, len(footnotes[0][0])):
                footnote_paragraphs = []
                for paragraph in footnotes[0][0][i]:
                    if paragraph.strip():
                        footnote_paragraphs.append(paragraph.replace("footnote", "").strip())
                if footnote_paragraphs:
                    f.write(f'{" <P> ".join(footnote_paragraphs)}\n')

def match_footnotes(main_file, footnotes_file):
    """
    מתאימה בין סימניות בהערות לבין הטקסט הראשי וכותבת קובץ JSON עם הקישורים.
    """
    with open(main_file, 'r', encoding='utf-8') as f:
        main_content = f.readlines()
    with open(footnotes_file, 'r', encoding='utf-8') as f:
        footnotes = f.readlines()
    matches = []
    for footnote_line in footnotes:
        footnote_num = re.match(r'(\d+)', footnote_line)
        if footnote_num:
            footnote_num = footnote_num.group(1)
            for i, main_line in enumerate(main_content):
                if f'<sup>{footnote_num}</sup>' in main_line:
                    matches.append({
                        'line_index_1': i+1,
                        'heRef_2': 'הערות',
                        'path_2': footnotes_file,
                        'line_index_2': int(footnote_num),
                        "Conection Type": "commentary"
                    })
                    break
    output_file = main_file.replace('.txt', '_links.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)

def zohar_to_otzaria(input_file):
    """
    מעבד את קובץ ה-DOCX:
    - ממיר את הטקסט הראשי ל-HTML ושומר כקובץ .txt.
    - אם קיימות הערות שוליים, מפיק קובץ הערות וקובץ קישורים.
    """
    main_file = input_file.replace('.docx', '.txt')
    main_file_name = os.path.basename(main_file)
    # המרת הטקסט הראשי
    has_footnotes = add_html_tags(input_file, main_file)
    
    # בדיקה האם קיימות הערות שוליים באמצעות docx2python
    with docx2python(input_file) as docx_content:
        docx_footnotes_count = len(docx_content.footnotes[0][0])
        has_actual_footnotes = docx_footnotes_count > 2  # נניח שאם יש יותר מ-2, קיימות הערות שוליים

    footnotes_filename = None
    if has_actual_footnotes:
        footnotes_file = main_file.replace(main_file_name, 'הערות על ' + main_file_name)
        extract_footnotes(input_file, footnotes_file)
        match_footnotes(main_file, footnotes_file)
        footnotes_filename = f'הערות על {main_file_name}'

    return {
        'הקובץ הראשי': main_file_name,
        'קובץ הערות': footnotes_filename if footnotes_filename is not None else "אין בקובץ הערות שוליים, לא נוצר קובץ",
        'קובץ קישור בין המסמך להערות (links)': main_file_name.replace('.txt', '_links.json') if has_actual_footnotes else "לא נוצר קובץ"
    }

def main(page: ft.Page):
    page.title = "ממיר מסמכי וורד לאוצריא"
    page.padding = 30
    page.rtl = True
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.ALWAYS

    files_to_process = []
    is_processing = False
    found_footnotes = False  # יהפוך ל-True אם לפחות קובץ אחד כולל הערות שוליים

    files_text = ft.Text("לא נבחרו קבצים", size=16)
    progress = ft.ProgressBar(visible=False)
    status_text = ft.Text("")
    results = ft.Column(spacing=10)

    # הגדרת כפתור "התחל עיבוד" כלא פעיל (disabled) בהתחלה
    start_button = ft.ElevatedButton(
        "התחל עיבוד",
        icon=ft.Icons.PLAY_ARROW_ROUNDED,
        disabled=True,
        on_click=lambda _: process_files()
    )

    def process_file(file_path):
        return zohar_to_otzaria(str(file_path))

    def add_result_card(filename, output_files):
        card = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text(filename, weight=ft.FontWeight.BOLD),
                    ft.Text("קבצים שנוצרו:", size=14),
                    ft.Column(
                        [ft.Text(f"• {name}: {path}") for name, path in output_files.items()],
                        spacing=2
                    )
                ]),
                padding=15
            )
        )
        results.controls.append(card)
        page.update()

    def process_files():
        nonlocal is_processing, found_footnotes
        is_processing = True
        results.controls.clear()
        total = len(files_to_process)
        progress.visible = True
        page.update()

        found_footnotes = False
        for i, file_path in enumerate(files_to_process, 1):
            status_text.value = f"מעבד קובץ {i} מתוך {total}: {file_path.name}"
            progress.value = i / total
            page.update()

            try:
                output_files = process_file(file_path)
                add_result_card(file_path.name, output_files)
                if (output_files["קובץ הערות"] != "אין בקובץ הערות שוליים, לא נוצר קובץ" and
                    output_files["קובץ קישור בין המסמך להערות (links)"] != "לא נוצר קובץ"):
                    found_footnotes = True
            except Exception as e:
                add_result_card(file_path.name, {"שגיאה": str(e)})

        if total > 0:
            base_msg = "הסתיים בהצלחה!"
            if found_footnotes:
                base_msg += "\nאת הקובץ שמסתיים ב - links.json יש להכניס לתיקיית links שבתוך התיקייה הראשית של אוצריא"
            status_text.value = base_msg
        else:
            status_text.value = ""
        progress.visible = False
        is_processing = False
        page.update()

    def pick_files_result(e: ft.FilePickerResultEvent):
        nonlocal files_to_process
        if not e.files:
            return
        files_to_process = [Path(f.path) for f in e.files]
        files_text.value = f"נבחרו {len(files_to_process)} קבצים"
        start_button.disabled = (len(files_to_process) == 0)
        page.update()

    def pick_folder_result(e: ft.FilePickerResultEvent):
        nonlocal files_to_process
        if not e.path:
            return
        folder = Path(e.path)
        files_to_process = list(folder.glob("**/*.docx"))
        files_text.value = f"נבחרו {len(files_to_process)} קבצים מהתיקייה"
        start_button.disabled = (len(files_to_process) == 0)
        page.update()

    pick_files_dialog = ft.FilePicker(on_result=pick_files_result)
    pick_folder_dialog = ft.FilePicker(on_result=pick_folder_result)

    page.overlay.extend([pick_files_dialog, pick_folder_dialog])
    page.add(
        ft.Column([
            ft.Text("ממיר קבצי וורד לאוצריא", size=32, weight=ft.FontWeight.BOLD),
            ft.Row([
                ft.ElevatedButton(
                    "בחר קבצים",
                    icon=ft.Icons.FILE_UPLOAD,
                    on_click=lambda _: pick_files_dialog.pick_files(
                        allowed_extensions=["docx"],
                        allow_multiple=True
                    )
                ),
                ft.ElevatedButton(
                    "בחר תיקייה",
                    icon=ft.Icons.FOLDER_OPEN,
                    on_click=lambda _: pick_folder_dialog.get_directory_path()
                ),
            ], spacing=10),
            files_text,
            start_button,
            progress,
            status_text,
            results
        ], spacing=20)
    )

if __name__ == "__main__":
    ft.app(target=main)
