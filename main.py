import docx
from docx.enum.style import WD_STYLE_TYPE
from bs4 import BeautifulSoup
import re

def add_html_tags(docx_file, output_file):
    # פתיחת קובץ ה-DOCX
    doc = docx.Document(docx_file)
    footnote_count = 1
    has_footnotes = False
    
    # Check if document has any footnotes
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.footnote is not None:
                has_footnotes = True
                break
        if has_footnotes:
            break
    
    # יצירת אובייקט BeautifulSoup
    soup = BeautifulSoup('<html><body></body></html>', 'html.parser')
    body = soup.body

    for paragraph in doc.paragraphs:
        if paragraph.style is not None and paragraph.style.type == WD_STYLE_TYPE.PARAGRAPH and paragraph.style.name.startswith('Heading'):
            # טיפול בכותרות
            level = int(paragraph.style.name.split()[-1])
            header_tag = soup.new_tag(f'h{level}')
            
            # טיפול בהערות שוליים בכותרות
            current_text = ""
            for run in paragraph.runs:
                if has_footnotes and run.footnote is not None:
                    if current_text:
                        if re.search(r'<[^>]+>', current_text):
                            header_tag.append(BeautifulSoup(current_text, 'html.parser'))
                        else:
                            header_tag.append(current_text)
                        current_text = ""
                    sub_tag = soup.new_tag('sup')
                    sub_tag.string = str(footnote_count)
                    footnote_count += 1
                    header_tag.append(sub_tag)
                elif run.font.bold:
                    if current_text:
                        if re.search(r'<[^>]+>', current_text):
                            header_tag.append(BeautifulSoup(current_text, 'html.parser'))
                        else:
                            header_tag.append(current_text)
                        current_text = ""
                    b_tag = soup.new_tag('b')
                    b_tag.string = run.text
                    header_tag.append(b_tag)
                else:
                    current_text += run.text
            
            # Add any remaining text
            if current_text:
                if re.search(r'<[^>]+>', current_text):
                    header_tag.append(BeautifulSoup(current_text, 'html.parser'))
                else:
                    header_tag.append(current_text)
            body.append(header_tag)
            body.append('\n')
        else:
            # טיפול בפסקאות רגילות - לא יוצרים תג p
            current_text = ""
            for run in paragraph.runs:
                if has_footnotes and run.footnote is not None:
                    if current_text:
                        if re.search(r'<[^>]+>', current_text):
                            body.append(BeautifulSoup(current_text, 'html.parser'))
                        else:
                            body.append(current_text)
                        current_text = ""
                    sub_tag = soup.new_tag('sup')
                    sub_tag.string = str(footnote_count)
                    footnote_count += 1
                    body.append(sub_tag)
                elif run.font.bold:
                    if current_text:
                        if re.search(r'<[^>]+>', current_text):
                            body.append(BeautifulSoup(current_text, 'html.parser'))
                        else:
                            body.append(current_text)
                        current_text = ""
                    b_tag = soup.new_tag('b')
                    b_tag.string = run.text
                    body.append(b_tag)
                else:
                    current_text += run.text
            
            # Add any remaining text
            if current_text:
                if re.search(r'<[^>]+>', current_text):
                    body.append(BeautifulSoup(current_text, 'html.parser'))
                else:
                    body.append(current_text)
            body.append('\n')

    # שמירת התוצאה לקובץ HTML
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(str(soup).replace('<html><body>', '').replace('</body></html>', ''))

    return has_footnotes


from docx2python import *

def extract_footnotes(docx_file, output_file):
    # פתיחת קובץ ה-DOCX
    with docx2python(docx_filename=docx_file) as docx_content:
        footnotes = docx_content.footnotes
        with open(output_file, 'w', encoding='utf-8') as f:
            for i in range(2, len(footnotes[0][0])):
                # חיבור כל הפסקאות של הערת השוליים עם תג <P> ביניהן
                footnote_paragraphs = []
                for paragraph in footnotes[0][0][i]:
                    if paragraph.strip():  # רק אם הפסקה לא ריקה
                        footnote_paragraphs.append(paragraph.replace("footnote", "").strip())
                
                # כתיבת מספר הערת השוליים ואחריו כל הפסקאות מחוברות עם <P>
                if footnote_paragraphs:                 
                    f.write(f'{" <P> ".join(footnote_paragraphs)}\n')  # חיבור הפסקאות עם <P>

import re
import json
def match_footnotes(main_file, footnotes_file):
    # קריאת הקובץ הראשי
    with open(main_file, 'r', encoding='utf-8') as f:
        main_content = f.readlines()

    # קריאת קובץ ההערות
    with open(footnotes_file, 'r', encoding='utf-8') as f:
        footnotes = f.readlines()

    # מילון לשמירת התאמות
    matches = []

    # מעבר על כל שורה בקובץ ההערות
    for footnote_line in footnotes:
        # חיפוש המספר בתחילת השורה
        footnote_num = re.match(r'(\d+)', footnote_line)
        if footnote_num:
            footnote_num = footnote_num.group(1)
            # חיפוש ההתאמה בקובץ הראשי
            for i, main_line in enumerate(main_content):
                if f'<sup>{footnote_num}</sup>' in main_line:
                    matches.append({'line_index_1':i+1,'heRef_2':'הערות','path_2':footnotes_file,\
                                            'line_index_2':int(footnote_num),"Conection Type":"commentary"})
                    break

    # כתיבת התוצאות לקובץ JSON
    output_file = main_file.replace('.txt', '_links.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)



def zohar_to_otzaria(input_file):
    main_file = input_file.replace('.docx', '.txt')
    main_file_name = os.path.basename(main_file)
    has_footnotes = add_html_tags(input_file, main_file)
    
    footnotes_filename = None
    if has_footnotes:
        footnotes_file = main_file.replace(main_file_name, 'הערות על '+main_file_name)
        extract_footnotes(input_file, footnotes_file)
        match_footnotes(main_file, footnotes_file)
        footnotes_filename = f'הערות על {main_file_name}'
    
    return {
        'main_file': main_file_name,
        'footnotes_file': footnotes_filename,
        'links_file': main_file_name.replace('.txt', '_links.json') if has_footnotes else None
    }


import flet as ft
from pathlib import Path
import time
import threading
import os

def main(page: ft.Page):
    # הגדרות בסיסיות לדף
    page.title = "ממיר מסמכי וורד לאוצריא"
        
    page.padding = 30
    page.rtl = True
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.ALWAYS
    
    # משתנים גלובליים למעקב
    files_to_process = []
    is_processing = False
    
    # רכיבי ממשק
    files_text = ft.Text("לא נבחרו קבצים", size=16)
    progress = ft.ProgressBar(visible=False)
    status_text = ft.Text("")
    results = ft.Column(spacing=10)
    
    def process_file( file_path):
        """
        הפונקציה שלך לעיבוד הקובץ
        יש להחליף את הקוד כאן בקוד האמיתי שלך
        """
        return zohar_to_otzaria(str(file_path))
        
      
    
    def add_result_card(filename, output_files):
        """מוסיף כרטיס תוצאה לרשימה"""
        results.controls.append(
            ft.Card(
                content=ft.Container(
                    content=ft.Column([
                        ft.Text(filename, weight=ft.FontWeight.BOLD),
                        ft.Text("קבצים שנוצרו:", size=14),
                          
                        ft.Column([
                            ft.Text(f"• {name}: {path}")
                            for name, path in output_files.items()
                        ], spacing=2,)
                    ]),
                    padding=15
                )
            )
        )
        page.update()
    
    def process_files():
        """מעבד את כל הקבצים שנבחרו"""
        nonlocal is_processing
        is_processing = True
        results.controls.clear()
        total = len(files_to_process)
        
        progress.visible = True
        page.update()
        
        for i, file_path in enumerate(files_to_process, 1):
            # עדכון ממשק
            status_text.value = f"מעבד קובץ {i} מתוך {total}: {file_path.name}"
            progress.value = i / total
            page.update()
            
            # עיבוד הקובץ
            try:
                output_files = process_file(file_path)
                add_result_card(file_path.name, output_files)
            except Exception as e:
                add_result_card(file_path.name, {"שגיאה": str(e)})
        
        # סיום
        status_text.value = "הסתיים בהצלחה!" if total > 0 else ""
        progress.visible = False
        is_processing = False
        page.update()
    
    def pick_files_result(e: ft.FilePickerResultEvent):
        """מטפל בבחירת קבצים"""
        if not e.files:
            return
            
        nonlocal files_to_process
        files_to_process = [Path(f.path) for f in e.files]
        files_text.value = f"נבחרו {len(files_to_process)} קבצים"
        page.update()
    
    def pick_folder_result(e: ft.FilePickerResultEvent):
        """מטפל בבחירת תיקייה"""
        if not e.path:
            return
            
        nonlocal files_to_process
        folder = Path(e.path)
        files_to_process = list(folder.glob("**/*.docx"))
        files_text.value = f"נבחרו {len(files_to_process)} קבצים מהתיקייה"
        page.update()
    
    def start_processing(e):
        """מתחיל את העיבוד בthread נפרד"""
        if is_processing or not files_to_process:
            return
        threading.Thread(target=process_files, daemon=True).start()
    
    # יצירת file pickers
    pick_files_dialog = ft.FilePicker(
        on_result=pick_files_result,
      
    )
    
    pick_folder_dialog = ft.FilePicker(
        on_result=pick_folder_result
    )
    
    page.overlay.extend([pick_files_dialog, pick_folder_dialog])
    
    # הוספת כל הרכיבים לדף
    page.add(
        ft.Column([
            ft.Text("ממיר קבצי וורד לאוצריא", size=32, weight=ft.FontWeight.BOLD),
            ft.Row([
                ft.ElevatedButton(
                    "בחר קבצים",
                    icon=ft.icons.FILE_UPLOAD,
                    on_click=lambda _: pick_files_dialog.pick_files(allowed_extensions=["docx"],allow_multiple=True)
                ),
                ft.ElevatedButton(
                    "בחר תיקייה",
                    icon=ft.icons.FOLDER_OPEN,
                    on_click=lambda _: pick_folder_dialog.get_directory_path()
                ),
            ], spacing=10),
            files_text,
            ft.ElevatedButton(
                "התחל עיבוד",
                icon=ft.icons.PLAY_ARROW_ROUNDED,
                on_click=start_processing
            ),
            progress,
            status_text,
            results
        ], spacing=20)
    )

if __name__ == "__main__":
    ft.app(target=main)
