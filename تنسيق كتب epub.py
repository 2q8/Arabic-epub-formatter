import os
import re
from docx import Document  # للتعامل مع ملفات docx
from ebooklib import epub  # لإنشاء ومعالجة ملفات EPUB
from bs4 import BeautifulSoup  # لتحليل وتعديل HTML داخل EPUB
#  متطلبات السكربت:
# pip install python-docx EbookLib beautifulsoup4

num = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
anum = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩']
fnum = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹']

def replace_text(text):
    # إجراء التعديلات اللازمة على النص
    text = text.replace(".  ", ".^p")
    text = text.replace("!  ", "!^p")
    text = text.replace(":  ", ":^p")
    text = text.replace("؟  ", "؟^p")
    text = text.replace("  ", " ").replace(" .", ".").replace(" ،", "،")
    text = text.replace(" ؛", "؛").replace(" :", ":").replace(" !", "!")
    text = text.replace(" ؟", "؟")

    # تغيير الأرقام العربية والفارسية إلى الأرقام الإنجليزية
    for i in range(len(num)):
        text = text.replace(anum[i], num[i])
        text = text.replace(fnum[i], num[i])
    
    # تقسيم النص إلى فقرات جديدة
    newParagraphs = [line for line in text.split("^p")]
    return newParagraphs

def process_docx(fn):
    # قراءة محتوى ملف docx
    doc = Document(fn)
    alltext = ""
    for paragraph in doc.paragraphs:
        paragraph_text = re.sub("Page [0-9]+", '', paragraph.text)  # إزالة ترقيم الصفحات
        if paragraph_text:
            alltext += paragraph_text + "  "
    
    # استبدال التكرارات والأخطاء في النصوص
    newParagraphs = replace_text(alltext)

    # إعداد ملف EPUB جديد
    book = epub.EpubBook()
    book.set_identifier('id123456')
    book.set_title('My Book')
    book.set_language('ar')

    # إعداد المحتوى بتنسيق HTML داخل ملف EPUB
    content = '<html><head><meta charset="utf-8"/></head><body dir="rtl" lang="ar">'
    for par in newParagraphs:
        content += f'<p>{par}</p>'
    content += '</body></html>'

    # إضافة المحتوى إلى ملف EPUB
    chapter = epub.EpubHtml(title='Chapter 1', file_name='chap_01.xhtml', lang='ar')
    chapter.content = content
    book.add_item(chapter)

    # إعداد الفهرس والعناصر
    book.toc = (epub.Link('chap_01.xhtml', 'Chapter 1', 'chap1'),)
    book.spine = ['nav', chapter]

    # إضافة ملفات CSS لتنسيق الكتاب
    style = 'body { font-family: Arial; font-size: 14pt; }'
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=style)
    book.add_item(nav_css)

    # حفظ الكتاب بصيغة EPUB
    epub_file = f"./@2q8/{os.path.splitext(fn)[0]}.epub"
    epub.write_epub(epub_file, book, {})
    print(f"!تم إنشاء ملف EPUB من DOCX: {epub_file}")

def process_epub(fn):
    # فتح ملف EPUB الحالي وتحليل محتوياته
    book = epub.read_epub(fn)

    for item in book.get_items():
        if item.get_type() == epub.EpubHtml:
            # استخدام BeautifulSoup لتحليل وتعديل HTML
            soup = BeautifulSoup(item.get_body_content(), 'html.parser')

            # تعديل النصوص داخل الفقرات
            for p in soup.find_all('p'):
                if p.text:
                    new_paragraphs = replace_text(p.text)
                    p.clear()
                    p.append(' '.join(new_paragraphs))

            # تحديث المحتوى الجديد
            item.set_content(str(soup).encode('utf-8'))

    # حفظ ملف EPUB المعدل
    epub_file = f"./@2q8/{os.path.splitext(fn)[0]}_modified.epub"
    epub.write_epub(epub_file, book, {})
    print(f"!تم تعديل ملف EPUB: {epub_file}")

# التأكد من وجود مجلد لحفظ الملفات المعدلة
try:
    os.mkdir("@2q8")
except FileExistsError:
    pass

# معالجة كل الملفات في المجلد الحالي (docx وepub)
for docfile in os.listdir():
    if docfile.endswith(".docx"):
        print(f"...{docfile} جاري تنسيق ملف DOCX")
        process_docx(docfile)
    elif docfile.endswith(".epub"):
        print(f"...{docfile} جاري تعديل ملف EPUB")
        process_epub(docfile)
