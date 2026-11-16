"""Build the three documents of practical work 3 as separate .docx files.

The documents are filled in from the templates handed out in class: consent to personal
data processing, cookie notice, consent to advertising mailings. The company data is shared
between them and lives in the constants below, so a change is made once.

Usage:
    python .claude/tools/build_consent_documents.py
"""

from pathlib import Path

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

import report_base as rb

OUT_DIR = Path(__file__).resolve().parents[2] / "Task 3rd Personal Data Consent" / "Solution"

COMPANY = "ООО «Дрим Сникерс»"
INN = "6164280573"
ADDRESS = "344002, г. Ростов-на-Дону, ул. Большая Садовая, д. 70, офис 12"
SITE = "https://dreamsneakers.ru"
EMAIL = "info@dreamsneakers.ru"

CONSENT_TITLE = "Согласие на обработку персональных данных"
CONSENT = [
    f"Я (далее – Заявитель) даю согласие {COMPANY}, ИНН {INN}, адрес {ADDRESS} (далее – "
    f"Оператор) на автоматизированную, а также без использования средств автоматизации, "
    f"обработку моих персональных данных, предоставленных мною Оператору посредством "
    f"сети Интернет через сайт {SITE}, а именно: фамилия, имя; номер телефона; адрес "
    f"электронной почты; адрес доставки,",

    "включая сбор, запись, систематизацию, накопление, хранение, уточнение "
    "(обновление, изменение), извлечение, использование, передачу (предоставление, "
    "доступ), блокирование, удаление и уничтожение, а также на вышеуказанную обработку "
    "иных моих персональных данных, полученных в результате их обработки, включая "
    "обработку третьими лицами.",

    "Целями обработки персональных данных является: оформление и исполнение заказов, "
    "заключение и исполнение договора розничной купли-продажи, доставка товаров, связь "
    "со мной по вопросам заказа, ответы на мои обращения через форму обратной связи.",

    "Указанное согласие дано на срок три года с даты его предоставления.",

    "В случае, если в течение указанного срока с Заявителем будет заключен договор "
    "розничной купли-продажи, настоящее Согласие Заявителя сохраняет силу в течение "
    "всего срока действия каждого договора, а также в течение трёх лет с даты "
    "прекращения обязательств сторон по договорам.",

    "Я уведомлен, что предоставление указанного выше согласия является моим правом, а "
    "не обязанностью, данные согласия могут быть отозваны в порядке, предусмотренном "
    "законодательством Российской Федерации. Согласие может быть отозвано субъектом "
    "персональных данных путем обращения к Оператору с заявлением, оформленным в "
    "письменной форме.",
]

COOKIE_TEXT = (
    f"{COMPANY} обрабатывает файлы cookie. Они помогают нам делать этот сайт удобнее для "
    f"пользователей. Продолжая работу с сайтом {SITE}, вы соглашаетесь с обработкой "
    f"файлов cookie вашего браузера. Однако вы можете запретить обработку некоторых "
    f"типов файлов cookie в настройках вашего браузера либо на странице «Уведомление об "
    f"использовании файлов cookie»."
)

ADS_TITLE = "Согласие на получение рекламных рассылок"
ADS = [
    f"Я (далее – Заявитель) даю согласие {COMPANY}, ИНН {INN}, адрес {ADDRESS} (далее – "
    f"Владелец сайта) на направление мне на указанные мной на сайте {SITE} контактные "
    f"данные (номер телефона и/или электронную почту) сообщений в информационных, "
    f"рекламно-информационных целях об услугах (сервисах) Владельца сайта, а именно: о "
    f"новых поступлениях кроссовок, об акциях, скидках и распродажах, об открытии и "
    f"режиме работы магазинов, в виде sms-сообщений; электронных писем; сообщений в "
    f"мессенджере Telegram.",

    "Настоящим я гарантирую, что указанные мною на Сайте абонентские номера телефонов "
    "принадлежат мне, в случае прекращения использования мною абонентского номера "
    "телефона, указанного на Сайте, я обязуюсь проинформировать об этом Владельца сайта.",

    f"Рекламная рассылка осуществляется до отписки физического лица от информации "
    f"путем перехода по специальной ссылке «Отписаться от рассылки» в рассылаемых "
    f"письмах или путем направления соответствующего запроса на электронную почту "
    f"{EMAIL}.",
]


def centered(doc, text):
    rb.para(doc, text, indent=Cm(0), align=WD_ALIGN_PARAGRAPH.CENTER)


def text_document(title, paragraphs):
    doc = rb.new_document()
    centered(doc, title)
    rb.blank(doc)
    for text in paragraphs:
        rb.para(doc, text)
        rb.blank(doc)
    return doc


def cell_text(cell, text, width):
    cell.width = Cm(width)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if text:
        run = p.add_run(text)
        run.font.name = rb.FONT
        run.font.size = rb.SIZE


def cookie_document():
    widths = [11.0, 5.5]
    doc = rb.new_document()
    t = doc.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    rows = [(COOKIE_TEXT, "Согласен"), ("", "Настрою cookies сам")]
    for row, (left, right) in zip(t.rows, rows):
        cell_text(row.cells[0], left, widths[0])
        cell_text(row.cells[1], right, widths[1])
    # python-docx puts the table first, the empty paragraph after it is harmless
    return doc


def build():
    rb.save(text_document(CONSENT_TITLE, CONSENT),
            OUT_DIR / "Согласие на обработку персональных данных.docx")
    rb.save(cookie_document(),
            OUT_DIR / "Уведомление об использовании cookie.docx")
    rb.save(text_document(ADS_TITLE, ADS),
            OUT_DIR / "Согласие на получение рекламных рассылок.docx")


if __name__ == "__main__":
    build()
