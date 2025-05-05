# bot.py
import os
import re
import logging
from datetime import datetime
from io import BytesIO
from typing import Dict, List, Optional

from telegram import Update, ForceReply
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters
from pdfplumber import open as pdf_open
import pandas as pd
import pytesseract
from PIL import Image
import io

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# Токен вашего бота
BOT_TOKEN = "7945896025:AAGnPOIU5C5DMl6YtfVgbIzScR-n12vcaIw"

# Регулярные выражения для извлечения данных
INN_PATTERN = r'(?:ИНН\s*)?(\d{10}|\d{12})'
DATE_PATTERN = r'«(\d{1,2})»\s+([А-Яа-я]+)\s+(\d{4})'
LEGAL_NAME_PATTERN = r'(ООО|АО|ПАО)\s+["«]?([А-Я][а-яА-ЯёЁ\-\s]+)[»"]?'
OUTGOING_NUMBER_PATTERN = r'(?:Исх\.?\s*(?:№|N)?\s*)([\w\d\-/]+)\s+(?:от|from)'

class PDFProcessor:
    @staticmethod
    def extract_text(pdf_path: str) -> str:
        """Извлекает текст из PDF-файла"""
        try:
            with pdf_open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
                return text
        except Exception as e:
            logging.error(f"Ошибка извлечения текста из PDF: {e}")
            try:
                img = Image.open(pdf_path)
                text = pytesseract.image_to_string(img, lang='rus')
                return text

            except Exception as e:
                logging.error(f"Ошибка OCR: {e}")
                return ""

    @staticmethod
    def extract_data(text: str) -> Dict[str, Optional[str]]:
        """Извлекает данные из текста"""
        # Извлечение ИНН
        inn_match = re.search(INN_PATTERN, text)
        inn = inn_match.group(1) if inn_match else None

        # Извлечение даты подписания
        date_match = re.search(DATE_PATTERN, text)
        date = None
        if date_match:
            try:
                day = date_match.group(1)
                month_name = date_match.group(2).lower()
                year = date_match.group(3)
                
                month_map = {
                    'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04',
                    'мая': '05', 'июня': '06', 'июля': '07', 'августа': '08',
                    'сентября': '09', 'октября': '10', 'ноября': '11', 'декабря': '12'
                }
                
                if month_name in month_map:
                    month = month_map[month_name]
                    date_str = f"{day.zfill(2)}.{month}.{year}"
                    date = datetime.strptime(date_str, '%d.%m.%Y').strftime('%d.%m.%Y')
            except (ValueError, KeyError):
                pass

        # Извлечение названия юрлица
        legal_name_match = re.search(LEGAL_NAME_PATTERN, text)
        legal_name = legal_name_match.group(2) if legal_name_match else None

        # Извлечение исходящего номера
        outgoing_match = re.search(OUTGOING_NUMBER_PATTERN, text)
        outgoing_number = outgoing_match.group(1) if outgoing_match else None

        return {
            'inn': inn,
            'date': date,
            'legal_name': legal_name,
            'outgoing_number': outgoing_number
        }

class ExcelExporter:
    @staticmethod
    def create_excel(data: List[Dict]) -> BytesIO:
        """Создает Excel-файл из данных"""
        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

class TelegramBot:
    def __init__(self):
        self.processor = PDFProcessor()
        self.exporter = ExcelExporter()
        self.processed_files = []

    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Обработчик команды /start"""
        user = update.effective_user
        await update.message.reply_html(
            rf"Привет, {user.mention_html()}! Отправьте мне PDF-файлы с гарантийными письмами.",
            reply_markup=ForceReply(selective=True),
        )

    async def handle_document(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Обработчик загрузки документов"""
        document = update.message.document

        if not document.file_name.lower().endswith('.pdf'):
            await update.message.reply_text("Пожалуйста, загрузите файл в формате PDF")
            return

        # Ограничение на количество файлов (до 20)
        if len(self.processed_files) >= 20:
            await update.message.reply_text("Максимальное количество файлов - 20")
            return

        try:
            # Скачиваем файл
            file = await document.get_file()
            file_path = f"downloads/{document.file_id}.pdf"

            os.makedirs("downloads", exist_ok=True)
            await file.download_to_drive(file_path)

            # Извлекаем текст
            text = self.processor.extract_text(file_path)

            # Извлекаем данные
            data = self.processor.extract_data(text)
            data['filename'] = document.file_name

            # Сохраняем результат и сразу создаем Excel
            self.processed_files.append(data)

            # Создаем и отправляем Excel файл после каждой обработки
            excel_file = self.exporter.create_excel(self.processed_files)
            await update.message.reply_document(
                document=excel_file,
                filename="garanty_letters_data.xlsx",
                caption=f"Данные из гарантийных писем (обработано файлов: {len(self.processed_files)})"
            )
            await update.message.reply_text(f"Файл {document.file_name} обработан и добавлен в Excel")

        except Exception as e:
            logging.error(f"Ошибка обработки файла: {e}")
            await update.message.reply_text("Произошла ошибка при обработке файла")

    async def send_excel(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Отправляет Excel-файл с результатами"""
        if not self.processed_files:
            await update.message.reply_text("Нет обработанных файлов для экспорта")
            return

        try:
            excel_file = self.exporter.create_excel(self.processed_files)
            await update.message.reply_document(
                document=excel_file,
                filename="garanty_letters_data.xlsx",
                caption="Данные из гарантийных писем"
            )
            self.processed_files.clear()
        except Exception as e:
            logging.error(f"Ошибка создания Excel-файла: {e}")
            await update.message.reply_text("Произошла ошибка при создании Excel-файла")

    async def clear_data(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Очищает накопленные данные"""
        self.processed_files.clear()
        await update.message.reply_text("Накопленные данные очищены")

    def run(self):
        """Запуск бота"""
        application = Application.builder().token(BOT_TOKEN).build()

        # Регистрация обработчиков
        application.add_handler(CommandHandler("start", self.start))
        application.add_handler(CommandHandler("clear", self.clear_data))
        application.add_handler(MessageHandler(filters.Document.ALL, self.handle_document))
        application.add_handler(CommandHandler("send", self.send_excel))

        # Запуск бота
        application.run_polling()

if __name__ == "__main__":
    bot = TelegramBot()
    bot.run()# PDF_bot
