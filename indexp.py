from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, MessageHandler, ContextTypes, ConversationHandler, filters
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import logging

# === Konfigurasi Google Sheets API ===
SERVICE_ACCOUNT_FILE = 'config/project_credentials.json'  # Ganti dengan lokasi file JSON kredensial
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1qIKl-LaYETBs38sZgOmQdD7UDUP4a_tnQdwFNrr1w3s'  # Ganti dengan ID Spreadsheet Anda
RANGE_NAME = 'DataProject!B:M'  # Sesuaikan dengan range data Anda

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build('sheets', 'v4', credentials=credentials)

# === Konstanta untuk ConversationHandler ===
ASK_ID, ASK_COLUMN, ASK_VALUE = range(3)

# Menambahkan dictionary untuk mengonversi kode kolom menjadi nama kolom
COLUMN_NAMES = {
    "B": "ID ORDER",
    "C": "NAMA CUSTOMER",
    "D": "ALAMAT",
    "E": "JENIS CUCI",
    "F": "KATEGORI ITEM",
    "G": "STATUS ORDER",
    "H": "TOTAL BIAYA",
    "I": "TANGGAL MASUK",
    "J": "TANGGAL SELESAI",
    "K": "METODE PEMBAYARAN",
    "L": "KETERANGAN",
    "M": "KONTAK"
}

# === Fungsi untuk Membaca Data dari Spreadsheet ===
def read_spreadsheet():
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        rows = result.get('values', [])
        return rows
    except Exception as e:
        print(f"Terjadi kesalahan saat mengambil data: {e}")
        return None

# === Fungsi untuk Memperbarui Data di Spreadsheet ===
def update_spreadsheet(row, column, value):
    try:
        range_to_update = f'DataProject!{column}{row}'  # Contoh: Update cell A2
        body = {'values': [[value]]}
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_to_update,
            valueInputOption='RAW',
            body=body
        ).execute()
        return True
    except Exception as e:
        print(f"Terjadi kesalahan saat memperbarui data: {e}")
        return False

# === Handler Command /start ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("Cek Order", callback_data='cek_order')],
        [InlineKeyboardButton("Update Order", callback_data='update_order')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Pilih salah satu opsi:", reply_markup=reply_markup)

# === Handler CallbackQuery untuk Cek Data ===
async def cek_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("Silakan masukkan ID ORDER yang ingin dicari:")
    return ASK_ID

# === Handler CallbackQuery untuk Update Data ===
async def update_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("Silakan masukkan ID ORDER yang ingin diperbarui:")
    return ASK_ID

async def cek_order_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    order_id = update.message.text  # ID ORDER yang dimasukkan oleh pengguna
    data = read_spreadsheet()

    if data:
        matched_row = next((row for row in data if str(row[0]).strip() == order_id), None)

        if matched_row:
            fields = [
                "ID ORDER", "NAMA CUSTOMER", "ALAMAT", "JENIS CUCI", "KATEGORI ITEM", 
                "STATUS ORDER", "TOTAL BIAYA", "TANGGAL MASUK", "TANGGAL SELESAI", "METODE PEMBAYARAN", 
                "KETERANGAN", "KONTAK"
            ]

            message = "Data ditemukan:\n\n"
            for i, field in enumerate(fields):
                value = matched_row[i] if i < len(matched_row) and matched_row[i] else "---"
                message += f"{field}: {value}\n"
        else:
            message = f"ID ORDER {order_id} tidak ditemukan dalam data."
    else:
        message = "Gagal mengambil data dari Spreadsheet."

    await update.message.reply_text(message)
    return ConversationHandler.END

async def update_order_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    order_id = update.message.text.strip()
    context.user_data['id'] = order_id

    data = read_spreadsheet()

    if data and any(row[0].strip() == order_id for row in data):
        columns = [
            ("B", "ID ORDER"), ("C", "NAMA CUSTOMER"), ("D", "ALAMAT"), 
            ("E", "JENIS CUCI"), ("F", "KATEGORI ITEM"), ("G", "STATUS ORDER"),
            ("H", "TOTAL BIAYA"), ("I", "TANGGAL MASUK"), ("J", "TANGGAL SELESAI"),
            ("K", "METODE PEMBAYARAN"), ("L", "KETERANGAN"), ("M", "KONTAK")
        ]

        keyboard = [[InlineKeyboardButton(text=col_name, callback_data=col)] for col, col_name in columns]
        reply_markup = InlineKeyboardMarkup(keyboard)

        await update.message.reply_text(
            f"ID ORDER {order_id} ditemukan. Pilih DATA yang ingin diperbarui:",
            reply_markup=reply_markup
        )

        return ASK_COLUMN
    else:
        await update.message.reply_text(f"ID ORDER {order_id} tidak ditemukan dalam data.")
        return ConversationHandler.END

async def update_order_column(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    column = query.data
    context.user_data['column'] = column

    column_name = COLUMN_NAMES.get(column, "Unknown Column")

    await query.edit_message_text(f"Anda memilih kolom {column_name}. Masukkan value baru untuk data ini:")

    return ASK_VALUE

async def update_order_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    order_id = context.user_data['id']
    column = context.user_data['column']

    data = read_spreadsheet()
    row_number = next((i + 1 for i, row in enumerate(data) if str(row[0]).strip() == order_id), None)

    if row_number:
        range_to_update = f'DataProject!{column}{row_number}'
        body = {'values': [[value]]}

        try:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_to_update,
                valueInputOption='RAW',
                body=body
            ).execute()
            message = "Data berhasil diperbarui."
        except Exception as e:
            message = f"Gagal memperbarui data: {e}"
    else:
        message = f"ID ORDER {order_id} tidak ditemukan dalam data."

    await update.message.reply_text(message)
    return ConversationHandler.END

# === Fungsi untuk Mengakhiri Percakapan ===
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Operasi dibatalkan. Ketik /start untuk memulai kembali.")
    return ConversationHandler.END

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# === Main Program ===
if __name__ == '__main__':
    app = ApplicationBuilder().token('YOUR_BOT_TOKEN').build()  # Ganti dengan token bot Anda

    conv_handler_cek = ConversationHandler(
        entry_points=[CallbackQueryHandler(cek_order, pattern='^cek_order$')],
        states={
            ASK_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, cek_order_id)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    conv_handler_update = ConversationHandler(
        entry_points=[CallbackQueryHandler(update_order, pattern='^update_order$')],
        states={
            ASK_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, update_order_id)],
            ASK_COLUMN: [CallbackQueryHandler(update_order_column)],
            ASK_VALUE: [MessageHandler(filters.TEXT & ~filters.COMMAND, update_order_value)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler_cek)
    app.add_handler(conv_handler_update)

    print("Bot berjalan...")
    app.run_polling()
