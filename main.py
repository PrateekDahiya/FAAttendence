import openpyxl
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext
from flask import Flask, request
import logging
import os
from difflib import get_close_matches
from threading import Thread
import asyncio

# Token and Webhook URL from environment variables
TOKEN = os.getenv("YOUR_BOT_TOKEN")
WEBHOOK_URL = os.getenv("YOUR_WEBHOOK_URL")

# Logging setup
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Initialize Flask app and bot variables
app = Flask(__name__)
bot = None
application = None

# Webhook route for receiving updates from Telegram
@app.route(f'/{TOKEN}', methods=['POST'])
def telegram_webhook():
    json_str = request.get_data(as_text=True)
    update = Update.de_json(json_str, bot)
    application.update_handler(update)
    return 'ok'

# Function to set the webhook URL for the bot
async def set_webhook():
    webhook_url = f"{WEBHOOK_URL}/{TOKEN}"
    await bot.set_webhook(webhook_url)

# Function to start Flask in a separate thread
def start_flask():
    app.run(host='0.0.0.0', port=5000)

# Function to start the bot and Flask
def main():
    global bot, application

    # Initialize the bot and application with token
    bot = Application.builder().token(TOKEN).build()
    application = Application.builder().token(TOKEN).build()

    # Adding command and message handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("sendfile", send_file))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error)

    # Set webhook
    asyncio.run(set_webhook())

    # Start Flask in a separate thread
    Thread(target=start_flask).start()

    # Keep the event loop running indefinitely
    asyncio.get_event_loop().run_forever()

# Start command handler
async def start(update: Update, context: CallbackContext):
    await update.message.reply_text('Hello! Send me a message, and I will respond. Use /sendfile to get the file.')

# Command to send a file (e.g., Excel file)
async def send_file(update: Update, context: CallbackContext):
    chat_id = update.message.chat_id
    file_path = './FAAtt.xlsx'
    try:
        with open(file_path, 'rb') as file:
            await context.bot.send_document(chat_id=chat_id, document=file)
            await update.message.reply_text("Here is the file.")
    except Exception as e:
        logging.error(f"Failed to send file: {e}")
        await update.message.reply_text("Failed to send the file.")

# Help command handler
async def help_command(update: Update, context: CallbackContext):
    help_text = '''
    Here are the available commands:
    /start - Start the bot
    /sendfile - Get the attendance file
    '''
    await update.message.reply_text(help_text)

# Handler for plain text messages (roll numbers or names for attendance)
async def handle_message(update: Update, context: CallbackContext):
    text = update.message.text.lower()
    if text.startswith('/sendfile'):
        await send_file(update, context)
    else:
        try:
            # Assuming roll number or name processing function here
            roll_number = int(text)
            res = await update_attendance(roll_number)
            if res != "success":
                await update.message.reply_text(res)
        except ValueError:
            await update.message.reply_text(
                "Invalid input. Please enter a valid roll number or use the /sendfile command."
            )

# Error handler
async def error(update: Update, context: CallbackContext):
    logging.error(f'Update {update} caused error {context.error}')

# Placeholder for update_attendance function (you can implement this as per your logic)
async def update_attendance(roll_number):
    # Mock function for demonstration purposes
    return "success"

# Main entry point
if __name__ == '__main__':
    main()
