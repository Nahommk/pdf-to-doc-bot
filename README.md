# PDF to DOC Telegram Bot

A Telegram bot that converts PDF files to Microsoft Word DOC format (97-2003).

## Features
- ✅ Convert PDF to DOC format
- ✅ Extract text and tables from PDFs
- ✅ Support for multi-page documents
- ✅ Preserve formatting where possible
- ✅ Free hosting on Render

## Deploy on Render

### Quick Deploy
[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy)

### Manual Deploy
1. Fork this repository
2. Create account on [Render](https://render.com)
3. Create new Background Worker
4. Connect your GitHub repository
5. Add environment variable: `BOT_TOKEN`
6. Deploy!

## Environment Variables
- `BOT_TOKEN`: Your Telegram Bot Token from [@BotFather](https://t.me/botfather)

## Getting Bot Token
1. Open Telegram and search for [@BotFather](https://t.me/botfather)
2. Send `/newbot` command
3. Choose a name for your bot
4. Choose a username (must end with 'bot')
5. Copy the token and add to Render

## Usage
1. Start the bot with `/start`
2. Send any PDF file (max 20MB)
3. Receive converted DOC file
4. Use `/help` for more information

## Commands
- `/start` - Welcome message
- `/help` - Help and instructions
- `/about` - About the bot
- `/stats` - Your conversion statistics

## Limitations
- Maximum file size: 20MB
- Text-based PDFs only (no OCR for scanned documents)
- Images are not preserved
- Complex formatting may be simplified

## Tech Stack
- Python 3.11
- python-telegram-bot
- PyPDF2
- pdfplumber
- python-docx

## Support
For issues or questions, contact: @cybrick (Telegram).

## License
MIT License
