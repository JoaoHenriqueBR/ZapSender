# ZapSender

A Python CLI tool for automated WhatsApp mass messaging via WhatsApp Web using Selenium.

## Overview

ZapSender reads contact data from an Excel spreadsheet and sends personalized messages (with optional images) to each contact via WhatsApp Web automation.

## Project Structure

- `dispara_imagem.py` — Main script: sends images + text messages via WhatsApp Web
- `requirements.txt` — Python dependencies
- `README.MD` — Project documentation (Portuguese)

## Setup & Configuration

Before running, edit `dispara_imagem.py` and set:
- `ARQUIVO_EXCEL` — Path to your Excel file with contacts
- `CAMINHO_IMAGEM` — Path to the image to send
- `CELULAR` — Column name for phone numbers in the spreadsheet
- `NOME` — Column name for contact names in the spreadsheet

## Running

The workflow runs `python dispara_imagem.py` as a console app.

**Requirements:**
- A Chrome/Chromium browser must be available for Selenium
- On Linux, `xclip` is needed for clipboard image support (`sudo apt install xclip`)
- The script requires WhatsApp Web login via QR code scan in the browser

## Platform Notes

- Originally designed for Windows (used `pywin32` for clipboard)
- Adapted for cross-platform: uses `xclip` on Linux, `win32clipboard` on Windows
- The script is interactive: it opens a browser, waits for QR code scan, then processes the contact list

## Dependencies

- `pandas` + `openpyxl` — Excel file reading
- `pillow` — Image processing
- `selenium` + `webdriver-manager` — Browser automation
- `python-dotenv` — Environment variable support
- `requests` — HTTP requests
