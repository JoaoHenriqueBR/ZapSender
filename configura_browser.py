"""
ZapSender - Browser configuration script

Copyright (C) 2026  João Henrique Alves Ferreira <joaohenrique.jh103@protonmail.com>

DISCLAIMER: This software is not affiliated, associated, authorized,
endorsed by, or in any way officially connected with WhatsApp or Meta Platforms, Inc.
The official WhatsApp website can be found at https://www.whatsapp.com.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import os
import shutil

from selenium import webdriver


SUPPORTED_BROWSERS = {
    "chrome",
    "chromium",
    "brave",
}

BROWSER_ALIASES = {
    "google-chrome": "chrome",
    "google chrome": "chrome",
}

BROWSER_BINARIES = {
    "chrome": ["google-chrome", "google-chrome-stable", "chrome"],
    "chromium": ["chromium", "chromium-browser"],
    "brave": ["brave-browser", "brave"],
}


def normalizar_browser(browser):
    browser_normalizado = (browser or "chrome").strip().lower()
    return BROWSER_ALIASES.get(browser_normalizado, browser_normalizado)


def encontrar_binario_browser(browser):
    browser_normalizado = normalizar_browser(browser)

    for binary_name in BROWSER_BINARIES.get(browser_normalizado, []):
        binary_path = shutil.which(binary_name)
        if binary_path:
            return binary_path

    return None


def criar_driver(browser="chrome", browser_binary=None):
    browser_normalizado = normalizar_browser(browser)

    if browser_normalizado not in SUPPORTED_BROWSERS:
        suportados = ", ".join(sorted(SUPPORTED_BROWSERS))
        raise ValueError(
            f"Navegador '{browser}' não suportado. Opções disponíveis: {suportados}."
        )

    options = webdriver.ChromeOptions()

    if browser_binary:
        binary_path = os.path.abspath(browser_binary)
        if not os.path.exists(binary_path):
            raise FileNotFoundError(f"Binário do navegador não encontrado: {binary_path}")
        options.binary_location = binary_path
    else:
        binary_path = encontrar_binario_browser(browser_normalizado)
        if binary_path:
            options.binary_location = binary_path

    return webdriver.Chrome(options=options)
