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
import platform
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
    "chrome": ["google-chrome", "google-chrome-stable", "chrome", "chrome.exe"],
    "chromium": ["chromium", "chromium-browser", "chromium.exe"],
    "brave": ["brave-browser", "brave", "brave.exe"],
}

WINDOWS_BROWSER_PATHS = {
    "chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe",
    ],
    "chromium": [
        r"C:\Program Files\Chromium\Application\chrome.exe",
        r"C:\Program Files (x86)\Chromium\Application\chrome.exe",
        r"%LOCALAPPDATA%\Chromium\Application\chrome.exe",
    ],
    "brave": [
        r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
        r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\Application\brave.exe",
    ],
}


def normalizar_browser(browser):
    browser_normalizado = (browser or "chrome").strip().lower()
    return BROWSER_ALIASES.get(browser_normalizado, browser_normalizado)


def _encontrar_binario_no_path(browser_normalizado):
    for binary_name in BROWSER_BINARIES.get(browser_normalizado, []):
        binary_path = shutil.which(binary_name)
        if binary_path:
            return binary_path

    return None


def _expandir_caminho_windows(path):
    localappdata = os.environ.get("LOCALAPPDATA", "")
    return path.replace("%LOCALAPPDATA%", localappdata)


def _encontrar_binario_windows(browser_normalizado):
    for candidate in WINDOWS_BROWSER_PATHS.get(browser_normalizado, []):
        expanded_candidate = _expandir_caminho_windows(candidate)
        if expanded_candidate and os.path.exists(expanded_candidate):
            return expanded_candidate

    return None


def encontrar_binario_browser(browser):
    browser_normalizado = normalizar_browser(browser)

    binary_path = _encontrar_binario_no_path(browser_normalizado)
    if binary_path:
        return binary_path

    if platform.system() == "Windows":
        return _encontrar_binario_windows(browser_normalizado)

    return None


def _resolver_caminho_informado(browser_binary):
    binary_path = browser_binary if os.path.isabs(browser_binary) else os.path.abspath(browser_binary)
    if not os.path.exists(binary_path):
        raise FileNotFoundError(f"Binário do navegador não encontrado: {binary_path}")
    return binary_path


def criar_driver(browser="chrome", browser_binary=None):
    browser_normalizado = normalizar_browser(browser)

    if browser_normalizado not in SUPPORTED_BROWSERS:
        suportados = ", ".join(sorted(SUPPORTED_BROWSERS))
        raise ValueError(
            f"Navegador '{browser}' não suportado. Opções disponíveis: {suportados}."
        )

    options = webdriver.ChromeOptions()

    if browser_binary:
        options.binary_location = _resolver_caminho_informado(browser_binary)
    else:
        binary_path = encontrar_binario_browser(browser_normalizado)
        if binary_path:
            options.binary_location = binary_path

    return webdriver.Chrome(options=options)
