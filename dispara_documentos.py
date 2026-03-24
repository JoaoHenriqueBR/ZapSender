import json
import random
import struct
import subprocess

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import platform

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = ''  # Nome do seu arquivo Excel
CAMINHO_DOCUMENTO = os.path.abspath('')  # Caminho do documento (.pdf, .docx, .xlsx, .pptx, etc.)
CELULAR = 'CELULAR'  # Nome da coluna de Números de Celular na planilha
NOME = 'NOME'        # Nome da coluna de nomes na planilha

# Mensagens são escolhidas no Loop principal e podem ser randomizadas por conversa

IS_WINDOWS = platform.system() == "Windows"

# Constante do Windows para arrastar e soltar arquivos via clipboard
CF_HDROP = 15


def copiar_texto_para_clipboard(texto):
    """Copia o texto (com emojis) para o Clipboard (Memória)"""
    if IS_WINDOWS:
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, texto)
            win32clipboard.CloseClipboard()
        except ImportError:
            print("⚠️ win32clipboard não disponível. Tente instalar pywin32.")
    else:
        try:
            proc = subprocess.Popen(
                ["xclip", "-selection", "clipboard"],
                stdin=subprocess.PIPE
            )
            proc.communicate(input=texto.encode("utf-8"))
        except FileNotFoundError:
            print("⚠️ xclip não encontrado. Instale com: sudo apt install xclip")


def copiar_arquivo_para_clipboard(caminho_arquivo):
    """
    Coloca a referência do arquivo na área de transferência no Windows.
    Simula um Ctrl+C feito no Explorer usando a estrutura DROPFILES (CF_HDROP).
    O WhatsApp Web reconhece esse formato e trata como upload de arquivo.
    Funciona para qualquer tipo de arquivo: PDF, DOCX, XLSX, etc.
    """
    if not os.path.exists(caminho_arquivo):
        print(f"❌ Arquivo não encontrado: {caminho_arquivo}")
        return

    try:
        import win32clipboard
    except ImportError:
        print("⚠️ win32clipboard não disponível. Tente instalar pywin32.")
        return

    caminho_abs = os.path.abspath(caminho_arquivo)

    # Monta a estrutura DROPFILES:
    # - offset=20: posição onde começa a lista de arquivos
    # - fWide=1: usa Unicode (UTF-16)
    offset = 20
    dropfiles_header = struct.pack("IIIII", offset, 0, 0, 0, 1)

    # O caminho termina com null duplo (exigência do formato)
    files_data = caminho_abs.encode("utf-16-le") + b"\0\0"

    data = dropfiles_header + files_data

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(CF_HDROP, data)
    win32clipboard.CloseClipboard()
    print("   -> Documento copiado para o clipboard (CF_HDROP).")


def anexar_documento_linux(driver, caminho, timeout=30):
    """
    Fallback para Linux: abre o menu de anexo via Selenium e faz o upload do documento.
    """
    attach_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-icon='attach-menu-plus']"))
    )
    attach_btn.click()
    time.sleep(1)

    file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    documento_input = None
    for fi in file_inputs:
        accept = fi.get_attribute("accept") or ""
        if "*" in accept or "pdf" in accept or not accept:
            documento_input = fi
            break

    if documento_input is None:
        documento_input = file_inputs[0]

    documento_input.send_keys(os.path.abspath(caminho))
    print("   -> Documento anexado via Selenium.")
    time.sleep(random.uniform(2, 4))


def localizar_legenda_documento(driver, timeout=20):
    """Localiza a caixa de legenda que aparece após anexar um documento."""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[data-tab='10'][contenteditable='true']"))
    )


def formatar_numero(numero):
    """
    Limpa e formata o número para o padrão internacional (55 + DDD + Numero).
    Remove espaços, traços e parênteses.
    """
    if pd.isna(numero):
        return None

    num_str = ''.join(filter(str.isdigit, str(numero)))

    if not num_str:
        return None

    if len(num_str) <= 11:
        num_str = "55" + num_str

    return num_str


def formatar_nome(nome_completo):
    if pd.isna(nome_completo) or str(nome_completo).strip() == "":
        return "Aluno(a)"

    primeiro_nome = str(nome_completo).split()[0]
    return primeiro_nome.capitalize()


def localizar_caixa_texto(driver, timeout=30):
    """Localiza a caixa onde digita o texto"""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "footer div[contenteditable='true']"))
    )


def main():
    global ARQUIVO_EXCEL, CAMINHO_DOCUMENTO, CELULAR, NOME

    # Lê configurações centralizadas do main.py (se disponíveis)
    _config = json.loads(os.environ.get("ZAPSENDER_CONFIG", "{}"))
    if _config:
        ARQUIVO_EXCEL = _config.get("ARQUIVO_EXCEL", ARQUIVO_EXCEL)
        CAMINHO_DOCUMENTO = _config.get("CAMINHO_ARQUIVO", CAMINHO_DOCUMENTO)
        CELULAR = _config.get("COLUNA_CELULAR", CELULAR)
        NOME = _config.get("COLUNA_NOME", NOME)
    _mensagens_custom = _config.get("MENSAGENS", [])

    # Valida o arquivo de mídia
    if CAMINHO_DOCUMENTO and not os.path.exists(CAMINHO_DOCUMENTO):
        print(f"❌ Documento não encontrado: {CAMINHO_DOCUMENTO}")
        return

    # 1. Carregar e Tratar Dados
    print("📂 Lendo arquivo Excel...")
    df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)

    lista_para_envio = []
    numeros_ja_adicionados = set()

    for _, row in df.iterrows():
        nome_tratado = formatar_nome(row.get(NOME))
        nums = [formatar_numero(row.get(CELULAR))]

        for num in nums:
            if num and num not in numeros_ja_adicionados:
                lista_para_envio.append({
                    'numero': num,
                    'nome': nome_tratado,
                })
                numeros_ja_adicionados.add(num)

    numeros_ja_processados = []

    lista_completa = [item for item in lista_para_envio if item['numero'] not in numeros_ja_processados]
    print(f"✅ Total de números únicos para envio: {len(lista_para_envio)}")
    print(f"✅ Números já processados: {len(numeros_ja_processados)}")
    print(f"✅ Para enviar: {len(lista_completa)}")

    if len(lista_completa) == 0:
        print("🎉 Todos os números já foram processados! Nada a fazer.")
        return

    # 2. Iniciar o Navegador
    print("🚀 Iniciando navegador. Por favor, escaneie o QR Code do WhatsApp Web.")
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")

    input("⚠️ Escaneie o QR Code e pressione ENTER aqui no terminal quando o WhatsApp carregar completamente...")

    total_enviados = 0
    total_invalidos = 0
    actions = ActionChains(driver)

    # 3. Loop de Envio
    for i, dados in enumerate(lista_completa):
        telefone = dados['numero']
        nome = dados['nome']

        saudacoes_inicio = ["Olá", "Oi", "Oie", "ATENÇÃO", "Hey", "Hello", "Hi", "Fala", "Ei", "E aí", "Salve", "Beleza"]
        saudacoes_pergunta = ["Tudo bem", "Como vai", "Como está"]
        emoji_saudacao = ["🌟", "✨", "🚀", "👋", "🚨", "😁", "😊", "🔥", "💙"]

        escolha_saudacao = random.choice(saudacoes_inicio)
        escolha_pergunta = random.choice(saudacoes_pergunta)
        escolha_emoji = random.choice(emoji_saudacao)
        escolha_emoji2 = random.choice(emoji_saudacao)
        escolha_emoji3 = random.choice(emoji_saudacao)

        try:
            print(f"[{i + 1}/{len(lista_completa)}] Enviando para {nome} ({telefone})...")

            if _mensagens_custom:
                template = random.choice(_mensagens_custom)
                try:
                    mensagem = template.format(
                        nome=nome,
                        escolha_saudacao=escolha_saudacao,
                        escolha_pergunta=escolha_pergunta,
                        escolha_emoji=escolha_emoji,
                        escolha_emoji2=escolha_emoji2,
                        escolha_emoji3=escolha_emoji3,
                    )
                except (KeyError, IndexError):
                    mensagem = template
            else:
                mensagem1 = f"""{escolha_emoji} {escolha_saudacao}, {nome}! {escolha_emoji}

Insira sua primeira mensagem para enviar"""

                mensagem2 = f"""{escolha_saudacao}, {nome}! Insira sua segunda mensagem para enviar {escolha_emoji3}"""

                mensagem3 = f"""{escolha_emoji} *{escolha_saudacao}, {nome}!* {escolha_emoji}

Insira sua terceira mensagem para enviar {escolha_emoji2}"""

                mensagem4 = f"""*{escolha_saudacao}, {nome}!

Insira sua quarta mensagem para enviar!"""

                mensagem = random.choice([mensagem1, mensagem2, mensagem3, mensagem4])

            # Abre a conversa direta
            link = f"https://web.whatsapp.com/send?phone={telefone}"
            driver.get(link)

            # 1. Aguarda o chat carregar
            try:
                localizar_caixa_texto(driver)
            except Exception:
                print(f"⚠️ Alerta: Não foi possível abrir o chat de {telefone}.")
                numeros_ja_processados.append(telefone)
                total_invalidos += 1
                continue

            # 2. Anexa o documento
            print("📄 Anexando documento...")
            if IS_WINDOWS:
                # Windows: copia o arquivo para o clipboard via CF_HDROP e cola na caixa de texto
                copiar_arquivo_para_clipboard(CAMINHO_DOCUMENTO)
                caixa_texto = localizar_caixa_texto(driver)
                caixa_texto.click()
                time.sleep(random.uniform(1, 2))
                caixa_texto.send_keys(Keys.CONTROL, 'v')
                # Aguarda o WhatsApp carregar a prévia do documento
                time.sleep(random.uniform(3, 5))
            else:
                # Linux: usa Selenium para interagir com o input de arquivo
                anexar_documento_linux(driver, CAMINHO_DOCUMENTO)

            # 3. Trava condicional para evitar repetição de mensagens
            teste_aleatorio = random.randint(1, 8)
            print(f"Teste lógico: {teste_aleatorio}")
            if teste_aleatorio != 1:
                # 4. Adiciona legenda ao documento (documentos suportam legenda no WhatsApp)
                try:
                    legenda = localizar_legenda_documento(driver, timeout=10)
                    legenda.click()
                except Exception:
                    # Se não encontrar a caixa de legenda, tenta a caixa de texto principal
                    legenda = localizar_caixa_texto(driver)
                    legenda.click()

                copiar_texto_para_clipboard(mensagem)
                actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                print("   -> Legenda colada.")
                time.sleep(random.uniform(2, 4))
            else:
                print("Legenda não adicionada para maior personalização.")

            # 5. Envia o documento (com ou sem legenda)
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(4)
            print("   -> Documento enviado!")

            total_enviados += 1
            numeros_ja_processados.append(telefone)

            # --- LÓGICA DE PAUSA ---
            if total_enviados > 0 and total_enviados % 30 == 0 and i < len(lista_completa) - 1:
                pausa = random.uniform(480, 1200)
                print(f"🛑 Limite de {total_enviados} envios atingido. Pausando por {pausa:.0f}s para segurança...")
                print(f"Números enviados - {len(numeros_ja_processados)}: {numeros_ja_processados}")
                time.sleep(pausa)
                print("▶️ Retomando envios...")
            elif total_enviados > 0 and total_enviados % 10 == 0 and i < len(lista_completa) - 1:
                pausa = random.uniform(120, 480)
                print(f"🛑 {total_enviados} envios atingido. Pausando por {pausa:.0f}s para segurança...")
                print(f"Números enviados - {len(numeros_ja_processados)}: {numeros_ja_processados}")
                time.sleep(pausa)
                print("▶️ Retomando envios...")
            else:
                if total_invalidos > 0 and total_invalidos % 3 == 0 and i < len(lista_completa) - 1:
                    pausa = random.uniform(60, 300)
                    print(f"🛑 {total_invalidos} números inválidos atingido. Pausando por {pausa:.0f}s...")
                    time.sleep(pausa)
                    print("▶️ Retomando envios...")
                    total_invalidos = 0
                else:
                    pausa = random.uniform(15, 30)
                    print(f'Esperando {pausa:.0f}s para o próximo envio.')
                    time.sleep(pausa)

        except Exception as e:
            print(f"❌ Erro ao enviar para {telefone}: {e}")

    print("🏁 Processo finalizado.")
    driver.quit()


if __name__ == "__main__":
    main()
