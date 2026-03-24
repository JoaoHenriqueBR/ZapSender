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

# Constante do Windows para arrastar e soltar arquivos via clipboard
CF_HDROP = 15

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = ''  # Nome do seu arquivo Excel
CAMINHO_AUDIO = os.path.abspath('')  # Caminho do arquivo de audio (.mp3, .ogg, .aac, .m4a, .wav)
CELULAR = 'CELULAR'  # Nome da coluna de Números de Celular na planilha
NOME = 'NOME'        # Nome da coluna de nomes na planilha

# Mensagens são escolhidas no Loop principal e podem ser randomizadas por conversa

IS_WINDOWS = platform.system() == "Windows"


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
    print("   -> Arquivo de audio copiado para o clipboard (CF_HDROP).")


def anexar_audio_linux(driver, caminho, timeout=30):
    """
    Fallback para Linux: abre o menu de anexo via Selenium e faz o upload do arquivo.
    """
    attach_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-icon='attach-menu-plus']"))
    )
    attach_btn.click()
    time.sleep(1)

    file_input = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
    )
    file_input.send_keys(os.path.abspath(caminho))
    print("   -> Arquivo de audio anexado via Selenium.")
    time.sleep(random.uniform(2, 4))


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
    global ARQUIVO_EXCEL, CAMINHO_AUDIO, CELULAR, NOME

    # Lê configurações centralizadas do main.py (se disponíveis)
    _config = json.loads(os.environ.get("ZAPSENDER_CONFIG", "{}"))
    if _config:
        ARQUIVO_EXCEL = _config.get("ARQUIVO_EXCEL", ARQUIVO_EXCEL)
        CAMINHO_AUDIO = _config.get("CAMINHO_ARQUIVO", CAMINHO_AUDIO)
        CELULAR = _config.get("COLUNA_CELULAR", CELULAR)
        NOME = _config.get("COLUNA_NOME", NOME)
    _mensagens_custom = _config.get("MENSAGENS", [])
    _teste_aleatorio_ativo = _config.get("TESTE_ALEATORIO", True)

    # Valida o arquivo de mídia
    if CAMINHO_AUDIO and not os.path.exists(CAMINHO_AUDIO):
        print(f"❌ Audio não encontrado: {CAMINHO_AUDIO}")
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

    # Determina o arquivo de rastreamento baseado no nome do Excel
    nome_base = os.path.splitext(os.path.basename(ARQUIVO_EXCEL))[0]
    arquivo_processados = f"{nome_base}_processados.txt"

    # Carrega números já processados do arquivo (se existir)
    numeros_ja_processados = []
    if os.path.exists(arquivo_processados):
        with open(arquivo_processados, "r", encoding="utf-8") as f:
            numeros_ja_processados = [linha.strip() for linha in f if linha.strip()]
        print(f"📋 {len(numeros_ja_processados)} números já processados carregados de '{arquivo_processados}'.")

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
                with open(arquivo_processados, "a", encoding="utf-8") as f:
                    f.write(telefone + "\n")
                total_invalidos += 1
                continue

            # 2. Anexa o arquivo de audio
            print("🎵 Anexando arquivo de audio...")
            if IS_WINDOWS:
                # Windows: copia o arquivo para o clipboard via CF_HDROP e cola na caixa de texto
                copiar_arquivo_para_clipboard(CAMINHO_AUDIO)
                caixa_texto = localizar_caixa_texto(driver)
                caixa_texto.click()
                time.sleep(random.uniform(1, 2))
                caixa_texto.send_keys(Keys.CONTROL, 'v')
                # Aguarda o WhatsApp carregar a prévia do arquivo
                time.sleep(random.uniform(3, 5))
            else:
                # Linux: usa Selenium para interagir com o input de arquivo
                anexar_audio_linux(driver, CAMINHO_AUDIO)

            # 3. Confirma o envio do audio (ENTER na tela de preview)
            # WhatsApp Web exibe uma preview do audio antes de enviar
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(random.uniform(2, 4))
            print("   -> Audio enviado!")

            # 4. Trava condicional para evitar repetição de mensagens
            if _teste_aleatorio_ativo:
                teste_aleatorio = random.randint(1, 8)
                print(f"Teste lógico: {teste_aleatorio}")
            else:
                teste_aleatorio = 2  # valor neutro: garante envio completo
            if teste_aleatorio != 1:
                # 5. Envia a mensagem de texto separadamente (audio não suporta legenda)
                caixa_texto = localizar_caixa_texto(driver)
                caixa_texto.click()
                time.sleep(random.uniform(1, 2))

                copiar_texto_para_clipboard(mensagem)
                actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                print("   -> Mensagem colada.")
                time.sleep(random.uniform(2, 4))

                actions.send_keys(Keys.ENTER).perform()
                print("   -> Mensagem enviada!")
            else:
                print("Mensagem de texto não enviada para maior personalização.")

            time.sleep(2)
            total_enviados += 1
            numeros_ja_processados.append(telefone)
            with open(arquivo_processados, "a", encoding="utf-8") as f:
                f.write(telefone + "\n")

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
