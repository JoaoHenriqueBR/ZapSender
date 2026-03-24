import random
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

# Verifica se o documento existe
if CAMINHO_DOCUMENTO and not os.path.exists(CAMINHO_DOCUMENTO):
    raise FileNotFoundError(f"Documento não encontrado em: {CAMINHO_DOCUMENTO}")

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


def anexar_documento(driver, caminho, timeout=30):
    """
    Abre o menu de anexo do WhatsApp Web e faz o upload do documento.
    Documentos suportam legenda, então o texto pode ser adicionado antes de enviar.
    """
    # Clica no botão de anexo (ícone de clipe)
    attach_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-icon='attach-menu-plus']"))
    )
    attach_btn.click()
    time.sleep(1)

    # Localiza o input de arquivo para documentos
    # O seletor 'input[accept="*"]' corresponde ao input de documentos no WhatsApp Web
    file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    documento_input = None
    for fi in file_inputs:
        accept = fi.get_attribute("accept") or ""
        if "*" in accept or "pdf" in accept or not accept:
            documento_input = fi
            break

    if documento_input is None:
        documento_input = file_inputs[0]  # fallback para o primeiro input encontrado

    documento_input.send_keys(os.path.abspath(caminho))
    print("   -> Documento anexado.")
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

            mensagem1 = f"""{escolha_emoji} {escolha_saudacao}, {nome}! {escolha_emoji}

Insira sua primeira mensagem para enviar"""

            mensagem2 = f"""{escolha_saudacao}, {nome}! Insira sua segunda mensagem para enviar {escolha_emoji3}"""

            mensagem3 = f"""{escolha_emoji} *{escolha_saudacao}, {nome}!* {escolha_emoji}

Insira sua terceira mensagem para enviar {escolha_emoji2}"""

            mensagem4 = f"""*{escolha_saudacao}, {nome}!

Insira sua quarta mensagem para enviar!"""

            msgs = [mensagem1, mensagem2, mensagem3, mensagem4]
            mensagem = random.choice(msgs)

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
            anexar_documento(driver, CAMINHO_DOCUMENTO)

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
