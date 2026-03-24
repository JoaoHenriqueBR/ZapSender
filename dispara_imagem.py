import random
import subprocess
import sys

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
from io import BytesIO
from PIL import Image
import urllib.parse
import platform

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = ''  # Nome do seu arquivo
CAMINHO_IMAGEM = os.path.abspath('')  # Nome da imagem (deve estar na mesma pasta ou caminho completo)
CELULAR = 'CELULAR' # Nome da coluna de Números de Celular na planilha
NOME = 'NOME' # Nome da coluna de nomes na planilha

# Mensagens são escolhidas no Loop principal e podem ser randomizadas por conversa

# Verifica se a imagem existe
if CAMINHO_IMAGEM and not os.path.exists(CAMINHO_IMAGEM):
    raise FileNotFoundError(f"Imagem não encontrada em: {CAMINHO_IMAGEM}")

IS_WINDOWS = platform.system() == "Windows"

def copiar_imagem_para_clipboard(caminho):
    """
    Carrega a imagem e a coloca na área de transferência.
    Suporta Windows (win32clipboard) e Linux (xclip).
    """
    if IS_WINDOWS:
        try:
            import win32clipboard
            image = Image.open(caminho)
            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
        except ImportError:
            print("⚠️ win32clipboard não disponível. Tente instalar pywin32.")
    else:
        # Linux: converte para PNG e usa xclip
        image = Image.open(caminho)
        png_path = "/tmp/_zapsender_clip.png"
        image.convert("RGBA").save(png_path, "PNG")
        try:
            subprocess.run(
                ["xclip", "-selection", "clipboard", "-t", "image/png", "-i", png_path],
                check=True
            )
        except FileNotFoundError:
            print("⚠️ xclip não encontrado. Instale com: sudo apt install xclip")
        except subprocess.CalledProcessError as e:
            print(f"⚠️ Erro ao copiar imagem para clipboard: {e}")


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
        # Linux: usa xclip
        try:
            proc = subprocess.Popen(
                ["xclip", "-selection", "clipboard"],
                stdin=subprocess.PIPE
            )
            proc.communicate(input=texto.encode("utf-8"))
        except FileNotFoundError:
            print("⚠️ xclip não encontrado. Instale com: sudo apt install xclip")


def formatar_numero(numero):
    """
    Limpa e formata o número para o padrão internacional (55 + DDD + Numero). - Por enquanto, suporte apenas a numeros brasileiros.
    Remove espaços, traços e parênteses.
    """
    if pd.isna(numero):
        return None

    # Converte para string e remove caracteres não numéricos
    num_str = ''.join(filter(str.isdigit, str(numero)))

    if not num_str:
        return None

    # Se não tiver o código do país (55), adiciona. (Assumindo Brasil)
    # Ajuste essa lógica se sua base já tiver o 55
    if len(num_str) <= 11:
        num_str = "55" + num_str

    return num_str


def formatar_nome(nome_completo):
    if pd.isna(nome_completo) or str(nome_completo).strip() == "":
        return "Aluno(a)"  # Fallback se não tiver nome

    # Pega o primeiro nome
    primeiro_nome = str(nome_completo).split()[0]

    # Capitaliza (primeira letra maiúscula, resto minúscula)
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

    # Cria uma lista única de números (para não mandar duplicado)
    lista_para_envio = []
    numeros_ja_adicionados = set() # para evitar duplicidade na lista

    for _, row in df.iterrows():
        nome_tratado = formatar_nome(row.get(NOME))
        nums = [formatar_numero(row.get(CELULAR))] # Em caso de variás colunas, só adicionar mais outro item na lista com a função formatar_numero

        for num in nums:
            # Se o número for válido e ainda não estiver na lista
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
    # options.add_argument("--headless") # Não use headless para WhatsApp, pode dar bloqueio
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")

    # Aguarda o usuário logar (procura pelo elemento da lista de chats)
    input("⚠️ Escaneie o QR Code e pressione ENTER aqui no terminal quando o WhatsApp carregar completamente...")

    total_enviados = 0
    total_invalidos = 0
    actions = ActionChains(driver)  # Inicializa o controlador de teclado

    # 3. Loop de Envio
    for i, dados in enumerate(lista_completa):
        telefone = dados['numero']
        nome = dados['nome']

        saudacoes = ["Olá", "Oi", "Tudo bem", "Oie", "ATENÇÃO", "Hey", "Hello", "Hi", "Fala", "Ei", "E aí", "Salve", "Beleza", "Como vai", "Como está"]
        emoji_saudacao = ["🌟", "✨", "🚀", "👋", "🚨", "😁", "😊", "🔥", "💙"]

        escolha_saudacao = random.choice(saudacoes)
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

            # Escolhe uma mensagem de forma aleatória
            msgs = [mensagem1, mensagem2, mensagem3, mensagem4]
            mensagem = random.choice(msgs)

            # 1. Copia a imagem para o clipboard
            print("📋 Copiando imagem para a memória...")
            copiar_imagem_para_clipboard(CAMINHO_IMAGEM)

            # Abre a conversa direta
            link = f"https://web.whatsapp.com/send?phone={telefone}"
            driver.get(link)

            # 2. Aguarda o chat carregar
            try:
                caixa_texto = localizar_caixa_texto(driver)
            except Exception:
                # Se não achou o anexo em 20s, assume que deu erro no carregamento ou número inválido
                print(f"⚠️ Alerta: Não foi possível abrir o chat de {telefone}.")
                numeros_ja_processados.append(telefone)
                total_invalidos += 1
                continue  # Pula para o próximo número imediatamente


            # 3. Clica na caixa e cola (Ctrl + V)
            caixa_texto.click()
            time.sleep(random.uniform(1, 4))

            # Trava de Condicional *opcional*, para evitar o máximo a repetição (spam) de imagens
            teste_aleatorio = random.randint(1, 8)
            print(f"Teste lógico: {teste_aleatorio}")
            if teste_aleatorio != 3:
                caixa_texto.send_keys(Keys.CONTROL, 'v')

                # 4. Aguarda a pré-visualização da imagem aparecer
                print("   -> Imagem colada, aguardando prévia...")
                # Pausa vital: O WhatsApp precisa de 2 a 3 segundos para carregar a prévia da imagem.
                # Se der ENTER antes da prévia aparecer, nada acontece.
                time.sleep(random.uniform(3, 6))
            else:
                print("Imagem não enviada para maior personalização.")

            # Trava de Condicional *opcional*, para evitar o máximo a repetição (spam) de mensagens
            if teste_aleatorio != 1:
                # 5. Carrega o Texto na Memória (Substitui a imagem no clipboard)
                copiar_texto_para_clipboard(mensagem)

                # 6. Cola o Texto na Legenda
                # O foco já está automaticamente na legenda quando a imagem abre
                actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

                print("   -> Legenda colada.")
                time.sleep(random.uniform(2, 6))  # Tempo visual para conferência
            else:
                print("Legenda não enviada para maior personalização.")



            # 7. Pressiona ENTER
            actions.send_keys(Keys.ENTER).perform()

            time.sleep(4)  # Espera confirmar o envio visualmente
            print("   -> ENTER pressionado. Enviado!")

            total_enviados += 1

            numeros_ja_processados.append(telefone)


            # --- LÓGICA DE PAUSA (Constraint do Usuário) ---

            # Verifica se já enviamos 30 e se não é o último da lista
            if total_enviados > 0 and total_enviados % 30 == 0 and i < len(lista_completa) - 1:
                pausa = random.uniform(480, 1200)
                print(f"🛑 Limite de {total_enviados} envios atingido. Pausando por {pausa} segundos para segurança...")
                print(f"Números enviados - {len(numeros_ja_processados)}: {numeros_ja_processados}")
                time.sleep(pausa)
                print("▶️ Retomando envios...")
            # Verifica se já enviamos 10 e se não é o último da lista
            elif total_enviados > 0 and total_enviados % 10 == 0 and i < len(lista_completa) - 1:
                pausa = random.uniform(120, 480)
                print(f"🛑 {total_enviados} envios atingido. Pausando por {pausa} segundos para segurança...")
                print(f"Números enviados - {len(numeros_ja_processados)}: {numeros_ja_processados}")
                time.sleep(pausa)
                print("▶️ Retomando envios...")
            else:
                if total_invalidos > 0 and total_invalidos % 3 == 0 and i < len(lista_completa) - 1:
                    pausa = random.uniform(60, 300)
                    print(f"🛑 {total_invalidos} números inválidos atingido. Pausando por {pausa} segundos para segurança...")
                    time.sleep(pausa)
                    print("▶️ Retomando envios...")
                    total_invalidos = 0
                else:
                    pausa = random.uniform(15, 30)
                    print(f'Esperando {pausa} segundos para o próximo envio.')
                    time.sleep(pausa)


        except Exception as e:
            print(f"❌ Erro ao enviar para {telefone}: {e}")
            # Em caso de erro, continua para o próximo número

    print("🏁 Processo finalizado.")
    driver.quit()


if __name__ == "__main__":
    main()
