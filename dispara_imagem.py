import random

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
import win32clipboard # Do pacote pywin32
from PIL import Image # Do pacote Pillow
import urllib.parse

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = 'Relatório de Alunos (3).xlsx'  # Nome do seu arquivo
CAMINHO_IMAGEM = os.path.abspath('indica_amigo.jpg')  # Nome da imagem (deve estar na mesma pasta ou caminho completo)
COLUNA_ALUNO = 'Celular do Aluno'
COLUNA_RESPONSAVEL = 'Celular do Responsável Financeiro'

MENSAGEM_LEGENDA = """Olá! Esperamos que esteja tudo bem com você e sua família. ✨

Temos uma novidade incrível para você que faz parte da família *CNA Interlagos* !

Que tal estudar com seus amigos e ainda garantir prêmios? Com a nossa nova campanha de indicação, *todo mundo sai ganhando* :

🎁 *Para VOCÊ:* Um Voucher de *R$ 50,00* na Kalunga por cada amigo que se matricular! 
🎁 *Para seu AMIGO:* Um *SUPER DESCONTO* especial + um minicurso gratuito à escolha (Inglês, Espanhol ou Criação de Jogos)!

Quanto mais amigos você indicar, mais vouchers você acumula. 💸

*Para participar é muito simples:* basta enviar nosso número para seus amigos e preencher os dados deles aqui: 👉 https://forms.gle/4JdJPhLmhipjqJVZ8

Bora trazer a sua galera para o CNA? Se tiver qualquer dúvida, é só me chamar por aqui! 📲"""

# Verifica se a imagem existe
if not os.path.exists(CAMINHO_IMAGEM):
    raise FileNotFoundError(f"Imagem não encontrada em: {CAMINHO_IMAGEM}")

def copiar_imagem_para_clipboard(caminho):
    """
    Carrega a imagem e a coloca na área de transferência do Windows (Ctrl+C manual)
    """
    image = Image.open(caminho)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def copiar_texto_para_clipboard(texto):
    """Copia o texto (com emojis) para o Clipboard (Memória)"""
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    # CF_UNICODETEXT é essencial para emojis e acentos funcionarem
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, texto)
    win32clipboard.CloseClipboard()

def formatar_numero(numero):
    """
    Limpa e formata o número para o padrão internacional (55 + DDD + Numero).
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

def localizar_caixa_texto(driver, timeout=30):
    """Localiza a caixa onde digita o texto (muito mais fácil de achar que o clipe)"""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "footer div[contenteditable='true']"))
    )


def main():
    # 1. Carregar e Tratar Dados
    print("📂 Lendo arquivo Excel...")
    df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)

    # Cria uma lista única de números (para não mandar duplicado se aluno e resp forem o mesmo)
    lista_completa = set()

    for index, row in df.iterrows():
        num_aluno = formatar_numero(row.get(COLUNA_ALUNO))
        num_resp = formatar_numero(row.get(COLUNA_RESPONSAVEL))

        if num_aluno: lista_completa.add(num_aluno)
        if num_resp: lista_completa.add(num_resp)

    lista_completa = list(lista_completa)
    print(f"✅ Total de números únicos para envio: {len(lista_completa)}")

    numeros_ja_processados = []

    lista_envio = [num for num in lista_completa if num not in numeros_ja_processados]

    print(f"✅ Números já processados: {len(numeros_ja_processados)}")
    print(f"🚀 Restam para enviar: {len(lista_envio)}")

    if len(lista_envio) == 0:
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
    actions = ActionChains(driver)  # Inicializa o controlador de teclado

    # 3. Loop de Envio
    for i, telefone in enumerate(lista_envio):
        try:
            print(f"[{i + 1}/{len(lista_envio)}] Enviando para {telefone}...")

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
                continue  # Pula para o próximo número imediatamente

            # 3. Clica na caixa e cola (Ctrl + V)
            caixa_texto.click()
            time.sleep(1)
            caixa_texto.send_keys(Keys.CONTROL, 'v')

            # 4. Aguarda a pré-visualização da imagem aparecer
            print("   -> Imagem colada, aguardando prévia...")
            # Pausa vital: O WhatsApp precisa de 2 a 3 segundos para carregar a prévia da imagem.
            # Se der ENTER antes da prévia aparecer, nada acontece.
            time.sleep(3)

            # 5. Carrega o Texto na Memória (Substitui a imagem no clipboard)
            copiar_texto_para_clipboard(MENSAGEM_LEGENDA)

            # 6. Cola o Texto na Legenda
            # O foco já está automaticamente na legenda quando a imagem abre
            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

            print("   -> Legenda colada.")
            time.sleep(1.5)  # Tempo visual para conferência

            # 7. Pressiona ENTER
            actions.send_keys(Keys.ENTER).perform()

            time.sleep(2)  # Espera confirmar o envio visualmente
            print("   -> ENTER pressionado. Enviado!")

            total_enviados += 1

            # --- LÓGICA DE PAUSA (Constraint do Usuário) ---
            time.sleep(random.uniform(5,9))

            # Pausa de 10 minutos a cada 60 envios
            # Verifica se já enviamos 60 e se não é o último da lista
            if total_enviados > 0 and total_enviados % 60 == 0 and i < len(lista_envio) - 1:
                print(f"🛑 Limite de 60 envios atingido. Pausando por 10 minutos para segurança...")
                time.sleep(600)  # 600 segundos = 10 minutos
                print("▶️ Retomando envios...")

        except Exception as e:
            print(f"❌ Erro ao enviar para {telefone}: {e}")
            # Em caso de erro, continua para o próximo número

    print("🏁 Processo finalizado.")
    driver.quit()


if __name__ == "__main__":
    main()