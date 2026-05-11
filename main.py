"""
ZapSender - Message and Media Sender Menu Script

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


import json
import os
import subprocess
import sys

from configura_browser import encontrar_binario_browser, normalizar_browser

TIPOS_ARQUIVO = {
    "1": ("dispara_imagem.py",     "Imagens + Mensagens",     "imagem"),
    "2": ("dispara_audio.py",      "Audios + Mensagens",      "audio"),
    "3": ("dispara_documentos.py", "Documentos + Mensagens",  "documento"),
}

PLACEHOLDERS = (
    "  Variaveis disponíveis: {nome}, {escolha_saudacao}, {escolha_pergunta},\n"
    "                         {escolha_emoji}, {escolha_emoji2}, {escolha_emoji3}"
)


def exibir_menu():
    print()
    print("=" * 45)
    print("         ZapSender - Menu Principal")
    print("=" * 45)
    print("  [1] Imagens    + Mensagens")
    print("  [2] Audios     + Mensagens")
    print("  [3] Documentos + Mensagens")
    print("  [0] Sair")
    print("=" * 45)
    return input("  Escolha uma opcao: ").strip()


def coletar_mensagem(numero):
    print(f"\n  --- Mensagem {numero} ---")
    print(PLACEHOLDERS)
    print("  (Digite FIM em uma linha nova para finalizar)")
    linhas = []
    while True:
        linha = input("  ")
        if linha.strip().upper() == "FIM":
            break
        linhas.append(linha)
    return "\n".join(linhas)


def coletar_configuracoes(tipo):
    print()
    print("=" * 45)
    print("         Configuracoes de Envio")
    print("=" * 45)

    while True:
        arquivo_excel = input("\n  Caminho do arquivo Excel: ").strip()
        if not arquivo_excel:
            print("  ❌ Caminho do arquivo Excel nao informado.")
            continue
        if not os.path.isfile(arquivo_excel):
            print(f"  ❌ Arquivo Excel nao encontrado: {arquivo_excel}")
            continue
        break

    while True:
        caminho_arquivo = input(f"  Caminho do {tipo}: ").strip()
        if not caminho_arquivo:
            print(f"  ❌ Caminho do {tipo} nao informado.")
            continue
        if not os.path.isfile(caminho_arquivo):
            print(f"  ❌ Arquivo do tipo '{tipo}' nao encontrado: {caminho_arquivo}")
            continue
        break

    coluna_celular = input("  Nome da coluna de celular [padrão: CELULAR]: ").strip() or "CELULAR"
    coluna_nome = input("  Nome da coluna de nomes   [padrão: NOME]: ").strip() or "NOME"
    print("\nConfigurações concluídas com sucesso...")
    print("="*45)
    print("\nNavegador a ser usado: ")
    print("- chrome [Escolhe o Google Chrome]")
    print("- chromium [Escolhe o Chromium]")
    print("- brave [Escolhe o Brave]\n")
    browser = normalizar_browser(
        input("Escolha o seu (padrao: chrome): ").strip() or "chrome"
    )

    browser_binary = ""
    browser_detectado = encontrar_binario_browser(browser)
    if browser_detectado:
        print(f"  ✅ Navegador localizado automaticamente: {browser_detectado}")
    else:
        print(f"  ⚠️  Nao foi possivel localizar automaticamente o navegador '{browser}'.")
        print("  ℹ️  Se estiver no Windows, informe o caminho completo do executavel (.exe).")
        browser_binary = input("  Informe o caminho do binario do navegador: ").strip()

    resp = input("\n  Ativar testes aleatorios na estrutura do envio? [s/N]: ").strip().lower()
    teste_aleatorio_ativo = resp in ("s", "sim", "y", "yes")
    if teste_aleatorio_ativo:
        print("  ⚠️  Testes ativos: arquivo ou mensagem podem ser omitidos aleatoriamente.")
    else:
        print("  ✅  Testes desativados: arquivo e mensagem sempre enviados.")

    while True:
        try:
            qtd = int(input("\n  Quantas mensagens deseja configurar? (minimo 1): ").strip())
            if qtd >= 1:
                break
            print("  ⚠️  Configure pelo menos 1 mensagem.")
        except ValueError:
            print("  ⚠️  Entrada invalida. Digite um numero inteiro.")

    print(f"  Uma das {qtd} sera escolhida aleatoriamente a cada envio.")

    mensagens = [coletar_mensagem(i) for i in range(1, qtd + 1)]

    return {
        "ARQUIVO_EXCEL": arquivo_excel,
        "CAMINHO_ARQUIVO": caminho_arquivo,
        "COLUNA_CELULAR": coluna_celular,
        "COLUNA_NOME": coluna_nome,
        "BROWSER": browser,
        "BROWSER_BINARY": browser_binary,
        "TESTE_ALEATORIO": teste_aleatorio_ativo,
        "MENSAGENS": mensagens,
    }


def main():
    while True:
        opcao = exibir_menu()

        if opcao == "0":
            print("\nSaindo do ZapSender. Ate mais!")
            break
        elif opcao in TIPOS_ARQUIVO:
            arquivo, descricao, tipo = TIPOS_ARQUIVO[opcao]
            caminho_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), arquivo)

            if not os.path.exists(caminho_script):
                print(f"\n❌ Arquivo '{arquivo}' nao encontrado.")
                continue

            config = coletar_configuracoes(tipo)

            env = os.environ.copy()
            env["ZAPSENDER_CONFIG"] = json.dumps(config, ensure_ascii=False)

            print(f"\n▶️  Iniciando: {descricao}...\n")
            subprocess.run([sys.executable, caminho_script], env=env)
            print(f"\n✅ '{descricao}' finalizado. Voltando ao menu...")
        else:
            print("\n⚠️  Opcao invalida. Tente novamente.")


if __name__ == "__main__":
    main()
