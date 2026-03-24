import json
import os
import subprocess
import sys

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

    arquivo_excel = input("\n  Caminho do arquivo Excel: ").strip()
    caminho_arquivo = input(f"  Caminho do {tipo}: ").strip()

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
