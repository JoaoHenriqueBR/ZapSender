import subprocess
import sys
import os


def exibir_menu():
    print()
    print("=" * 45)
    print("         ZapSender - Menu Principal")
    print("=" * 45)
    print("  [1] Imagens   + Mensagens")
    print("  [2] Audios    + Mensagens")
    print("  [3] Documentos + Mensagens")
    print("  [0] Sair")
    print("=" * 45)
    return input("  Escolha uma opcao: ").strip()


def main():
    scripts = {
        "1": ("dispara_imagem.py",    "Imagens + Mensagens"),
        "2": ("dispara_audio.py",     "Audios + Mensagens"),
        "3": ("dispara_documentos.py","Documentos + Mensagens"),
    }

    while True:
        opcao = exibir_menu()

        if opcao == "0":
            print("\nSaindo do ZapSender. Ate mais!")
            break
        elif opcao in scripts:
            arquivo, descricao = scripts[opcao]
            caminho = os.path.join(os.path.dirname(__file__), arquivo)

            if not os.path.exists(caminho):
                print(f"\n❌ Arquivo '{arquivo}' nao encontrado.")
                continue

            print(f"\n▶️  Iniciando: {descricao}...\n")
            subprocess.run([sys.executable, caminho])
            print(f"\n✅ '{descricao}' finalizado. Voltando ao menu...")
        else:
            print("\n⚠️  Opcao invalida. Tente novamente.")


if __name__ == "__main__":
    main()
