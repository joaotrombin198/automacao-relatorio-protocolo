import subprocess
import time
import sys
import traceback

max_tentativas = None  # ou número inteiro, ex: 5; None = ilimitado
tentativa_atual = 0
timeout_segundos = 600  # 10 minutos

while True:
    tentativa_atual += 1
    print(f"\n[INFO] Iniciando execução do gerar-relatorio.py - tentativa {tentativa_atual}...\n")

    try:
        # Executa o script gerar-relatorio.py num subprocesso com timeout
        resultado = subprocess.run(
            [sys.executable, "gerar-relatorio.py"],
            check=True,
            timeout=timeout_segundos
        )

        # Se chegou aqui, rodou sem erros
        print("\n[SUCESSO] gerar-relatorio.py finalizado com sucesso!\n")
        break  # sai do loop, encerra main.py

    except subprocess.TimeoutExpired:
        print(f"\n[ERRO] gerar-relatorio.py ultrapassou timeout de {timeout_segundos} segundos e foi encerrado.\n")
    except subprocess.CalledProcessError as e:
        print(f"\n[ERRO] gerar-relatorio.py retornou erro de execução: {e}\n")
    except Exception:
        print("\n[ERRO] Ocorreu um erro inesperado ao tentar executar gerar-relatorio.py:")
        traceback.print_exc()

    if max_tentativas is not None and tentativa_atual >= max_tentativas:
        print(f"[INFO] Número máximo de tentativas ({max_tentativas}) atingido. Encerrando.")
        break

    print("[INFO] Aguardando 10 segundos antes de tentar novamente...")
    time.sleep(10)
