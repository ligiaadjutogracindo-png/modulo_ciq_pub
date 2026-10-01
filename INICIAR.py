"""
INICIAR.py
-----------
Script auxiliar para rodar o app Desempenho Analítico sem precisar abrir CMD/PowerShell.
"""
import subprocess
import sys
import os

PASTA = os.path.dirname(os.path.abspath(__file__))
os.chdir(PASTA)

print("=" * 60)
print("Desempenho Analítico - preparando ambiente")
print("=" * 60)

print("\n[0/3] Verificando se o pip está disponível...")
tem_pip = subprocess.run([sys.executable, "-m", "pip", "--version"],
                          capture_output=True, text=True).returncode == 0

if not tem_pip:
    print("pip não encontrado. Tentando recuperar com ensurepip...")
    r = subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"],
                        capture_output=True, text=True)
    print(r.stdout)
    if r.returncode != 0:
        print("STDERR:", r.stderr)
        print("\nNão foi possível recuperar o pip automaticamente.")
        print("Detalhe do erro:")
        print(r.stderr)
        sys.exit(1)
    else:
        print("pip recuperado com sucesso via ensurepip.")
else:
    print("pip já está disponível.")

print("\n[1/2] Instalando bibliotecas necessárias (pode levar um minuto)...")
r = subprocess.run([
    sys.executable, "-m", "pip", "install", "-q",
    "streamlit", "pandas", "plotly", "openpyxl",
], capture_output=True, text=True)
print(r.stdout)
if r.returncode != 0:
    print("\nERRO ao instalar as bibliotecas.")
    print("STDOUT:", r.stdout)
    print("STDERR:", r.stderr)
    sys.exit(1)
print("Bibliotecas instaladas com sucesso.")

print("\n[2/2] Iniciando o app (vai abrir uma aba no navegador)...")
print("Deixe esta execução rodando enquanto estiver usando o app.")
print("Para fechar o app, clique no quadrado vermelho (Stop) no VS Code.\n")

try:
    subprocess.check_call([sys.executable, "-m", "streamlit", "run", "app.py"])
except subprocess.CalledProcessError as e:
    print("\nERRO ao iniciar o Streamlit.")
    print("Detalhe do erro:", e)
    sys.exit(1)
except KeyboardInterrupt:
    print("\nApp encerrado.")
