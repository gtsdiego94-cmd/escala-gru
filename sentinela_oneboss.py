import os
import time
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Pastas que serão vigiadas
PASTAS_PARA_VIGIAR = [
    r"C:\Relatorios",
    r"C:\Export"
]

REPO_DIR = r"C:\Users\User\Desktop\escala-gru"
SCRIPT_PROCESSAR = os.path.join(REPO_DIR, "atualizar_vendas_python.py")
PYTHON_EXE = r"C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"

class RelatorioWatcherHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.ultimo_processamento = 0

    def processar(self, caminho_arquivo):
        # Apenas arquivos de relatório
        ext = os.path.splitext(caminho_arquivo)[1].lower()
        if ext not in [".xlsx", ".xls", ".csv"]:
            return

        nome = os.path.basename(caminho_arquivo)
        if nome.startswith("~$") or "relatorio" not in nome.lower():
            return

        # Evita disparos duplicados em menos de 10 segundos
        agora = time.time()
        if agora - self.ultimo_processamento < 10:
            return

        print("\n" + "=" * 60)
        print(f"🚨 NOVO RELATÓRIO DETECTADO INSTANTANEAMENTE!")
        print(f"📁 Arquivo: {nome}")
        print("⏳ Aguardando 3 segundos para conclusão da gravação do OneBoss...")
        print("=" * 60)

        # Espera o arquivo ser fechado pelo OneBoss
        time.sleep(3)

        print("\n⚡ DISPARANDO ATUALIZAÇÃO AUTOMÁTICA VIA PYTHON...")
        self.ultimo_processamento = time.time()

        try:
            subprocess.run([PYTHON_EXE, SCRIPT_PROCESSAR], check=True)
            print("\n✅ CICLO COMPLETO CONCLUÍDO COM SUCESSO!")
            print("👀 Sentinela continua vigiando a pasta para novos relatórios...\n")
        except Exception as e:
            print(f"❌ Erro ao executar processador: {e}")

    def on_created(self, event):
        if not event.is_directory:
            self.processar(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self.processar(event.src_path)

def iniciar_sentinela():
    print("=" * 60)
    print("  🛡️  SENTINELA ONEBOSS ATIVADO (MONITORAMENTO EM TEMPO REAL)")
    print("=" * 60)
    print("O sentinela está vigiando as pastas:")
    for pasta in PASTAS_PARA_VIGIAR:
        if not os.path.exists(pasta):
            os.makedirs(pasta, exist_ok=True)
        print(f"   👁️  {pasta}")

    print("\nAssim que um novo relatório cair na pasta, ele atualiza a página na hora!")
    print("Pressione Ctrl + C para encerrar o sentinela.\n")

    event_handler = RelatorioWatcherHandler()
    observer = Observer()

    for pasta in PASTAS_PARA_VIGIAR:
        observer.schedule(event_handler, path=pasta, recursive=False)

    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    iniciar_sentinela()
