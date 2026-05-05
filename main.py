import pyautogui
import time
import os
from dotenv import load_dotenv
from modules.clicar_imagem import clicar_imagem
from modules.razao import automacao_razao
from modules.calcular_competencia import calcular_competencia
from modules.conectar_vpn import conectar_vpn
from modules.abrir_app_agent import habilitar_app_agent
from modules.logger_excel import LogExecucao
from modules.enviar_email import enviar_email_resultado
from config.list_filial import LISTA_FILIAIS

if __name__ == "__main__":

    # Carregar variáveis de ambiente
    load_dotenv()

    # Diretório raiz do projeto (onde está o main.py)
    # O log será salvo em: <RAIZ_PROJETO>/log/log_automacao.xlsx
    RAIZ_PROJETO = os.path.dirname(os.path.abspath(__file__))

    # Calcula a competência (último dia do mês anterior)
    competencia_anterior = calcular_competencia()

    # conectar_vpn()

    try:
        NOME_APP = os.getenv("NOME_APP")
        USER     = os.getenv("USER")
        SENHA    = os.getenv("SENHA")

        if not NOME_APP or not USER or not SENHA:
            print("Erro: Variáveis de ambiente não configuradas!")
            print(f"NOME_APP: {NOME_APP}")
            print(f"USER: {USER}")
            print(f"SENHA: {'***' if SENHA else None}")
            exit(1)

        print("✅ Variáveis carregadas com sucesso!")

    except Exception as e:
        print("Erro ao carregar variáveis de ambiente:", e)
        exit(1)

    pyautogui.press('win')
    pyautogui.write(NOME_APP, interval=0.1)
    pyautogui.press('enter')
    time.sleep(10)

    pyautogui.press('enter')
    time.sleep(20)

    habilitar_app_agent()
    time.sleep(15)

    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    pyautogui.press('backspace')

    pyautogui.write(USER.upper(), interval=0.1)
    pyautogui.press('tab')
    pyautogui.write(SENHA, interval=0.1)
    pyautogui.press('enter')

    time.sleep(5)
    pyautogui.click(x=100, y=200)
    time.sleep(5)
    pyautogui.press('end')
    time.sleep(2)

    if not clicar_imagem("data/botao_entrar.png", confidence=0.8,
                         timeout=15, descricao="Botão entrar"):
        print("Erro ao clicar no botão entrar.")
        exit(1)

    time.sleep(3)

    log_razao = LogExecucao(raiz_projeto=RAIZ_PROJETO)
    log_razao.iniciar_execucao(
        tipo="Razão",
        competencia=competencia_anterior,
        filiais=LISTA_FILIAIS,
    )

    automacao_razao(competencia_anterior, log=log_razao)

    resumo_razao = log_razao.finalizar_execucao()
    enviar_email_resultado(resumo_razao)

    print("\n✅ Automação concluída com sucesso!")

    pyautogui.keyDown('ctrl')
    pyautogui.press('q')
    pyautogui.keyUp('ctrl')
    time.sleep(1)

    if not clicar_imagem("data/botao_finalizar.png", confidence=0.8,
                         timeout=15, descricao="Botão Finalizar"):
        print("Erro ao clicar no botão finalizar.")

    pyautogui.press('f11')

    pyautogui.keyDown('alt')
    pyautogui.press('f4')
    pyautogui.keyUp('alt')

    exit(0)