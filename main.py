import os
from dotenv import load_dotenv, dotenv_values

# Carregar variáveis de ambiente antes de importar os módulos,
# pois alguns leem configurações (ex.: PAUSA_CAMPO) na importação
load_dotenv()

import pyautogui
import time
import sys
from modules.clicar_imagem import clicar_imagem
from modules.localizar_imagem import localizar_imagem, salvar_screenshot
from modules.razao import automacao_razao
from modules.calcular_competencia import calcular_competencia
from modules.conectar_vpn import conectar_vpn
from modules.abrir_app_agent import habilitar_app_agent
from modules.logger_excel import LogExecucao
from modules.enviar_email import enviar_email_resultado
from config.list_filial import LISTA_FILIAIS


def clicar_campo(caminho_icone, descricao, deslocamento_x=80, timeout=20):
    """
    Localiza o ícone que fica dentro de um campo de texto e clica à direita dele,
    na área de digitação, para garantir que o campo receba o foco.
    """
    localizacao = localizar_imagem(caminho_icone, confidence=0.8,
                                   timeout=timeout, descricao=descricao)
    if not localizacao:
        return False
    centro = pyautogui.center(localizacao)
    pyautogui.click(centro.x + deslocamento_x, centro.y)
    print(f"✓ Clicou em: {descricao}")
    return True


def fazer_login(nome_app, usuario, senha):
    """Abre o Protheus, habilita o WebAgent e faz login. Lança RuntimeError se falhar."""
    pyautogui.press('win')
    pyautogui.write(nome_app, interval=0.1)
    pyautogui.press('enter')

    # Aguarda a tela "Parâmetros Iniciais" (SIGAMDI / ambiente) abrir e confirma no OK.
    # O tempo de abertura do Webapp varia, por isso não usa espera fixa.
    if not clicar_imagem("data/botao_ok_parametros_iniciais.png", confidence=0.8,
                         timeout=90, descricao="OK (Parâmetros Iniciais)"):
        print("⚠️ Tela Parâmetros Iniciais não encontrada — tentando com Enter.")
        pyautogui.press('enter')

    # O habilitar_app_agent aguarda o botão aparecer (até 40s), por isso não há sleep fixo aqui
    habilitar_app_agent(timeout=40)
    time.sleep(15)

    # O cursor abre no campo da senha: clica no campo do usuário e limpa o valor
    # preenchido automaticamente (ex.: "SEDE\larissa.miranda")
    if not clicar_campo("data/campo_usuario.png", "Campo usuário"):
        raise RuntimeError("Campo de usuário não encontrado na tela de login.")
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    time.sleep(0.5)
    pyautogui.write(usuario.upper(), interval=0.1)
    time.sleep(1)

    # Fecha a lista de sugestões que o sistema abre ao digitar o usuário
    pyautogui.press('esc')
    time.sleep(1)

    if not clicar_campo("data/campo_senha.png", "Campo senha"):
        raise RuntimeError("Campo de senha não encontrado na tela de login.")
    time.sleep(1)
    pyautogui.write(senha, interval=0.1)
    pyautogui.press('enter')

    time.sleep(5)
    pyautogui.click(x=100, y=200)
    time.sleep(5)
    pyautogui.press('end')
    time.sleep(2)

    if not clicar_imagem("data/botao_entrar.png", confidence=0.8,
                         timeout=15, descricao="Botão entrar"):
        raise RuntimeError("Erro ao clicar no botão entrar (login não concluído).")

    time.sleep(3)


def encerrar_protheus():
    """Faz logout e fecha o Protheus."""
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


if __name__ == "__main__":

    # Diretório raiz do projeto (onde está o main.py)
    # O log será salvo em: <RAIZ_PROJETO>/log/log_automacao.xlsx
    RAIZ_PROJETO = os.path.dirname(os.path.abspath(__file__))
    os.chdir(RAIZ_PROJETO)  # garante que os caminhos relativos (data/, log/) funcionem

    # Calcula a competência (último dia do mês anterior, ou COMPETENCIA do .env)
    competencia_anterior = calcular_competencia()
    # competencia_anterior = "31/08/2026"

    # conectar_vpn()

    # USER/SENHA são lidos direto do arquivo .env: variáveis do sistema com o mesmo
    # nome (ex.: USER definido pelo Git Bash) teriam prioridade no os.getenv.
    valores_env = dotenv_values(os.path.join(RAIZ_PROJETO, ".env"))
    NOME_APP = os.getenv("NOME_APP")
    USER     = os.getenv("PROTHEUS_USER") or valores_env.get("USER")
    SENHA    = os.getenv("PROTHEUS_SENHA") or valores_env.get("SENHA")

    if not NOME_APP or not USER or not SENHA:
        print("Erro: Variáveis de ambiente não configuradas!")
        print(f"NOME_APP: {NOME_APP}")
        print(f"USER: {USER}")
        print(f"SENHA: {'***' if SENHA else None}")
        sys.exit(1)

    print("✅ Variáveis carregadas com sucesso!")

    # Filiais a processar: as passadas na linha de comando
    # (ex.: python main.py 01010037 01010038) ou, se nenhuma, a LISTA_FILIAIS completa
    filiais_execucao = [f.strip() for f in sys.argv[1:] if f.strip()] or list(LISTA_FILIAIS)
    if sys.argv[1:]:
        print(f"🎯 Processando somente as filiais informadas: {', '.join(filiais_execucao)}")

    log_razao = LogExecucao(raiz_projeto=RAIZ_PROJETO)
    log_razao.iniciar_execucao(
        tipo="Razão",
        competencia=competencia_anterior,
        filiais=filiais_execucao,
    )

    mensagem_nao_processada = "Filial não processada"
    codigo_saida = 0

    try:
        fazer_login(NOME_APP, USER, SENHA)
        automacao_razao(competencia_anterior, log=log_razao, filiais=filiais_execucao)
        print("\n✅ Automação concluída!")

    except Exception as e:
        print(f"\n❌ Automação interrompida: {e}")
        salvar_screenshot("automacao_interrompida")
        mensagem_nao_processada = f"Não processada — automação interrompida: {e}"
        codigo_saida = 1

    finally:
        # Log e e-mail são enviados sempre, mesmo se a automação parar no meio
        resumo_razao = log_razao.finalizar_execucao(mensagem_nao_processada)
        enviar_email_resultado(resumo_razao)

    try:
        encerrar_protheus()
    except Exception as e:
        print(f"⚠️ Erro ao encerrar o Protheus: {e}")

    sys.exit(codigo_saida)
