from modules.clicar_imagem import clicar_imagem
from modules.localizar_imagem import localizar_imagem, salvar_screenshot
import time
from config.list_filial import LISTA_FILIAIS
import pyautogui
import pyperclip
from datetime import datetime
from modules.aguardar_download_inteligente import aguardar_download_completo, listar_pids_excel
import os

# Pausa entre etapas de teclado que não têm confirmação visual (tab, digitação de campos).
# Pode ser ajustada no .env (PAUSA_CAMPO=3) sem mexer no código.
PAUSA_CAMPO = float(os.getenv("PAUSA_CAMPO", "3"))

# Pausa curta antes de procurar um botão: o clicar_imagem já espera o botão
# aparecer (até o timeout), então aqui basta dar tempo para a tela "assentar".
PAUSA_ANTES_CLIQUE = 1.5

# Quantas vezes tentar novamente as filiais que falharam (ao final da primeira passada)
TENTATIVAS_EXTRAS = int(os.getenv("TENTATIVAS_EXTRAS", "1"))

IMG_CONSULTA_CONTABIL = "data/cosulta_contabil.png"


def colar_texto(texto: str) -> None:
    """
    Cola o texto via área de transferência (Ctrl+V).
    Mais rápido que digitar e funciona com acentos e com '\\' em qualquer layout de teclado.
    """
    pyperclip.copy(texto)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')


def _clicar_se_visivel(caminho_imagem, descricao, **kwargs) -> bool:
    """Clica na imagem somente se ela já estiver na tela (sem esperar nem gerar screenshot)."""
    localizacao = localizar_imagem(caminho_imagem, timeout=2, descricao=descricao,
                                   screenshot_se_falhar=False, **kwargs)
    if not localizacao:
        return False
    pyautogui.click(pyautogui.center(localizacao))
    print(f"✓ Clicou em: {descricao}")
    time.sleep(2)
    return True


def _consulta_visivel() -> bool:
    """Verifica se a opção Consulta Contábil está na tela (com o clique de foco do início da filial)."""
    if localizar_imagem(IMG_CONSULTA_CONTABIL, timeout=2, screenshot_se_falhar=False):
        return True
    pyautogui.click(x=100, y=200)
    time.sleep(2)
    return localizar_imagem(IMG_CONSULTA_CONTABIL, timeout=2, screenshot_se_falhar=False) is not None


def voltar_estado_inicial(tentativas: int = 4) -> bool:
    """
    Tenta deixar a tela no ponto de partida de cada filial (menu Específicos,
    com a opção Consulta Contábil visível), fechando o que ficou aberto após um erro.

    O Protheus Webapp não fecha os diálogos com Esc, por isso a ordem é:
      1. "Fechar" em mensagens de erro (ex.: "O intervalo de Lote é inválido")
      2. "Cancelar" na janela de Parametros
      3. "X" da aba da rotina (Rel. Consulta Mov. Contabeis)
      4. Esc, como último recurso

    Returns:
        True se a opção Consulta Contábil ficou visível.
    """
    print("↩️  Voltando para a tela inicial da rotina...")

    largura, _ = pyautogui.size()
    regiao_abas = (0, 25, largura, 40)  # barra azul das abas, no topo da tela

    for _ in range(tentativas):
        if _consulta_visivel():
            return True

        clicou = False
        clicou |= _clicar_se_visivel("data/botao_fechar.png", "Fechar (mensagem de erro)")
        clicou |= _clicar_se_visivel("data/botao_cancelar_parametros.png", "Cancelar (Parametros)")

        if not clicou:
            clicou = _clicar_se_visivel("data/fechar_aba_rotina.png", "Fechar aba da rotina",
                                        region=regiao_abas)

        if not clicou:
            pyautogui.press('esc')
            time.sleep(2)

    if _consulta_visivel():
        return True

    print("⚠️ Não foi possível confirmar o retorno à tela inicial da rotina.")
    salvar_screenshot("falha_voltar_estado_inicial")
    return False


def processar_filial(filial: str, competencia: str) -> str:
    """
    Executa o fluxo completo de uma filial: preenche os parâmetros do relatório,
    salva o arquivo e aguarda o download/conversão.

    Returns:
        Caminho do arquivo .xlsx gerado.

    Raises:
        RuntimeError em caso de falha em qualquer etapa.
    """
    # Foco na tela
    pyautogui.click(x=100, y=200)
    time.sleep(PAUSA_ANTES_CLIQUE)

    if not clicar_imagem(IMG_CONSULTA_CONTABIL, confidence=0.8,
                         timeout=15, descricao="Opção consulta_contabil"):
        raise RuntimeError("Erro ao clicar na opção consulta_contabil.")

    time.sleep(4)
    pyautogui.press('tab', presses=2, interval=0.5)

    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    time.sleep(2)
    pyautogui.typewrite(filial, interval=0.3)

    time.sleep(2)
    if not clicar_imagem("data/botao_confirmar.png", confidence=0.9, grayscale=False,
                         timeout=15, descricao="Botão Confirmar"):
        raise RuntimeError("Erro ao clicar no botão Confirmar.")

    # O clicar_imagem abaixo já aguarda o botão aparecer por até 15s
    time.sleep(3)

    # Botão reforma tributária (não fatal)
    if not clicar_imagem("data/botao_reforma_tributaria.png", confidence=0.8,
                         timeout=15, descricao="Botão Reforma Tributária"):
        print("⚠️  Botão Reforma Tributária não encontrado — continuando.")

    time.sleep(PAUSA_CAMPO)
    pyautogui.press('backspace')
    time.sleep(PAUSA_CAMPO)
    pyautogui.typewrite(filial, interval=0.2)

    time.sleep(PAUSA_CAMPO)
    pyautogui.press('backspace')
    time.sleep(PAUSA_CAMPO)
    pyautogui.typewrite(filial, interval=0.2)

    # Data início
    pyautogui.press('backspace')
    competencia_inicio = f"01/{competencia[3:]}"
    pyautogui.typewrite(competencia_inicio, interval=0.2)
    time.sleep(PAUSA_CAMPO)

    # Data fim
    pyautogui.press('backspace')
    pyautogui.typewrite(competencia, interval=0.2)
    time.sleep(PAUSA_CAMPO)

    pyautogui.press('backspace')
    pyautogui.press('tab', presses=1, interval=0.5)
    time.sleep(PAUSA_CAMPO)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite("zzzzzz", interval=0.2)
    time.sleep(PAUSA_CAMPO)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite("4", interval=0.2)
    time.sleep(PAUSA_CAMPO)

    pyautogui.press('tab', presses=1, interval=0.5)
    time.sleep(PAUSA_CAMPO)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite("5", interval=0.2)
    time.sleep(PAUSA_CAMPO)

    pyautogui.press('tab', presses=1, interval=0.5)
    pyautogui.press('backspace')
    time.sleep(PAUSA_CAMPO)

    pyautogui.press('tab', presses=1, interval=0.5)
    pyautogui.press('backspace')
    pyautogui.typewrite("zzzzzzzzzzzzzzz", interval=0.2)
    time.sleep(PAUSA_ANTES_CLIQUE)

    if not clicar_imagem("data/botao_diretorio.png", confidence=0.8,
                         timeout=15, descricao="Botão diretório"):
        raise RuntimeError("Erro ao clicar no botão para selecionar diretório.")

    time.sleep(PAUSA_ANTES_CLIQUE)

    if not clicar_imagem("data/nome_diretorio.png", confidence=0.8,
                         timeout=15, descricao="Botão nome diretório"):
        raise RuntimeError("Erro ao clicar no botão para selecionar nome diretório.")

    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    # Monta caminho de saída
    data           = datetime.strptime(competencia, "%d/%m/%Y")
    ano            = data.year
    mes            = str(data.month).zfill(2)
    caminho_fixo   = os.getenv("CAMINHO_FIXO_RAZAO")
    if not caminho_fixo:
        raise RuntimeError("Variável CAMINHO_FIXO_RAZAO não configurada no .env.")
    caminho_fixo_completo = f"{caminho_fixo}\\{ano}\\{mes}_{ano}"
    nome_arquivo = f"Razao_Filial_{filial}"
    caminho_fixo_completo_com_filial = f"{caminho_fixo_completo}\\{nome_arquivo}"

    os.makedirs(caminho_fixo_completo, exist_ok=True)

    print(f"📂 Caminho completo: {caminho_fixo_completo}")
    print(f"📝 Nome do arquivo : {nome_arquivo}")
    time.sleep(2)

    pyautogui.press('backspace')
    colar_texto(caminho_fixo_completo_com_filial)

    time.sleep(2)
    if not clicar_imagem("data/abrir_diretorio.png", confidence=0.9, grayscale=False,
                         timeout=15, descricao="Botão Abrir Diretório"):
        raise RuntimeError("Erro ao clicar no botão Abrir Diretório.")

    time.sleep(PAUSA_ANTES_CLIQUE)

    # Marca o momento do disparo do download e os Excel já abertos,
    # para não confundir com arquivos/planilhas de antes
    pids_excel_antes = listar_pids_excel()
    inicio_download  = time.time()

    if not clicar_imagem("data/botao_salvar_arquivo_final.png", confidence=0.8,
                         timeout=15, descricao="Botão Salvar Arquivo"):
        raise RuntimeError("Erro ao clicar no botão Salvar Arquivo.")

    print("🔍 Aguardando conclusão do download...")
    time.sleep(PAUSA_ANTES_CLIQUE)

    if not clicar_imagem("data/botao_fechar.png", confidence=0.8,
                         timeout=15, descricao="Botão Fechar"):
        raise RuntimeError("Erro ao clicar no botão Fechar.")

    sucesso, arquivo_baixado, tempo_gasto = aguardar_download_completo(
        diretorio_temp=caminho_fixo_completo,
        nome_arquivo_esperado=nome_arquivo,
        inicio_download=inicio_download,
        pids_excel_antes=pids_excel_antes,
        timeout=1200,
        intervalo_verificacao=3,
    )

    if not sucesso:
        raise RuntimeError(f"Download não concluído para a filial {filial}.")

    print(f"⚡ Tempo total: {tempo_gasto:.1f}s")
    print(f"💾 Arquivo final: {arquivo_baixado}")
    return arquivo_baixado


def automacao_razao(competencia: str, log=None, filiais=None):
    """
    Automação para download do relatório de razão contábil por filial.

    Args:
        competencia: data no formato "dd/mm/YYYY" (último dia do mês).
        log: instância de LogExecucao (opcional). Quando fornecida, registra
             o resultado de cada filial automaticamente.
        filiais: lista de filiais a processar (padrão: LISTA_FILIAIS).

    Raises:
        RuntimeError se não conseguir abrir o menu da rotina.
    """
    print("🚀 Iniciando automação razão...")

    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8,
                         timeout=15, descricao="Menu Relatórios"):
        raise RuntimeError("Erro ao acessar o menu Relatórios.")

    time.sleep(PAUSA_ANTES_CLIQUE)

    if not clicar_imagem("data/especificos.png", confidence=0.8,
                         timeout=15, descricao="Opção Específicos"):
        raise RuntimeError("Erro ao clicar na opção Específicos.")

    time.sleep(PAUSA_CAMPO)

    pendentes = list(filiais if filiais else LISTA_FILIAIS)

    for tentativa in range(1, TENTATIVAS_EXTRAS + 2):
        if not pendentes:
            break

        if tentativa > 1:
            print(f"\n🔁 Nova tentativa ({tentativa}ª) para {len(pendentes)} filial(is): "
                  f"{', '.join(pendentes)}")
            voltar_estado_inicial()

        falharam = []

        for filial in pendentes:
            print(f"\n{'='*60}")
            print(f"🏢 Processando filial: {filial} (tentativa {tentativa})")
            print(f"{'='*60}\n")

            inicio_filial = datetime.now()   # marca início para calcular tempo no log

            try:
                arquivo_baixado = processar_filial(filial, competencia)
                print(f"✅ Filial {filial} processada com sucesso!")

                if log:
                    sufixo = f" ({tentativa}ª tentativa)" if tentativa > 1 else ""
                    log.registrar_filial(
                        filial=filial,
                        sucesso=True,
                        mensagem=f"Arquivo: {os.path.basename(arquivo_baixado)}{sufixo}",
                        inicio_filial=inicio_filial,
                    )

            except Exception as e:
                msg_erro = str(e)
                print(f"❌ Erro na filial {filial}: {msg_erro}")
                salvar_screenshot(f"erro_filial_{filial}")
                falharam.append(filial)

                if log:
                    log.registrar_filial(
                        filial=filial,
                        sucesso=False,
                        mensagem=f"{msg_erro} (tentativa {tentativa})",
                        inicio_filial=inicio_filial,
                    )

                # Fecha o que ficou aberto para não contaminar a próxima filial
                voltar_estado_inicial()

        pendentes = falharam

    if not clicar_imagem("data/logo_p_voltar_a_tela_inicial.png", confidence=0.8,
                         timeout=15, descricao="Logo voltar tela inicial"):
        print("⚠️ Não clicou no logo para voltar à tela inicial — continuando.")

    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8,
                         timeout=15, descricao="Menu Relatórios (fechar)"):
        print("⚠️ Não conseguiu fechar o menu Relatórios — continuando.")

    print("\n" + "=" * 60)
    print("✅ Automação de razão concluída para todas as filiais!")
    print("=" * 60)
