from modules.clicar_imagem import clicar_imagem
import time
from config.list_filial import LISTA_FILIAIS
import pyautogui
from datetime import datetime
from modules.aguardar_download_inteligente import aguardar_download_completo, fechar_excel
import os


def automacao_razao(competencia: str, log=None):
    """
    Automação para download do relatório de razão contábil por filial.

    Args:
        competencia: data no formato "dd/mm/YYYY" (último dia do mês).
        log: instância de LogExecucao (opcional). Quando fornecida, registra
             o resultado de cada filial automaticamente.
    """
    print("🚀 Iniciando automação razão...")

    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8,
                         timeout=15, descricao="Menu Relatórios"):
        print("Erro ao acessar o menu Relatórios.")
        return

    time.sleep(2)

    if not clicar_imagem("data/especificos.png", confidence=0.8,
                         timeout=15, descricao="Opção Específicos"):
        print("Erro ao clicar na opção Específicos.")
        return

    for filial in LISTA_FILIAIS:
        print(f"\n{'='*60}")
        print(f"🏢 Processando filial: {filial}")
        print(f"{'='*60}\n")

        inicio_filial = datetime.now()   # marca início para calcular tempo no log

        try:
            # Foco na tela
            pyautogui.click(x=100, y=200)
            time.sleep(2)

            if not clicar_imagem("data/cosulta_contabil.png", confidence=0.8,
                                 timeout=15, descricao="Opção consulta_contabil"):
                raise RuntimeError("Erro ao clicar na opção consulta_contabil.")

            time.sleep(2)
            pyautogui.press('tab', presses=2, interval=0.5)

            time.sleep(2)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')

            time.sleep(2)
            pyautogui.typewrite(filial, interval=0.2)

            time.sleep(2)
            if not clicar_imagem("data/botao_confirmar.png", confidence=0.8,
                                 timeout=15, descricao="Botão Confirmar"):
                raise RuntimeError("Erro ao clicar no botão Confirmar.")

            time.sleep(8)

            # Botão reforma tributária (não fatal)
            if not clicar_imagem("data/botao_reforma_tributaria.png", confidence=0.8,
                                 timeout=15, descricao="Botão Reforma Tributária"):
                print("⚠️  Botão Reforma Tributária não encontrado — continuando.")

            time.sleep(2)
            pyautogui.press('backspace')
            time.sleep(2)
            pyautogui.typewrite(filial, interval=0.2)

            time.sleep(2)
            pyautogui.press('backspace')
            time.sleep(2)
            pyautogui.typewrite(filial, interval=0.2)

            # Data início
            pyautogui.press('backspace')
            competencia_inicio = f"01/{competencia[3:]}"
            pyautogui.typewrite(competencia_inicio, interval=0.2)
            time.sleep(2)

            # Data fim
            pyautogui.press('backspace')
            pyautogui.typewrite(competencia, interval=0.2)
            time.sleep(2)

            pyautogui.press('backspace')
            pyautogui.press('tab', presses=1, interval=0.5)
            time.sleep(2)

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.typewrite("zzzzzz", interval=0.2)
            time.sleep(2)

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.typewrite("4", interval=0.2)
            time.sleep(2)

            pyautogui.press('tab', presses=1, interval=0.5)
            time.sleep(2)

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.typewrite("5", interval=0.2)
            time.sleep(2)

            pyautogui.press('tab', presses=1, interval=0.5)
            pyautogui.press('backspace')
            time.sleep(2)

            pyautogui.press('tab', presses=1, interval=0.5)
            pyautogui.press('backspace')
            pyautogui.typewrite("zzzzzzzzzzzzzzz", interval=0.2)
            time.sleep(2)

            if not clicar_imagem("data/botao_diretorio.png", confidence=0.8,
                                 timeout=15, descricao="Botão diretório"):
                raise RuntimeError("Erro ao clicar no botão para selecionar diretório.")

            time.sleep(2)

            if not clicar_imagem("data/nome_diretorio.png", confidence=0.8,
                                 timeout=15, descricao="Botão nome diretório"):
                raise RuntimeError("Erro ao clicar no botão para selecionar nome diretório.")

            time.sleep(2)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            time.sleep(2)

            # Monta caminho de saída
            data           = datetime.strptime(competencia, "%d/%m/%Y")
            ano            = data.year
            mes            = str(data.month).zfill(2)
            caminho_fixo   = os.getenv("CAMINHO_FIXO_RAZAO")
            caminho_fixo_completo = f"{caminho_fixo}\\{ano}\\{mes}_{ano}"
            caminho_fixo_completo_com_filial = (
                f"{caminho_fixo_completo}\\Razao_Filial_{filial}"
            )

            if not os.path.exists(caminho_fixo_completo):
                os.makedirs(caminho_fixo_completo)

            print(f"📂 Caminho completo: {caminho_fixo_completo}")
            print(f"📝 Nome do arquivo : Razao_Filial_{filial}")
            time.sleep(2)

            pyautogui.press('backspace')
            pyautogui.typewrite(caminho_fixo_completo_com_filial, interval=0.1)

            time.sleep(2)
            if not clicar_imagem("data/abrir_diretorio.png", confidence=0.9,
                                 timeout=15, descricao="Botão Abrir Diretório"):
                raise RuntimeError("Erro ao clicar no botão Abrir Diretório.")

            time.sleep(2)
            if not clicar_imagem("data/botao_salvar_arquivo_final.png", confidence=0.8,
                                 timeout=15, descricao="Botão Salvar Arquivo"):
                raise RuntimeError("Erro ao clicar no botão Salvar Arquivo.")

            print("🔍 Aguardando conclusão do download...")
            time.sleep(3)

            if not clicar_imagem("data/botao_fechar.png", confidence=0.8,
                                 timeout=15, descricao="Botão Fechar"):
                raise RuntimeError("Erro ao clicar no botão Fechar.")

            time.sleep(3)

            sucesso, arquivo_baixado, tempo_gasto = aguardar_download_completo(
                diretorio_temp=caminho_fixo_completo,
                nome_arquivo_esperado=f"Razao_Filial_{filial}",
                timeout=900,
                intervalo_verificacao=3,
            )

            if not sucesso:
                raise RuntimeError(f"Download não concluído para a filial {filial}.")

            print(f"⚡ Tempo total: {tempo_gasto:.1f}s")
            print(f"💾 Arquivo final: {arquivo_baixado}")
            print(f"✅ Filial {filial} processada com sucesso!")

            # Registra sucesso no log
            if log:
                log.registrar_filial(
                    filial=filial,
                    sucesso=True,
                    mensagem=f"Arquivo: {os.path.basename(arquivo_baixado)}",
                    inicio_filial=inicio_filial,
                )

        except Exception as e:
            msg_erro = str(e)
            print(f"❌ Erro na filial {filial}: {msg_erro}")

            # Registra erro no log
            if log:
                log.registrar_filial(
                    filial=filial,
                    sucesso=False,
                    mensagem=msg_erro,
                    inicio_filial=inicio_filial,
                )
            continue   # segue para a próxima filial mesmo com erro

    if not clicar_imagem("data/logo_p_voltar_a_tela_inicial.png", confidence=0.8,
                         timeout=15, descricao="Logo voltar tela inicial"):
        print("Erro ao clicar no logo para voltar à tela inicial.")

    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8,
                         timeout=15, descricao="Menu Relatórios (fechar)"):
        print("Erro ao fechar o menu Relatórios.")
        exit(1)

    print("\n" + "=" * 60)
    print("✅ Automação de razão concluída para todas as filiais!")
    print("=" * 60)