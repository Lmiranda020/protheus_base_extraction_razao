from modules.clicar_imagem import clicar_imagem
import time
from config.list_filial import LISTA_FILIAIS
import pyautogui
from datetime import datetime
from modules.aguardar_download_inteligente import aguardar_download_completo, fechar_excel
import os

def automacao_razao(competencia):
    """
    Automação para download do relatório de centro de custo
    """ 
    print("🚀Iniciando automação razão...")
    # no menu a oção "Relatórios"
    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8, timeout=15, descricao="Menu Relatórios"):
        print("Erro ao acessar o menu Relatórios.")
        return
    
    time.sleep(2)
    # clicar na opção "específicos"
    if not clicar_imagem("data/especificos.png", confidence=0.8, timeout=15, descricao="Opção especificos"):
        print("Erro ao clicar na opção Específicos.")
        return

    for filial in LISTA_FILIAIS:
        print(f"\n{'='*60}")
        print(f"🏢 Processando filial: {filial}")
        print(f"{'='*60}\n")
        
        # clicar no meio da tela para garantir o foco
        pyautogui.click(x=100, y=200)
        
        time.sleep(2)
        # clicar na opção "cosulta_contabil"
        if not clicar_imagem("data/cosulta_contabil.png", confidence=0.8, timeout=15, descricao="Opção cosulta_contabil"):
            print("Erro ao clicar na opção cosulta_contabil.")
            return
        
        time.sleep(2)
        
        # clicar duas vezes o tab
        pyautogui.press('tab', presses=2, interval=0.5)

        time.sleep(2)
        # seleciona todo o campo 
        pyautogui.keyDown('ctrl')
        pyautogui.press('a')
        pyautogui.keyUp('ctrl')

        time.sleep(2)
        # apaga o conteúdo do campo
        pyautogui.press('backspace')

        time.sleep(2)
        # digita a filial
        pyautogui.typewrite(filial, interval=0.2)

        # clica no botão "Confirmar"
        time.sleep(2)
        if not clicar_imagem("data/botao_confirmar.png", confidence=0.8, timeout=15, descricao="Botão Confirmar"):
            print("Erro ao clicar no botão Confirmar.")
            return
        
        time.sleep(8)

        # clicar no botão reforma tributaria
        if not clicar_imagem("data/botao_reforma_tributaria.png", confidence=0.8, timeout=15, descricao="Botão Reforma Tributária"):
            print("Erro ao clicar no botão Reforma Tributária.")

        time.sleep(2)
        # apaga o conteúdo do campo
        pyautogui.press('backspace')

        time.sleep(2)
        # digita a filial
        pyautogui.typewrite(filial, interval=0.2)

        time.sleep(2)
        # apaga o conteúdo do campo
        pyautogui.press('backspace')

        time.sleep(2)
        # digita a filial
        pyautogui.typewrite(filial, interval=0.2)

        #digita a data do primeiro mes
        # apaga o conteúdo do campo
        pyautogui.press('backspace')
        competencia_inicio = f"01/{competencia[3:]}"
        pyautogui.typewrite(competencia_inicio, interval=0.2)

        # escreve a competencia final
        # apaga o conteúdo do campo 
        pyautogui.press('backspace')
        pyautogui.typewrite(competencia, interval=0.2)        

        # clicar duas vezes o tab
        pyautogui.press('backspace')
        pyautogui.press('tab', presses=1, interval=0.5)


        #digitar zzzzz
        pyautogui.keyDown('ctrl')
        pyautogui.press('a')
        pyautogui.keyUp('ctrl')
        pyautogui.press('backspace', presses=1, interval=0.1)
        pyautogui.typewrite("zzzzzz", interval=0.2)


        pyautogui.keyDown('ctrl')
        pyautogui.press('a')
        pyautogui.keyUp('ctrl')
        pyautogui.press('backspace', presses=1, interval=0.1)
        pyautogui.typewrite("4", interval=0.2)

        # clicar duas vezes o tab
        pyautogui.press('tab', presses=1, interval=0.5)

        pyautogui.keyDown('ctrl')
        pyautogui.press('a')
        pyautogui.keyUp('ctrl')
        pyautogui.press('backspace', presses=1, interval=0.1)
        pyautogui.typewrite("5", interval=0.2)

        # dar um tab
        pyautogui.press('tab', presses=1, interval=0.5)
        pyautogui.press('backspace', presses=1, interval=0.1)
        
        pyautogui.press('tab', presses=1, interval=0.5)
        pyautogui.press('backspace', presses=1, interval=0.1)
        pyautogui.typewrite("zzzzzzzzzzzzzzz", interval=0.2)

        if not clicar_imagem("data/botao_diretorio.png", confidence=0.8, timeout=15, descricao="Botão diretório"):
            print("Erro ao clicar no botão para selecionar diretório.")
            return

        if not clicar_imagem("data/nome_diretorio.png", confidence=0.8, timeout=15, descricao="Botão nome diretorio"):
            print("Erro ao clicar no botão para selecionar nome diretório.")
            return
        
        pyautogui.keyDown('ctrl')
        pyautogui.press('a')
        pyautogui.keyUp('ctrl')
        pyautogui.press('backspace', presses=1, interval=0.1)

        # pyautogui.press('tab', presses=1, interval=0.5)

        # definir o nome do diretório
        data = datetime.strptime(competencia, "%d/%m/%Y")
        ano = data.year
        mes = data.month
        caminho_fixo = os.getenv("CAMINHO_FIXO_RAZAO")
        caminho_fixo_completo = f"{caminho_fixo}\\{ano}\\{mes}_{ano}"
        caminho_fixo_completo_com_filial = f"{caminho_fixo_completo}\\Razao_Filial_{filial}"

        # vefificar se não existe o diretório, se não existir criar
        if not os.path.exists(caminho_fixo_completo):
            os.makedirs(caminho_fixo_completo)
        
        print(f"📂 Caminho completo: {caminho_fixo_completo}")
        print(f"📝 Nome do arquivo: Razao_Filial_{filial}")

        # digita o caminho completo (pasta + nome do arquivo)
        pyautogui.press('backspace', presses=1, interval=0.1)
        pyautogui.typewrite(caminho_fixo_completo_com_filial, interval=0.1)

        # clicar em abrir
        time.sleep(2)   
        if not clicar_imagem("data/abrir_diretorio.png", confidence=0.9, timeout=15, descricao="Botão Abrir Diretório"):
            print("Erro ao clicar no botão Abrir Diretório.")
            return

        time.sleep(2)
        # clicar no botão "Salvar" da janela de salvar arquivo
        if not clicar_imagem("data/botao_salvar_arquivo_final.png", confidence=0.8, timeout=15, descricao="Botão Salvar Arquivo"):
            print("Erro ao clicar no botão Salvar Arquivo na janela de salvar.")
            return

        print("🔍 Aguardando conclusão do download...")

        time.sleep(8)
        print(f"Clicando em fechar...")
        if not clicar_imagem("data/botao_fechar.png", confidence=0.8, timeout=15, descricao="Botão Fechar"):
            print("Erro ao clicar no botão fechar.")
            return
        
        time.sleep(10)  # Aguarda um pouco antes de monitorar
        
        # Aguarda download e converte para Excel
        # Passa o diretório (sem nome do arquivo) e o nome esperado do arquivo
        sucesso, arquivo_baixado, tempo_gasto = aguardar_download_completo(
            diretorio_temp=caminho_fixo_completo,
            nome_arquivo_esperado=f"Razao_Filial_{filial}",
            timeout=100,  
            intervalo_verificacao=2
        )
        
        if not sucesso:
            print(f"❌ Erro: Download não concluído para a filial {filial}")
            continue
        
        print(f"⚡ Tempo total: {tempo_gasto:.1f} segundos")
        print(f"💾 Arquivo final: {arquivo_baixado}")
        
        print(f"✅ Filial {filial} processada com sucesso!")

    # clica no menu a opção "Relatórios", para fechar o menu aberto inicialmente
    if not clicar_imagem("data/menu_relatorios.png", confidence=0.8, timeout=15, descricao="Menu Relatórios"):
        print("Erro ao acessar o menu Relatórios.")
        exit(1)
    
    print("\n" + "="*60)
    print("✅ Automação do centro de custo concluída para todas as filiais!")
    print("="*60)