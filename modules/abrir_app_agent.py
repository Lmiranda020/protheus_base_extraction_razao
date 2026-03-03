import selenium
import time
from modules.clicar_imagem import clicar_imagem

def habilitar_app_agent():
    time.sleep(2)

    if not clicar_imagem("data/botao_confirmar_agent_v2.png", confidence=0.8, timeout=15, descricao="Botão app web agent"):
        time.sleep(2)
        # se não encontrou o primeiro, tenta o segundo modelo
        if not clicar_imagem("data/botao_confirmar_agent.png", confidence=0.8, timeout=15, descricao="Botão app web agent v2"):
            print("Erro ao clicar no botão WebAgent.")
        
    time.sleep(8)