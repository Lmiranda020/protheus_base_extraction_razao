import time
import pyautogui
from modules.localizar_imagem import localizar_imagem, salvar_screenshot
from modules.clicar_imagem import afastar_mouse

IMAGENS_BOTAO_AGENT = ["data/botao_confirmar_agent_v2.png", "data/botao_confirmar_agent.png"]


def habilitar_app_agent(timeout=40, tentativas=3):
    """
    Clica em "Abrir Web Agent" no diálogo do navegador assim que ele aparecer
    e confirma que o diálogo fechou.

    O navegador ignora cliques feitos logo após o diálogo surgir (proteção contra
    cliques acidentais), por isso espera alguns segundos antes de clicar e,
    se o diálogo continuar aberto, tenta de novo.
    """
    localizacao = localizar_imagem(IMAGENS_BOTAO_AGENT, confidence=0.8,
                                   timeout=timeout, descricao="Botão app web agent")
    if not localizacao:
        print("Erro ao clicar no botão WebAgent.")
        return False

    for tentativa in range(1, tentativas + 1):
        # Dá tempo para o diálogo passar a aceitar cliques
        time.sleep(3)

        pyautogui.moveTo(pyautogui.center(localizacao), duration=0.3)
        pyautogui.click()
        print(f"✓ Clicou em: Botão app web agent (tentativa {tentativa})")
        afastar_mouse()

        time.sleep(3)
        localizacao = localizar_imagem(IMAGENS_BOTAO_AGENT, confidence=0.8, timeout=2,
                                       descricao="Botão app web agent (ainda aberto?)",
                                       screenshot_se_falhar=False)
        if not localizacao:
            print("✅ Diálogo do WebAgent fechado")
            time.sleep(5)
            return True

        print("⚠️ Diálogo do WebAgent continua aberto — tentando novamente.")

    # Último recurso: o foco do diálogo fica em "Cancelar"; Shift+Tab vai para
    # "Abrir Web Agent" e Enter confirma
    print("⌨️ Tentando abrir o WebAgent pelo teclado...")
    pyautogui.hotkey('shift', 'tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(3)

    if localizar_imagem(IMAGENS_BOTAO_AGENT, confidence=0.8, timeout=2,
                        descricao="Botão app web agent (ainda aberto?)",
                        screenshot_se_falhar=False):
        print("❌ Não foi possível fechar o diálogo do WebAgent.")
        salvar_screenshot("falha_webagent")
        return False

    print("✅ Diálogo do WebAgent fechado (teclado)")
    time.sleep(5)
    return True
