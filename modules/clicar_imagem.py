import pyautogui
from modules.localizar_imagem import localizar_imagem


def afastar_mouse():
    """
    Move o mouse para a borda esquerda da tela, para que o efeito de hover
    não altere a aparência do botão e atrapalhe a próxima busca.
    (Não usa os cantos, que disparam o FAILSAFE do pyautogui.)
    """
    _, altura = pyautogui.size()
    pyautogui.moveTo(5, altura // 2)


def clicar_imagem(caminho_imagem, confidence=0.8, timeout=10, descricao="",
                  grayscale=True, region=None, afastar_mouse_apos=True):
    """
    Procura uma imagem na tela e clica no centro dela

    Args:
        caminho_imagem: caminho da imagem ou lista de caminhos alternativos
        confidence: precisão da busca (0.0 a 1.0)
        timeout: tempo máximo de espera em segundos
        descricao: descrição do botão para logs
        grayscale: busca em tons de cinza (mais rápida)
        region: (left, top, width, height) para limitar a área de busca
        afastar_mouse_apos: tira o mouse de cima do botão após o clique

    Returns:
        True se encontrou e clicou, False caso contrário
    """
    localizacao = localizar_imagem(
        caminho_imagem, confidence=confidence, timeout=timeout,
        descricao=descricao, grayscale=grayscale, region=region,
    )
    if not localizacao:
        return False

    pyautogui.click(pyautogui.center(localizacao))
    print(f"✓ Clicou em: {descricao if descricao else caminho_imagem}")

    if afastar_mouse_apos:
        afastar_mouse()

    return True
