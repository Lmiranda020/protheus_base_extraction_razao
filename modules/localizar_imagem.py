import os
import time
from datetime import datetime
import pyautogui

PASTA_SCREENSHOTS = os.path.join("log", "screenshots")


def salvar_screenshot(motivo: str) -> str | None:
    """
    Salva um print da tela atual em log/screenshots/ para facilitar a análise de falhas.

    Returns:
        Caminho do arquivo salvo ou None se não conseguiu salvar.
    """
    try:
        os.makedirs(PASTA_SCREENSHOTS, exist_ok=True)
        nome_seguro = "".join(c if c.isalnum() or c in "-_" else "_" for c in motivo)[:80]
        caminho = os.path.join(
            PASTA_SCREENSHOTS,
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome_seguro}.png",
        )
        pyautogui.screenshot(caminho)
        print(f"📸 Screenshot salvo: {caminho}")
        return caminho
    except Exception as e:
        print(f"⚠️ Não foi possível salvar screenshot: {e}")
        return None


def localizar_imagem(caminho_imagem, confidence=0.8, timeout=10, descricao="",
                     grayscale=True, region=None, screenshot_se_falhar=True):
    """
    Procura uma imagem na tela, repetindo até encontrar ou esgotar o timeout.

    Args:
        caminho_imagem: caminho da imagem ou lista de caminhos alternativos
                        (ex.: botão normal e botão com hover). Retorna o primeiro encontrado.
        confidence: precisão da busca (0.0 a 1.0)
        timeout: tempo máximo de espera em segundos
        descricao: descrição da imagem para logs
        grayscale: busca em tons de cinza (mais rápida e tolerante a variações de cor)
        region: (left, top, width, height) para limitar a área de busca
        screenshot_se_falhar: salva um print da tela quando não encontra

    Returns:
        A localização (Box) se encontrou, None caso contrário
    """
    caminhos = [caminho_imagem] if isinstance(caminho_imagem, str) else list(caminho_imagem)
    nome = descricao if descricao else ", ".join(caminhos)

    existentes = [c for c in caminhos if os.path.exists(c)]
    for c in caminhos:
        if c not in existentes:
            print(f"⚠️ Imagem de referência não existe: {c}")
    if not existentes:
        print(f"✗ Nenhuma imagem de referência disponível para: {nome}")
        return None

    print(f"Procurando por: {nome}")

    tempo_inicial = time.time()

    while time.time() - tempo_inicial < timeout:
        for caminho in existentes:
            try:
                localizacao = pyautogui.locateOnScreen(
                    caminho, confidence=confidence, grayscale=grayscale, region=region
                )
                if localizacao:
                    print(f"✓ Imagem localizada: {nome}")
                    return localizacao
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                print(f"Erro ao procurar imagem {caminho}: {e}")

        time.sleep(0.5)

    print(f"✗ Imagem não encontrada: {nome}")
    if screenshot_se_falhar:
        salvar_screenshot(f"nao_encontrou_{nome}")
    return None


def aguardar_imagem(caminho_imagem, timeout=30, **kwargs) -> bool:
    """Aguarda uma imagem aparecer na tela. Retorna True se apareceu dentro do timeout."""
    return localizar_imagem(caminho_imagem, timeout=timeout, **kwargs) is not None
