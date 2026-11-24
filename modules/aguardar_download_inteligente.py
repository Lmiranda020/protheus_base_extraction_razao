import os
import time
import psutil
import win32com.client
import win32gui
import win32process
import pythoncom


def listar_pids_excel() -> set:
    """Retorna o conjunto de PIDs dos processos do Excel em execução."""
    pids = set()
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                pids.add(proc.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def pids_excel_com_arquivo(nome_arquivo, pids_candidatos) -> set:
    """
    Retorna os PIDs (dentre 'pids_candidatos') que têm uma janela visível do Excel
    cujo título contém 'nome_arquivo' — ou seja, o Excel que está com aquele arquivo aberto.
    """
    if not pids_candidatos:
        return set()

    nome = nome_arquivo.lower()
    encontrados = set()

    def _verificar_janela(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and nome in win32gui.GetWindowText(hwnd).lower():
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid in pids_candidatos:
                    encontrados.add(pid)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_verificar_janela, None)
    return encontrados


def verificar_e_aguardar_excel_aberto(pids_ignorar=None, nome_arquivo=None, timeout=320, intervalo=3):
    """
    Aguarda o Excel aberto pelo download (ignorando os PIDs em 'pids_ignorar', que já
    estavam abertos antes) por até 'timeout' segundos.

    Se 'nome_arquivo' for informado, só considera aberto quando houver uma janela do
    Excel com esse arquivo no título. Isso evita fechar o Excel antes de ele carregar
    o arquivo (arquivos grandes demoram a abrir).

    Returns:
        set com os PIDs a fechar (vazio se nenhum Excel novo apareceu).
    """
    pids_ignorar = pids_ignorar or set()
    alvo = f"'{nome_arquivo}'" if nome_arquivo else "um Excel novo"
    print(f"⏳ Aguardando o Excel abrir {alvo} (timeout: {timeout}s)...")
    tempo_inicio = time.time()
    novos = set()

    while (time.time() - tempo_inicio) < timeout:
        try:
            novos = listar_pids_excel() - pids_ignorar
            if novos and not nome_arquivo:
                print(f"✅ Excel detectado após {(time.time() - tempo_inicio):.1f}s")
                # aguarda um pouco para garantir que o Excel tenha aberto completamente
                time.sleep(10)
                return novos

            com_arquivo = pids_excel_com_arquivo(nome_arquivo, novos) if nome_arquivo else set()
            if com_arquivo:
                print(f"✅ Excel com o arquivo aberto detectado após {(time.time() - tempo_inicio):.1f}s")
                time.sleep(3)
                return com_arquivo
        except Exception as e:
            print(f"⚠️ Erro ao verificar Excel: {e}")

        time.sleep(intervalo)
        print(f"🔍 Excel ainda não abriu o arquivo... {int(time.time() - tempo_inicio)}s / {timeout}s")

    if novos:
        # Abriu Excel novo, mas a janela com o arquivo não foi identificada: fecha os novos
        print(f"⚠️ Janela do arquivo não identificada após {timeout}s — usando os Excel novos")
        return novos

    print(f"⚠️ Excel não foi detectado após {timeout}s")
    return set()


def tem_extensao(arquivo):
    """Verifica se um arquivo tem extensão"""
    nome = os.path.basename(arquivo)
    return '.' in nome and len(os.path.splitext(nome)[1]) > 0


def xml_excel_completo(caminho) -> bool:
    """
    Para o XML gerado pelo Protheus (arquivo sem extensão), confere se o arquivo já
    termina com a tag de fechamento </Workbook>, ou seja, se a gravação terminou.
    Arquivos com extensão não são verificados (retorna True).
    """
    if tem_extensao(caminho):
        return True
    try:
        with open(caminho, "rb") as arquivo:
            arquivo.seek(0, os.SEEK_END)
            arquivo.seek(max(0, arquivo.tell() - 4096))
            return b"</Workbook>" in arquivo.read()
    except OSError:
        return False


def aguardar_arquivo_estavel(caminho, timeout=900, intervalo=3, leituras_estaveis=5) -> bool:
    """
    Aguarda a gravação do arquivo terminar: o tamanho precisa parar de mudar por
    'leituras_estaveis' leituras seguidas e, no caso do XML do Protheus, o arquivo
    precisa terminar com </Workbook>. Arquivos grandes levam vários minutos.

    Returns:
        True se o arquivo ficou estável e completo dentro do 'timeout'.
    """
    tempo_inicio = time.time()
    tamanho_anterior = -1
    estaveis = 0
    ultimo_aviso = 0

    while time.time() - tempo_inicio < timeout:
        try:
            tamanho = os.path.getsize(caminho)
        except OSError:
            tamanho = -1

        if tamanho > 0 and tamanho == tamanho_anterior:
            estaveis += 1
            if estaveis >= leituras_estaveis:
                if xml_excel_completo(caminho):
                    print(f"✅ Arquivo estável e completo ({tamanho} bytes)")
                    return True
                # Tamanho parado, mas o XML ainda não foi fechado: o Protheus segue gravando
                estaveis = 0
        else:
            estaveis = 0

        decorrido = time.time() - tempo_inicio
        if decorrido - ultimo_aviso >= 30:
            ultimo_aviso = decorrido
            print(f"⏱️  Aguardando gravação do arquivo... {int(decorrido)}s / {timeout}s "
                  f"({max(tamanho, 0) / 1024 / 1024:.1f} MB)")

        tamanho_anterior = tamanho
        time.sleep(intervalo)

    print(f"⚠️ Arquivo não estabilizou em {timeout}s — seguindo mesmo assim")
    return False


def remover_arquivo_original(caminho, pids_ignorar=None, tentativas=3) -> bool:
    """
    Apaga o arquivo original após a conversão. Se ele estiver aberto em um Excel
    que abriu depois do download, fecha esse Excel e tenta de novo.
    """
    pids_ignorar = pids_ignorar or set()
    nome = os.path.basename(caminho)

    for tentativa in range(1, tentativas + 1):
        try:
            os.remove(caminho)
            print(f"🗑️ Arquivo original removido: {nome}")
            return True
        except FileNotFoundError:
            return True
        except OSError as e:
            print(f"⚠️ Não foi possível remover o arquivo original ({tentativa}/{tentativas}): {e}")
            if tentativa == tentativas:
                break
            novos = listar_pids_excel() - pids_ignorar
            alvo = pids_excel_com_arquivo(os.path.splitext(nome)[0], novos) or novos
            if alvo:
                print("🔄 Fechando Excel que está com o arquivo original aberto...")
                fechar_excel(pids=alvo)
            time.sleep(5)

    return False


def converter_para_excel_com_win32(arquivo_origem, nome_arquivo_destino, diretorio_destino):
    """
    Converte arquivo XML/sem extensão para Excel usando win32com (Excel nativo)

    Usa uma instância exclusiva do Excel (DispatchEx), para não interferir em
    planilhas que outra pessoa esteja usando na máquina.

    Args:
        arquivo_origem: caminho do arquivo original
        nome_arquivo_destino: nome desejado (sem extensão)
        diretorio_destino: diretório onde salvar o .xlsx

    Returns:
        tuple: (sucesso, caminho_arquivo_xlsx)
    """
    excel = None
    wb = None

    try:
        print(f"📄 Arquivo original: {os.path.basename(arquivo_origem)}")

        # Verifica se o arquivo tem extensão
        if not tem_extensao(arquivo_origem):
            print("📝 Arquivo sem extensão detectado (provavelmente XML do Excel)")

        print(f"🔄 Abrindo arquivo no Excel...")

        # Inicializa COM
        pythoncom.CoInitialize()

        # Abre uma instância nova e invisível do Excel
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Caminho absoluto do arquivo
        arquivo_abs = os.path.abspath(arquivo_origem)

        # Abre o arquivo (o Excel detecta automaticamente o formato XML)
        wb = excel.Workbooks.Open(arquivo_abs)
        print(f"✅ Arquivo aberto no Excel")

        # Caminho do arquivo Excel final
        arquivo_xlsx = os.path.join(diretorio_destino, f"{nome_arquivo_destino}.xlsx")
        arquivo_xlsx_abs = os.path.abspath(arquivo_xlsx)

        # Salva como XLSX (formato 51 = xlOpenXMLWorkbook)
        print(f"💾 Salvando como: {os.path.basename(arquivo_xlsx)}")
        wb.SaveAs(arquivo_xlsx_abs, FileFormat=51)

        # Informações sobre o arquivo
        num_sheets = wb.Worksheets.Count
        if num_sheets > 0:
            primeira_sheet = wb.Worksheets(1)
            linhas = primeira_sheet.UsedRange.Rows.Count
            colunas = primeira_sheet.UsedRange.Columns.Count
            print(f"📊 Dados: {linhas} linhas x {colunas} colunas")

        print(f"✅ Arquivo convertido com sucesso!")

        return True, arquivo_xlsx

    except Exception as e:
        print(f"❌ Erro ao converter arquivo: {e}")
        import traceback
        traceback.print_exc()
        return False, None

    finally:
        # Fecha o workbook
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass

        # Fecha o Excel
        try:
            if excel:
                excel.Quit()
        except:
            pass

        # Libera objetos COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

        # Aguarda um pouco
        time.sleep(2)


def aguardar_download_completo(diretorio_temp, nome_arquivo_esperado, inicio_download=None,
                               pids_excel_antes=None, timeout=200, intervalo_verificacao=2,
                               timeout_excel=320):
    """
    Aguarda o download ser concluído:
    1. Detecta o arquivo esperado, gerado/alterado depois de 'inicio_download'
    2. Aguarda o tamanho do arquivo estabilizar
    3. Fecha o Excel aberto pelo download (somente os processos novos)
    4. Converte o arquivo para .xlsx usando Excel nativo

    Args:
        diretorio_temp: diretório onde os arquivos são baixados
        nome_arquivo_esperado: nome base do arquivo esperado (ex: "Razao_Filial_01010009")
        inicio_download: timestamp (time.time()) de quando o download foi disparado.
                         Arquivos mais antigos (de execuções anteriores) são ignorados.
        pids_excel_antes: PIDs do Excel que já estavam abertos antes do download
                          (não serão fechados)
        timeout: tempo máximo de espera pelo arquivo em segundos (padrão: 200s)
        intervalo_verificacao: intervalo entre verificações em segundos (padrão: 2s)
        timeout_excel: tempo máximo de espera pela abertura do Excel (padrão: 320s)

    Returns:
        tuple: (sucesso, arquivo_path, tempo_decorrido)
    """
    print("🔍 Monitorando download...")
    print(f"📝 Procurando arquivo: {nome_arquivo_esperado}")
    print(f"📂 No diretório: {diretorio_temp}")

    # Garante que o diretório existe
    if not os.path.exists(diretorio_temp):
        print(f"❌ Diretório não existe: {diretorio_temp}")
        return False, None, 0

    tempo_inicio = time.time()
    tempo_decorrido = 0
    arquivo_novo = None
    # margem de 5s para diferenças de arredondamento do sistema de arquivos
    limite_mtime = (inicio_download - 5) if inicio_download else None
    ultimo_aviso = 0

    # Fase 1: Procura o arquivo específico
    print(f"📥 Procurando arquivo '{nome_arquivo_esperado}' (com ou sem extensão)...")

    while tempo_decorrido < timeout:
        try:
            candidatos = []
            for nome_arq in os.listdir(diretorio_temp):
                caminho_completo = os.path.join(diretorio_temp, nome_arq)

                # Pula se for diretório
                if os.path.isdir(caminho_completo):
                    continue

                # Remove extensão se tiver
                nome_sem_ext = os.path.splitext(nome_arq)[0]

                # Verifica se é o arquivo esperado (CASE-INSENSITIVE)
                if not (nome_sem_ext.lower() == nome_arquivo_esperado.lower() or
                        nome_arq.lower() == nome_arquivo_esperado.lower()):
                    continue

                # Ignora arquivos de execuções anteriores
                if limite_mtime and os.path.getmtime(caminho_completo) < limite_mtime:
                    continue

                candidatos.append(caminho_completo)

            if candidatos:
                # Prefere o arquivo sem extensão (formato original do Protheus)
                candidatos.sort(key=lambda c: tem_extensao(c))
                arquivo_novo = candidatos[0]
                print(f"✅ Arquivo encontrado: {os.path.basename(arquivo_novo)}")
                break

        except Exception as e:
            print(f"⚠️ Erro ao verificar arquivos: {e}")

        time.sleep(intervalo_verificacao)
        tempo_decorrido = time.time() - tempo_inicio

        if tempo_decorrido - ultimo_aviso >= 30:
            ultimo_aviso = tempo_decorrido
            print(f"⏱️  Procurando arquivo... {int(tempo_decorrido)}s / {timeout}s")

    if not arquivo_novo:
        print(f"❌ Arquivo '{nome_arquivo_esperado}' não foi encontrado")
        return False, None, time.time() - tempo_inicio

    # Fase 2: Aguarda estabilização do arquivo
    print("⏳ Aguardando estabilização do arquivo...")
    aguardar_arquivo_estavel(arquivo_novo)

    # Fase 3: Aguarda o Excel do download abrir e depois fecha somente ele
    novos_pids = verificar_e_aguardar_excel_aberto(
        pids_ignorar=pids_excel_antes,
        nome_arquivo=os.path.splitext(os.path.basename(arquivo_novo))[0],
        timeout=timeout_excel,
        intervalo=3,
    )
    if novos_pids:
        print("🔄 Fechando Excel aberto pelo download...")
        fechar_excel(pids=novos_pids)
        time.sleep(6)  # aguarda o Windows liberar o handle do arquivo
    else:
        print("⚠️ Excel não detectado, seguindo para conversão mesmo assim...")

    # Fase 4: Converte o arquivo usando Excel nativo
    nome_base_original = os.path.splitext(os.path.basename(arquivo_novo))[0]
    nome_arquivo_final = nome_base_original
    arquivo_xlsx_destino = os.path.join(diretorio_temp, f"{nome_arquivo_final}.xlsx")

    if os.path.normcase(os.path.abspath(arquivo_novo)) == os.path.normcase(os.path.abspath(arquivo_xlsx_destino)):
        # O arquivo já veio como .xlsx — não converte nem apaga
        tempo_decorrido = time.time() - tempo_inicio
        print(f"✅ Arquivo já está em .xlsx, conversão não necessária ({tempo_decorrido:.1f}s)")
        return True, arquivo_novo, tempo_decorrido

    print(f"💾 Convertendo '{nome_base_original}' para '{nome_arquivo_final}.xlsx'")

    sucesso, arquivo_xlsx = converter_para_excel_com_win32(
        arquivo_origem=arquivo_novo,
        nome_arquivo_destino=nome_arquivo_final,
        diretorio_destino=diretorio_temp
    )

    if sucesso:
        # Remove o arquivo original após conversão (fecha o Excel que o prender, se houver)
        remover_arquivo_original(arquivo_novo, pids_ignorar=pids_excel_antes)

        tempo_decorrido = time.time() - tempo_inicio
        print(f"✅ Processo concluído em {tempo_decorrido:.1f} segundos!")
        return True, arquivo_xlsx, tempo_decorrido
    else:
        print("❌ Falha na conversão do arquivo")
        tempo_decorrido = time.time() - tempo_inicio
        return False, None, tempo_decorrido


def fechar_excel(pids=None):
    """
    Fecha as instâncias do Excel indicadas em 'pids'.
    Se 'pids' não for informado, fecha todas as instâncias abertas.
    """
    try:
        fechou_algum = False
        for proc in psutil.process_iter(['name']):
            try:
                if not (proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper()):
                    continue
                if pids is not None and proc.pid not in pids:
                    continue
                print(f"🔄 Fechando Excel (PID: {proc.pid})...")
                proc.kill()
                fechou_algum = True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if fechou_algum:
            print("📊 Excel fechado com sucesso")
            time.sleep(4)
        else:
            print("ℹ️ Nenhuma instância do Excel estava aberta")
        return True
    except Exception as e:
        print(f"⚠️ Erro ao fechar Excel: {e}")
        return False
