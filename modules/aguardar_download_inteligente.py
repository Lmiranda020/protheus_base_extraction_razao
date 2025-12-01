import os
import time
import psutil
from pathlib import Path
import win32com.client
import pythoncom

def verificar_excel_aberto():
    """Verifica se há alguma instância do Excel aberta"""
    try:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                return True
        return False
    except Exception as e:
        print(f"Erro ao verificar Excel: {e}")
        return False

def tem_extensao(arquivo):
    """Verifica se um arquivo tem extensão"""
    nome = os.path.basename(arquivo)
    return '.' in nome and len(os.path.splitext(nome)[1]) > 0

def converter_para_excel_com_win32(arquivo_origem, nome_arquivo_destino, diretorio_destino):
    """
    Converte arquivo XML/sem extensão para Excel usando win32com (Excel nativo)
    
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
        
        # Abre o Excel em modo invisível
        excel = win32com.client.Dispatch("Excel.Application")
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

def aguardar_download_completo(diretorio_temp, nome_arquivo_esperado, timeout=200, intervalo_verificacao=2):
    """
    Aguarda o download ser concluído:
    1. Detecta arquivo específico (sem extensão)
    2. Aguarda estabilização
    3. Fecha Excel se estiver aberto
    4. Converte o arquivo para .xlsx usando Excel nativo
    
    Args:
        diretorio_temp: diretório onde os arquivos são baixados
        nome_arquivo_esperado: nome base do arquivo esperado (ex: "Razao_Filial_01010009")
        timeout: tempo máximo de espera em segundos (padrão: 200s)
        intervalo_verificacao: intervalo entre verificações em segundos (padrão: 2s)
    
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
    
    # Fase 1: Procura o arquivo específico
    print(f"📥 Procurando arquivo '{nome_arquivo_esperado}' (com ou sem extensão)...")
    
    while tempo_decorrido < timeout:
        try:
            # Lista TODOS os arquivos do diretório
            arquivos_agora = os.listdir(diretorio_temp)
            
            # Procura pelo arquivo esperado
            for nome_arq in arquivos_agora:
                caminho_completo = os.path.join(diretorio_temp, nome_arq)
                
                # Pula se for diretório
                if os.path.isdir(caminho_completo):
                    continue
                
                # Remove extensão se tiver
                nome_sem_ext = os.path.splitext(nome_arq)[0]
                
                # Verifica se é o arquivo esperado (CASE-INSENSITIVE)
                if (nome_sem_ext.lower() == nome_arquivo_esperado.lower() or 
                    nome_arq.lower() == nome_arquivo_esperado.lower()):
                    
                    arquivo_novo = caminho_completo
                    print(f"✅ Arquivo encontrado: {nome_arq}")
                    
                    # Verifica se tem extensão
                    if not tem_extensao(arquivo_novo):
                        print("📝 Arquivo sem extensão confirmado")
                    else:
                        print(f"📝 Arquivo com extensão: {os.path.splitext(nome_arq)[1]}")
                    
                    break
            
            if arquivo_novo:
                break
                
        except Exception as e:
            print(f"⚠️ Erro ao verificar arquivos: {e}")
        
        time.sleep(intervalo_verificacao)
        tempo_decorrido = time.time() - tempo_inicio
        
        if int(tempo_decorrido) % 10 == 0 and tempo_decorrido > 0:
            print(f"⏱️  Procurando arquivo... {int(tempo_decorrido)}s / {timeout}s")
    
    if not arquivo_novo:
        print(f"❌ Arquivo '{nome_arquivo_esperado}' não foi encontrado")
        return False, None, time.time() - tempo_inicio
    
    # Fase 2: Aguarda estabilização do arquivo
    print("⏳ Aguardando estabilização do arquivo...")
    time.sleep(3)
    
    # Fase 3: Fecha Excel se estiver aberto
    if verificar_excel_aberto():
        print("🔄 Fechando Excel aberto em background...")
        fechar_excel()
        time.sleep(3)
    
    # Fase 4: Converte o arquivo usando Excel nativo
    nome_base_original = os.path.splitext(os.path.basename(arquivo_novo))[0]
    nome_arquivo_final = nome_base_original
    
    print(f"💾 Convertendo '{nome_base_original}' para '{nome_arquivo_final}.xlsx'")
    
    sucesso, arquivo_xlsx = converter_para_excel_com_win32(
        arquivo_origem=arquivo_novo,
        nome_arquivo_destino=nome_arquivo_final,
        diretorio_destino=diretorio_temp
    )
    
    if sucesso:
        # Remove o arquivo original após conversão
        try:
            os.remove(arquivo_novo)
            print(f"🗑️ Arquivo original removido: {os.path.basename(arquivo_novo)}")
        except Exception as e:
            print(f"⚠️ Não foi possível remover arquivo original: {e}")
        
        tempo_decorrido = time.time() - tempo_inicio
        print(f"✅ Processo concluído em {tempo_decorrido:.1f} segundos!")
        return True, arquivo_xlsx, tempo_decorrido
    else:
        print("❌ Falha na conversão do arquivo")
        tempo_decorrido = time.time() - tempo_inicio
        return False, None, tempo_decorrido


def fechar_excel():
    """Fecha todas as instâncias do Excel abertas"""
    try:
        fechou_algum = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                print(f"🔄 Fechando Excel (PID: {proc.pid})...")
                proc.kill()
                fechou_algum = True
        
        if fechou_algum:
            print("📊 Excel fechado com sucesso")
            time.sleep(4)
            return True
        else:
            print("ℹ️ Nenhuma instância do Excel estava aberta")
            return True
    except Exception as e:
        print(f"⚠️ Erro ao fechar Excel: {e}")
        return False