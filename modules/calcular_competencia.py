import os
from datetime import datetime, timedelta

def calcular_competencia():
    '''
        Função para calcular a competência (último dia do mês anterior) a partir da data atual.

        Para reprocessar um mês específico, defina COMPETENCIA no .env
        (formato dd/mm/aaaa, ex.: COMPETENCIA=31/03/2026). Deixe vazio para o cálculo automático.
    '''
    competencia_env = (os.getenv("COMPETENCIA") or "").strip()
    if competencia_env:
        datetime.strptime(competencia_env, "%d/%m/%Y")  # valida o formato
        print(f"✓ Competência (definida no .env): {competencia_env}")
        return competencia_env

    hoje = datetime.now()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    competencia = ultimo_dia_mes_anterior.strftime("%d/%m/%Y")
    print(f"✓ Competência: {competencia}")
    return competencia