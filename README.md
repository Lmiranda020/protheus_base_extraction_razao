# 🤖 Automação Inteligente de Relatórios Gerenciais

> Solução desenvolvida para automatizar extração de relatórios de Centro de Custo e Consumo, reduzindo drasticamente o tempo de processamento e eliminando erros manuais.

## 💡 O Problema

Antes da automação, o processo manual de extração de relatórios era:
- ⏰ **Demorado**: ~3-4 horas por mês para processar todas as filiais
- 😫 **Repetitivo**: Mesmos cliques e preenchimentos dezenas de vezes
- ❌ **Propenso a erros**: Nomenclatura inconsistente e arquivos na pasta errada
- 👤 **Dependente**: Precisava de uma pessoa dedicada ao processo

## ✨ A Solução

Automação inteligente desenvolvida em Python que:
- 🚀 Processa múltiplas filiais automaticamente
- 🎯 Aplica filtros e configurações de forma precisa
- 📁 Organiza arquivos automaticamente (ano/mês)
- ⚡ **Sistema de download inteligente** que detecta quando terminou

## 📊 Resultados e Ganhos

### Tempo de Processamento

| Tarefa | Antes (Manual) | Depois (Automatizado) | Economia |
|--------|----------------|----------------------|----------|
| Centro de Custo (10 filiais) | ~60 min | **~8 min** | **87% mais rápido** |
| Consumo (10 filiais) | ~90 min | **~12 min** | **87% mais rápido** |
| **Total Mensal** | **~150 min** | **~20 min** | **💰 130 minutos economizados** |

### Impacto Anual
- ⏱️ **26 horas** de trabalho manual eliminadas por ano
- 💰 Equivalente a **3+ dias úteis** de produtividade recuperada
- 🎯 **Zero erros** de nomenclatura ou organização de arquivos
- 📈 Escalável para novas filiais sem custo adicional

### Diferencial Técnico: Download Inteligente

**Desafio:** Downloads de relatórios têm tempos variáveis e imprevisíveis.

**Solução implementada:**
Sistema de monitoramento inteligente que detecta automaticamente quando o download é concluído:

1. 🔍 Monitora criação do arquivo em tempo real
2. 📊 Detecta abertura automática do Excel
3. ✅ Confirma conclusão e prossegue imediatamente
4. 🔒 Fecha Excel automaticamente

**Resultado:** Não desperdiça tempo esperando. Segue para próxima filial assim que o arquivo está pronto!

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+** - Linguagem principal
- **PyAutoGUI** - Automação de interface
- **PSUtil** - Monitoramento de processos
- **Pillow** - Reconhecimento de imagens
- **Python-dotenv** - Gerenciamento de configurações

## 🚀 Como Funciona (Simplificado)

### Centro de Custo
```
1. Acessa sistema → Menu Relatórios
2. Para cada filial:
   ✓ Preenche dados automaticamente
   ✓ Configura exportação XLSX
   ✓ Nomeia: CC_{FILIAL}_{DATA}
   ✓ Monitora download inteligentemente
   ✓ Organiza em: /Ano/Mês_Ano/
3. Pronto! ✅
```

### Consumo (SD3)
```
1. Acessa sistema → Menu Consultas
2. Para cada filial:
   ✓ Aplica filtros de competência
   ✓ Configura dicionário de dados
   ✓ Exporta CSV delimitado
   ✓ Monitora download inteligentemente
   ✓ Renomeia e organiza automaticamente
3. Pronto! ✅
```

## 📦 Estrutura do Projeto

```
automacao-relatorios/
├── modules/
│   ├── aguardar_download_inteligente.py  # ⚡ Sistema inteligente
│   ├── clicar_imagem.py                  # Reconhecimento visual
│   └── mover_e_renomear_arquivo_baixado.py
├── config/
│   └── list_filial.py                    # Lista de filiais
├── data/                                 # Imagens de referência
├── automacao_centro_de_custo.py          # 📊 Script principal
├── automacao_consumo.py                  # 📈 Script principal
└── .env                                  # Configurações
```

## ⚙️ Setup Rápido

1. **Clone e instale dependências**
```bash
git clone <repo>
cd automacao-relatorios
pip install -r requirements.txt
```

2. **Configure o `.env`**
```env
CAMINHO_FIXO_CC=C:\Relatorios\CentroCusto
CAMINHO_FIXO_CONSUMO=C:\Relatorios\Consumo
DIRETORIO_TEMP=C:\Users\SeuUsuario\Downloads
```

3. **Execute**
```python
from automacao_centro_de_custo import automacao_centro_de_custo
from automacao_consumo import automacao_consumo

competencia = "31/10/2024"
automacao_centro_de_custo(competencia)
automacao_consumo(competencia)
```

## ⏰ Execução Automática

A automação está configurada no **Agendador de Tarefas do Windows** para executar automaticamente toda **primeira segunda-feira do mês**, garantindo que os relatórios estejam sempre prontos no início do período.

**Configuração:**
- 📅 Trigger: Primeira segunda-feira de cada mês
- 🕐 Horário: Definido para período de baixo uso do sistema
- 🔄 Execução: Totalmente desassistida
- ✅ Resultado: Relatórios prontos sem intervenção manual

## 💼 Valor para o Negócio

✅ **Redução de custos operacionais**
✅ **Aumento de produtividade da equipe**
✅ **Eliminação de erros humanos**
✅ **Processo padronizado e auditável**
✅ **Escalabilidade sem custo adicional**


## 👥 Autores

- **Desenvolvedor Principal** - *Larissa Miranda*

  ---
