# 🤖 Automação de Razão Contábil

## 📋 Sobre o Projeto

Sistema de automação desenvolvido para otimizar o processo de geração e salvamento de relatórios de razão contábil no sistema corporativo da empresa.

### 👨‍💻 Desenvolvedor
**Larissa Miranda**

---

## ❌ Problemas Identificados

Antes da implementação desta automação, o processo manual apresentava diversos gargalos operacionais:

### 1. **Lentidão no Salvamento**
- Cada relatório demorava vários minutos para ser gerado e salvo manualmente
- Processo repetitivo que consumia horas de trabalho produtivo da equipe
- Necessidade de acompanhamento constante durante a execução

### 2. **Falta de Padronização**
- Nomenclatura inconsistente dos arquivos salvos
- Diferentes colaboradores utilizavam padrões distintos de nomenclatura
- Dificuldade em localizar relatórios específicos posteriormente

### 3. **Desorganização de Pastas**
- Arquivos salvos em locais diversos sem estrutura definida
- Ausência de hierarquia clara de diretórios
- Perda de tempo procurando documentos em múltiplas pastas

### 4. **Erros Humanos**
- Esquecimento de salvar determinados relatórios
- Erros de digitação em nomes de arquivos
- Inconsistência nos períodos de competência selecionados

### 5. **Ausência de Rastreabilidade**
- Nenhum registro histórico de execuções anteriores
- Impossibilidade de identificar quais filiais foram processadas com sucesso
- Sem visibilidade sobre falhas ou inconsistências no processo


---

## ✅ Solução Implementada

Esta automação resolve todos os problemas mencionados através de:

- **Execução automática** de todo o fluxo de geração de relatórios
- **Padronização rigorosa** de nomenclatura e estrutura de pastas
- **Velocidade otimizada** com processamento em lote
- **Log incremental em Excel** com histórico completo de todas as execuções
- **Notificação automática por e-mail** ao final de cada execução

---

## 🚀 Funcionalidades

### Principais Recursos

1. **Cálculo Automático de Competência**
   - Identifica automaticamente o período anterior ao atual
   - Garante que os relatórios sejam sempre do mês correto

2. **Login Automatizado**
   - Credenciais gerenciadas de forma segura via variáveis de ambiente
   - Login automático no sistema corporativo

3. **Habilitação do App Agent**
   - Ativa módulos necessários para execução
   - Configura ambiente adequadamente

4. **Processamento de Razão Contábil**
   - Executa a rotina principal de geração de relatórios por filial
   - Salva arquivos com nomenclatura padronizada (`Razao_Filial_XXXX`)
   - Organiza documentos em estrutura de pastas definida (`CAMINHO\ANO\MM_AAAA\`)
   - Em caso de falha em uma filial, continua o processamento das demais

5. **Conversão Automática de Arquivos**
   - Converte arquivos XML (formato original de download) para XLSX
   - Remove arquivos XML após conversão bem-sucedida
   - Mantém apenas os arquivos Excel finalizados

6. **Log de Execução em Excel**
   - Registra automaticamente cada filial processada com status, tempo e mensagem
   - Salvo de forma **incremental** em `log/log_automacao.xlsx` — cada execução adiciona novas linhas sem apagar o histórico anterior
   - Colunas registradas: Data/Hora Início, Data/Hora Fim, Tipo, Competência, Filial, Status, Mensagem, Tempo (s)
   - Formatação visual com cores por status (✅ verde / ❌ vermelho) e zebra nas linhas

7. **Notificação Automática por E-mail**
   - Ao final da execução, envia e-mail automático para a equipe de custo
   - Corpo do e-mail em HTML com resumo visual: total de filiais, processadas com sucesso, com erro e taxa de sucesso
   - Lista detalhada das filiais OK e das filiais com erro
   - Log Excel anexado ao e-mail para consulta imediata
   - Suporta múltiplos destinatários (separados por vírgula no `.env`)

8. **Logout e Finalização Seguros**
   - Encerra sessão adequadamente
   - Fecha aplicações abertas
   - Limpa recursos utilizados

---

## 💻 Tecnologias Utilizadas

- **Python** - Linguagem principal
- **PyAutoGUI** - Automação de interface gráfica
- **python-dotenv** - Gerenciamento de variáveis de ambiente
- **openpyxl** - Geração e atualização do log em Excel
- **smtplib / email** - Envio de e-mail via Gmail (App Password)
- **Reconhecimento de imagens** - Para interação com elementos visuais

---


## 📁 Estrutura de Pastas Gerada

```
projeto/
├── log/
│   └── log_automacao.xlsx      ← log incremental de todas as execuções
├── data/                       ← imagens de referência para reconhecimento
├── modules/                    ← módulos da automação
├── config/                     ← configurações (lista de filiais, etc.)
├── main.py
└── .env
```

---

## ⏰ Agendamento Automático

A automação foi configurada para execução automática através do **Agendador de Tarefas do Windows**, programada para rodar em um dia específico do mês.

### Arquitetura de Execução

Para garantir a execução completa sem interrupções:

1. **Arquivo BAT** - Script batch que inicia a automação Python
2. **Agendador de Tarefas** - Executa o arquivo BAT no momento programado

Essa abordagem evita que o Agendador de Tarefas encerre a automação prematuramente em caso de processos mais demorados, garantindo que toda a rotina seja concluída adequadamente.

---

## 📊 Benefícios Mensurados

### Ganhos Operacionais

- ⏱️ **Redução de 100% no tempo** de processamento manual
- 📁 **100% de padronização** em nomenclatura de arquivos
- ✅ **Zero erros** de digitação ou esquecimento
- 📊 **Arquivos entregues em formato Excel** prontos para uso (conversão automática de XML)
- 🔍 **Facilidade na busca** de relatórios históricos
- 📋 **Rastreabilidade completa** com log incremental por filial e por execução
- 📧 **Notificação imediata** da equipe de custo ao término da extração
- 👥 **Liberação de tempo** da equipe para atividades estratégicas

---

**Última atualização**: Maio 2026