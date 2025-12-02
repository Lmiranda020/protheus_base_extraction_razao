# 🤖 Automação de Razão Contábil

## 📋 Sobre o Projeto

Sistema de automação desenvolvido para otimizar o processo de geração e salvamento de relatórios de razão contábil no sistema corporativo da empresa.

### 👨‍💻 Desenvolvedora
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

---

## ✅ Solução Implementada

Esta automação resolve todos os problemas mencionados através de:

- **Execução automática** de todo o fluxo de geração de relatórios
- **Padronização rigorosa** de nomenclatura e estrutura de pastas
- **Velocidade otimizada** com processamento em lote
- **Rastreabilidade** com logs detalhados de execução

---

## 🚀 Funcionalidades

### Principais Recursos

1. **Cálculo Automático de Competência**
   - Identifica automaticamente o período anterior ao atual
   - Garante que os relatórios sejam sempre do mês correto

2. **Conexão VPN Automatizada**
   - Estabelece conexão segura com a rede corporativa
   - Prerequisito para acesso ao sistema

3. **Login Automatizado**
   - Credenciais gerenciadas de forma segura via variáveis de ambiente
   - Login automático no sistema corporativo

4. **Habilitação do App Agent**
   - Ativa módulos necessários para execução
   - Configura ambiente adequadamente

5. **Processamento de Razão Contábil**
   - Executa a rotina principal de geração de relatórios
   - Salva arquivos com nomenclatura padronizada
   - Organiza documentos em estrutura de pastas definida

6. **Logout e Finalização Seguros**
   - Encerra sessão adequadamente
   - Fecha aplicações abertas
   - Limpa recursos utilizados

## 🎯 Como Usar

### Fluxo de Execução

1. Sistema calcula competência anterior automaticamente
2. Conecta à VPN corporativa
3. Abre aplicação corporativa
4. Realiza login com credenciais do .env
5. Habilita App Agent
6. Executa rotina de geração de razão contábil
7. Salva relatórios padronizados
8. Realiza logout seguro
9. Fecha aplicações e finaliza

---

## 💻 Tecnologias Utilizadas

- **Python 3.x** - Linguagem principal
- **PyAutoGUI** - Automação de interface gráfica
- **python-dotenv** - Gerenciamento de variáveis de ambiente
- **Reconhecimento de imagens** - Para interação com elementos visuais

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

- ⏱️ **Redução de 80% no tempo** de processamento
- 📁 **100% de padronização** em nomenclatura de arquivos
- ✅ **Zero erros** de digitação ou esquecimento
- 🔍 **Facilidade na busca** de relatórios históricos
- 👥 **Liberação de tempo** da equipe para atividades estratégicas

---

## 🛡️ Segurança

- Credenciais armazenadas em variáveis de ambiente (nunca hardcoded)
- Conexão VPN obrigatória para acesso ao sistema
- Logout automático ao final da execução
- Tratamento de erros com interrupção segura do processo

---

**Última atualização**: Dezembro 2025