# financial_operation_dashboard

# 💰 Controle Financeiro Pessoal - Excel VBA

> **Status:** 🚧 Em desenvolvimento ativo
> 
> **Versão atual:** v0.1.0 - Fluxo de Caixa concluído ✅ | Dashboard e Config em desenvolvimento 🚧

## 📋 Sobre o Projeto

Este projeto tem como objetivo criar uma planilha de **controle financeiro pessoal** utilizando Microsoft Excel combinado com Visual Basic for Applications (VBA), transformando a experiência tradicional de planilhas em uma **experiência de aplicativo** moderna e intuitiva.

O foco principal é demonstrar o **domínio completo das funcionalidades do Excel** e técnicas avançadas de **análise de dados**, provando que é possível criar soluções robustas e profissionais usando ferramentas que muitos subestimam.

## 🎯 Objetivos

- **Demonstrar habilidades avançadas em Excel:** Fórmulas complexas, automações VBA, validações de dados e formatação condicional
- **Criar experiência de usuário similar a um aplicativo:** Interface intuitiva, automações inteligentes e feedback visual em tempo real
- **Aplicar análise de dados:** Transformar dados brutos em insights visuais através de dashboards dinâmicos
- **Desenvolver solução prática:** Ferramenta real e funcional para gestão financeira pessoal

## 🏗️ Estrutura do Projeto

O projeto está organizado em três abas principais:

### 1️⃣ Fluxo de Caixa ✅ **[CONCLUÍDO]**
Aba principal para registro e acompanhamento de movimentações financeiras.

**Status:** Totalmente funcional e operacional

**Funcionalidades:**
- Registro de **despesas** com categorização por tópicos
- Registro de **ganhos** com controle de status
- **Automação de datas:** Preenchimento automático de datas de realização quando status muda para "Concluído"
- **Cálculo automático de saldo:** Atualização em tempo real baseado em ganhos e despesas
- **Sistema de investimentos:** Transferência automática entre Cofrinho e Investimentos com validação de saldo
- **Carteira digital:** Visualização consolidada de Saldo, Valor Investido e Cofrinho

**Campos principais:**
- Data Prevista
- Data de Realização (automática)
- Descrição
- Tópico (Gasto Fixo, Gasto Variável, Investimento, Cofrinho, etc.)
- Valor
- Status (Concluído, Pendente)

### 2️⃣ Dashboard 🚧 **[EM DESENVOLVIMENTO]**
Aba de visualização com gráficos e indicadores para análise visual dos dados financeiros.

**Status:** Planejado - ainda não implementado

**Recursos planejados:**
- Gráficos de evolução temporal de receitas e despesas
- Comparativos entre categorias de gastos
- Indicadores de performance financeira
- Análise de tendências e projeções

### 3️⃣ Config (Banco de Dados Auxiliar) 🚧 **[EM DESENVOLVIMENTO]**
Aba de configuração que funciona como banco de dados para listas e parâmetros do sistema.

**Status:** Planejado - ainda não implementado

**Conteúdo:**
- **Listas para Dropdowns:** Definição de opções para validação de dados
  - Tópicos (Gasto Fixo, Gasto Variável, Investimento, Cofrinho, etc.)
  - Status (Concluído, Pendente)
  - Outras categorias customizáveis
- **Parâmetros de configuração:** Valores padrão e regras de negócio
- **Tabelas auxiliares:** Dados de referência para fórmulas e automações

## 🔧 Tecnologias e Técnicas

### Excel Avançado
- **Fórmulas:** SOMASE, SOMASES, validações complexas
- **Formatação Condicional:** Feedback visual baseado em status
- **Validação de Dados:** Dropdowns dinâmicos conectados à aba Config
- **Células Mescladas e Layout:** Design profissional e organizado

### VBA (Visual Basic for Applications)
- **Event Handlers:** `Worksheet_Change` para automações em tempo real
- **Proteção de Planilha:** Gerenciamento inteligente de bloqueio/desbloqueio
- **Validações:** Verificação de saldo antes de transferências
- **Manipulação de Ranges:** Inserção e atualização dinâmica de dados
- **User Feedback:** Mensagens de erro e confirmação (MsgBox)

### Análise de Dados
- Agregação e sumarização de dados financeiros
- Cálculos automáticos de saldos e totalizadores
- Estruturação de dados para análise visual
- Preparação para dashboards dinâmicos

## ✨ Funcionalidades Implementadas

✅ **Automação de Datas**
- Data de realização preenchida automaticamente ao marcar como "Concluído"
- Data removida automaticamente ao voltar para "Pendente"

✅ **Sistema de Investimentos**
- Transferência automática de valores entre Cofrinho e Investimentos
- Validação de saldo disponível antes da transferência
- Registro automático de movimentações na tabela de despesas
- Uso de valores negativos para representar saídas do Cofrinho

✅ **Cálculo Automático de Saldo**
- Atualização em tempo real baseado em despesas e ganhos concluídos
- Fórmulas SOMASE para totalização por categoria
- Indicadores visuais na Carteira

✅ **Proteção e Segurança**
- Proteção de células críticas mantendo campos editáveis
- Tratamento de erros para evitar quebra de funcionalidades
- `Application.EnableEvents` gerenciado para prevenir loops infinitos

## 🚀 Próximos Passos

### Roadmap de Desenvolvimento

**Fase 1: Fluxo de Caixa** ✅ **CONCLUÍDO**
- [x] Sistema de registro de despesas e ganhos
- [x] Automação de datas baseado em status
- [x] Cálculo automático de saldo
- [x] Sistema de transferência entre Cofrinho e Investimentos
- [x] Validações e proteções de dados
- [x] Interface de Carteira digital

**Fase 2: Dashboard** 🚧 **EM ANDAMENTO**
- [ ] Desenvolver dashboards visuais com gráficos dinâmicos
- [ ] Implementar filtros e análises por período
- [ ] Criar relatórios automáticos mensais/anuais
- [ ] Adicionar indicadores de performance financeira
- [ ] Implementar análise de tendências e projeções

**Fase 3: Config** 🚧 **EM ANDAMENTO**
- [ ] Criar aba de configuração como banco de dados auxiliar
- [ ] Implementar listas dinâmicas para dropdowns
- [ ] Adicionar mais categorias e subcategorias de gastos customizáveis
- [ ] Criar parâmetros configuráveis pelo usuário

**Fase 4: Melhorias Futuras** 📋 **PLANEJADO**
- [ ] Implementar metas financeiras com acompanhamento visual
- [ ] Criar sistema de alertas para gastos acima da média
- [ ] Desenvolver análise preditiva de despesas
- [ ] Adicionar exportação de relatórios em PDF

## 📊 Estrutura de Dados

### Despesas
| Coluna | Descrição | Tipo |
|--------|-----------|------|
| C | Data Prevista | Data |
| D | Data Realização | Data (automática) |
| E | Descrição | Texto |
| F | Tópico | Lista (Config) |
| G | Valor | Moeda |
| H | Status | Lista (Config) |

### Ganhos
| Coluna | Descrição | Tipo |
|--------|-----------|------|
| V | Data Realização | Data (automática) |
| Y | Valor | Moeda |
| Z | Status | Lista (Config) |

### Carteira
| Item | Fórmula | Descrição |
|------|---------|-----------|
| Saldo | `=Ganhos - Despesas` | Saldo geral disponível |
| Valor Investido | `=SOMASE(F:F;"Investimento";G:G)` | Total em investimentos |
| Cofrinho | `=SOMASE(F:F;"Cofrinho";G:G)` | Total guardado no cofrinho |

## 🎓 Aprendizados e Desafios

Este projeto representa um estudo completo de como transformar o Excel de uma simples ferramenta de planilhas em uma **aplicação completa de gestão financeira**, demonstrando que com conhecimento aprofundado e criatividade, é possível criar soluções profissionais usando ferramentas acessíveis.

### 📍 Status Atual do Desenvolvimento

**Atualmente concluído:**
- ✅ **Fluxo de Caixa:** Totalmente funcional com todas as automações implementadas
- 🚧 **Dashboard:** Em desenvolvimento - estrutura sendo planejada
- 🚧 **Config:** Em desenvolvimento - ainda não iniciado

O foco inicial foi criar uma base sólida e totalmente funcional no Fluxo de Caixa, garantindo que todas as automações VBA e fórmulas trabalhem perfeitamente antes de expandir para outras áreas do projeto.

**Principais desafios superados:**
- Gerenciamento de eventos VBA sem quebrar funcionalidades existentes
- Trabalho com células mescladas e proteção de planilha
- Lógica de transferência entre categorias mantendo integridade dos dados
- Criação de experiência fluida e intuitiva para o usuário

---

## 📝 Notas de Desenvolvimento

**Versão atual:** 0.1.0 (Alpha)  
**Última atualização:** Fevereiro 2026  
**Desenvolvedor:** [Seu Nome]  

> "Demonstrando que Excel não é apenas uma planilha, mas uma plataforma completa de desenvolvimento." 🚀
