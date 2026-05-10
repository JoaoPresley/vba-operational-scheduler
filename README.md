# 📊 Automação de Programação de Atividades - Operação e serviços de campo

## 📌 Sobre o Projeto
Este sistema foi desenvolvido em **VBA (Visual Basic for Applications)** para otimizar o planejamento e a distribuição de tarefas de equipes técnicas. A solução utiliza um motor de processamento centralizado onde um único input de dados alimenta automaticamente múltiplas visões operacionais.

### O Fluxo de Trabalho:
1. **Input Central:** Preenchimento da aba de programação com OSs, técnicos e diretrizes.
2. **Processamento VBA:** O motor lê a matriz de dados, identifica períodos (Manhã/Tarde) e distribui as informações.
3. **Saída Automática 1 (Semanal):** Gera um painel de visão macro da semana.
4. **Saída Automática 2 (Diária):** Gera folhas de serviço detalhadas por dia, prontas para uso operacional.

## ⚠️ Nota sobre a Versão de Demonstração
A estrutura original deste sistema foi projetada para gerenciar simultaneamente três setores: **Operação (OCGR)**, **Três Lagoas (TLG)** e **Manutenção (MCGR)**. 

Para esta exibição no GitHub:
* As abas e dados referentes a **TLG** e **MCGR** foram removidos para garantir a privacidade e segurança de informações corporativas sensíveis.
* **O código-fonte (VBA) permanece íntegro:** Toda a lógica modular original foi mantida no projeto, demonstrando a capacidade do sistema de escalar para múltiplos setores através de parametrização.
* O funcionamento da planilha de **Operação (OCGR)** apresentada não sofreu qualquer alteração, operando exatamente como na versão de produção.

## 🛠️ Especificações Técnicas
* **Compatibilidade:** Desenvolvido para **Microsoft Excel 2007**, mantendo alta performance em versões legadas do Office.
* **Destaques de Código:** * Automação via Eventos (`Workbook_SheetActivate`).
    * Funções modulares que recebem parâmetros de setor (ex: `Call Prog_Diária("OCGR", Sh.Name)`).
    * Algoritmos de busca e organização de matrizes de dados.

## 📸 Demonstração (Versão Sanitizada)
*(Insira aqui os links das imagens que você limpou como "Tom 1", "Tom 2")*

## 📂 Estrutura do Repositório
* `/src`: Código-fonte exportado (`.bas` e `.cls`) para visualização direta no navegador.
* `Automacao_Operacional.xlsm`: Ficheiro Excel funcional (habilitado para macros).

---
*Projeto desenvolvido por [Seu Nome/GitHub User]*
