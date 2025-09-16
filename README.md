# Repositório de Scripts VBA para Automação de Rotinas

Este é um repositório dedicado a centralizar, mesclar e integrar códigos VBA (Visual Basic for Applications) projetados para maximizar a velocidade e eficiência de rotinas de automação.

O foco principal do projeto é a integração entre aplicações do Microsoft Office (Excel, Access, Word) e a automação de tarefas web (Web Scraping/Preenchimento de Formulários) através do Selenium.

## Tecnologias e Componentes

* **Linguagem Principal:** VBA (Visual Basic for Applications)
* **Plataforma:** Microsoft Excel (`.xlsm`), Microsoft Access (`.accdb`), Microsoft Word (`.docx`)
* **Automação Web:** Selenium Basic (uso explícito de `WebDriver` e `ChromeDriver`)
* **Templates:** Contém módulos (`.bas`, `.vb`), um banco de dados Access e um documento Word para preenchimento.

---

## Análise dos Principais Scripts

Este repositório contém várias macros de automação. As principais incluem:

### 1. Envio de Formulário em Massa (Send-Mass-Text.bas)

Script de automação web que utiliza Selenium para preencher e enviar um formulário web repetidamente com dados de uma planilha.

* **Função:** `enviarFormularioEmMassa()`.
* **Dependência:** Selenium Basic (WebDriver/ChromeDriver).
* **Fluxo de Processo:**
    1.  A macro lê dados da planilha Excel chamada "Dados".
    2.  Inicia um loop que continua até encontrar a string "Parar" na coluna A.
    3.  Para cada linha, ele instancia um novo `ChromeDriver` e navega até um formulário SurveyMonkey (`.../Y9Y6FFR`).
    4.  Utiliza seletores `FindElementByName` (ex: "683928983", "683932318") e `FindElementById` (ex: "683931881...") para preencher os campos do formulário (Nome, Email, Telefone, Sobre, Gênero) com os dados da planilha.
    5.  Envia o formulário usando `FindElementByXPath` para localizar o botão de envio.
    6.  Fecha o navegador e repete o processo para a próxima linha.

### 2. Abertura de PDFs em Lote (openingPDF_Files.vb)

Script utilitário para abrir múltiplos arquivos PDF listados em uma planilha.

* **Função:** `OpenPdfs()`.
* **Fluxo de Processo:**
    1.  Lê a "Planilha2" e faz um loop pela Coluna A.
    2.  Para cada célula contendo um nome de arquivo, ele executa um comando `Shell` para abrir o executável do Adobe Acrobat (`AcroRd32.exe`) com o caminho do arquivo especificado.

### 3. Integração de Dados (Access, Word e Excel)

O repositório está estruturado para uma automação de preenchimento de documentos:

* **Banco de Dados (`BaseDados.accdb`):** Contém as tabelas centrais, como "Pessoas", "Equipamentos" e "Tabela 1", que armazenam os dados brutos (ex: CPF, Nome Completo, Produto, Preço).
* **Modelo de Documento (`Contrato.docx`):** Um arquivo Word padrão que serve como modelo, contendo placeholders para os dados armazenados no Access (CPF, Nome Completo, RG, Produto, Preço, Data, etc.).
* **Arquivos de Orquestração:** Os arquivos `.xlsm` (como `Sistema_Banco_Firebird.xlsm`) e o módulo `DB-ACCESS.bas` são destinados a conter as macros VBA que executam a conexão com o banco de dados ( via ADO ou DAO) e realizam a automação de preenchimento (mala direta) do `Contrato.docx`.

### 4. Interface de Checklist Cíclico (CycleChecklist.vb)

Script VBA projetado para gerenciar o estado de uma interface de usuário (UI) dentro do Excel, atuando como um checklist de processo.

* **Função:** `Checklist_Cíclico()`.
* **Fluxo de Processo:**
    1.  Lê o número da OS de uma caixa de texto (Shape "txtOS").
    2.  O código referencia vários botões (Shapes) que representam etapas do processo, como "btnCriarPasta", "btnBaixarImagens", "btnAbrirLink", e "btnCriarRelatorio".
    3.  Ao final, marca todas as checkboxes (`msoShapeCheckbox`) como concluídas e limpa o campo "txtOS" para reiniciar o ciclo.

## Requisitos de Setup

1.  **Microsoft Office:** Requer Microsoft Excel e Microsoft Access (para os scripts de banco de dados) e Microsoft Word (para o modelo de contrato).
2.  **Selenium Basic:** O script `Send-Mass-Text.bas` requer a instalação do **Selenium Basic** (VBA) e o `ChromeDriver.exe` correspondente à versão do Google Chrome instalada.
3.  **Referências VBA:** Pode ser necessário habilitar referências específicas no editor VBA (Alt+F11 > Ferramentas > Referências), como `Microsoft ActiveX Data Objects` (para ADO/Access) e `Selenium Type Library` (para o script web).
