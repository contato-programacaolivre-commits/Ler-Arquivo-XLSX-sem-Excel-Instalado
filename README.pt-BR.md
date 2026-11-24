[🇺🇸 English](https://github.com/contato-programacaolivre-commits/Ler-Arquivo-XLSX-sem-Excel-Instalado/edit/main/README.md) | 🇧🇷 Português

# Como Ler Planilha Excel no Delphi Sem Office e Sem Drivers (`uLeitorXLSX.pas`)

Unit reutilizável (**`uLeitorXLSX.pas`**) e projeto VCL de exemplo demonstrando como **ler planilha Excel no Delphi** (arquivos **`.xlsx`**) e importar os dados diretamente para um **`TFDMemTable` + `TDBGrid`** usando apenas recursos nativos (`System.Zip` e varredura XML por cursor numérico) — sem precisar ter o Microsoft Office instalado, sem drivers externos (`ACE.OLEDB`) e sem estourar memória em planilhas grandes.

> 📖 **Tutorial completo com explicação detalhada, testes de bancada e dicas de produção:**  
> [Como Ler Planilha Excel no Delphi Sem Office e Sem Drivers](https://programacaolivre.com.br)

---

## 🚀 Como Usar `uLeitorXLSX.pas` no Seu Projeto

Basta adicionar `uLeitorXLSX` na cláusula `uses` do seu formulário e chamar os métodos de classe:

```delphi
// 1. Lista automaticamente todas as abas reais da planilha no ComboBox
TLeitorXLSX.ListarAbas('C:\Dados\planilha.xlsx', cbxAbas.Items);

// 2. Carrega a aba escolhida a partir da linha exata onde começa a tabela (ALinhaInicio >= 1)
TLeitorXLSX.CarregarParaMemTable('C:\Dados\planilha.xlsx', 0, memExcel, 1, True);

// 3. Ajusta automaticamente a largura das colunas no TDBGrid
TLeitorXLSX.AjustarLarguraColunasGrid(dbgExcel);
```

---

## 🎯 Principais Recursos e Blindagens

* **Zero Dependências Externas:** Não exige Microsoft Office, não instala *Access Database Engine* e elimina conflitos de arquitetura entre Office 64-bit e executáveis Delphi 32-bit (`Win32`).
* **Controle Explícito de Linha Inicial (`ALinhaInicio`):** Permite pular cabeçalhos institucionais ou títulos soltos nas primeiras linhas de relatórios oficiais e iniciar a importação exatamente na linha das colunas reais.
* **Alta Performance Sem `Out of Memory`:** Evita árvores DOM pesadas (`IXMLDocument`), percorrendo o XML interno por cursor numérico — testado na bancada com **26.635 registros e 75 colunas** em menos de 1 segundo em `Win32`.
* **Tratamento de Códigos e Células Vazias:** Converte notações científicas do Excel (`7.89E+12`) de volta para dígitos inteiros limpos, ignora células vazias formatadas até `XFD1` e ajusta automaticamente a largura das colunas na grade.

---

## 📂 Arquivos do Repositório

* `uLeitorXLSX.pas` — Unit desacoplada e comentada com toda a lógica de leitura do `.xlsx` e ajuste visual do `TDBGrid`.
* `uPrincipal.pas` / `uPrincipal.dfm` — Formulário VCL de demonstração pronto para compilar e testar.

---

## 🛠️ Compatibilidade

* **Desenvolvido e testado em:** Embarcadero Delphi Community Edition (Windows 32-bit e 64-bit).
* **Retrocompatível com:** Delphi 10 Seattle, Delphi XE e versões superiores com `System.Zip` e FireDAC (sintaxe clássica, sem *inline variables* e sem *warnings* de compilação).

---

## ⚠️ Nota de Boas Práticas

Os códigos deste repositório têm finalidade educacional e são disponibilizados para uso livre. Valide e homologue a rotina no seu próprio ambiente e estrutura de planilhas antes de implantar em sistemas de produção.

---

## 📄 Autor

Desenvolvido e mantido pelo hub **[Programação Livre](https://programacaolivre.com.br)**.
