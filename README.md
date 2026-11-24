🇺🇸 English | [🇧🇷 Português](README.pt-BR.md)

# How to Read Excel Files (.XLSX) in Delphi Without Office or External Drivers (`uLeitorXLSX.pas`)

A lightweight, dependency-free native Delphi unit (**`uLeitorXLSX.pas`**) and VCL sample project demonstrating how to read and import **`.xlsx`** spreadsheets directly into a **`TFDMemTable` + `TDBGrid`** using only native RTL units (`System.Zip` and fast numeric cursor XML parsing) — with zero Microsoft Office dependency, no `ACE.OLEDB` driver conflicts, and no `Out of memory` crashes on large spreadsheets.

> 📖 **Full step-by-step tutorial, benchmarks, and production tips (in Portuguese):**  
> [Como Ler Planilha Excel no Delphi Sem Office e Sem Drivers](https://programacaolivre.com.br)

---

## 🚀 Quick Start (`uLeitorXLSX.pas`)

Add `uLeitorXLSX` to your form's `uses` clause and call the class methods:

```delphi
// 1. Automatically list all real worksheet names into a ComboBox
TLeitorXLSX.ListarAbas('C:\Data\spreadsheet.xlsx', cbxAbas.Items);

// 2. Load the chosen sheet starting from the exact header row (ALinhaInicio >= 1)
TLeitorXLSX.CarregarParaMemTable('C:\Data\spreadsheet.xlsx', 0, memExcel, 1, True);

// 3. Auto-fit TDBGrid column widths based on header and sample rows
TLeitorXLSX.AjustarLarguraColunasGrid(dbgExcel);
```

---

## 🎯 Key Features & Production Safeguards

* **Zero External Dependencies:** Does not require Microsoft Excel, *Access Database Engine*, or third-party components. Eliminates 32-bit vs. 64-bit Office driver conflicts on client machines.
* **Explicit Starting Row Control (`ALinhaInicio`):** Easily skip institutional headers or title rows at the top of official/government reports and start reading at the exact table header row.
* **High Speed Without `Out of Memory`:** Avoids heavy DOM trees (`IXMLDocument`), scanning internal XML via numeric cursor — benchmarked with **26,635 rows and 75 columns** in under 1 second on a **32-bit (`Win32`)** executable.
* **Data & Grid Formatting:** Converts Excel scientific notation (`7.89E+12`) back to clean integer strings, ignores empty cells formatted up to `XFD1`, and auto-sizes `TDBGrid` columns.

---

## 📂 Repository Structure

* `uLeitorXLSX.pas` — Decoupled, commented native `.xlsx` reader unit and `TDBGrid` auto-fit helper.
* `uPrincipal.pas` / `uPrincipal.dfm` — Ready-to-compile VCL demo application.

---

## 🛠️ Compatibility

* **Tested on:** Embarcadero Delphi Community Edition (Windows 32-bit and 64-bit).
* **Backward-compatible with:** Delphi 10 Seattle, Delphi XE, and newer versions supporting `System.Zip` and FireDAC (classic syntax, no inline variables, zero compiler warnings).

---

## ⚠️ Disclaimer

The code provided in this repository is intended for educational and general development purposes. Always test and validate the routines in your own staging environment and spreadsheet structures before deploying to production.

---

## 📄 Author

Maintained by **[Programação Livre](https://programacaolivre.com.br)**.
