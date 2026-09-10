# Organizador-Imposto-de-Renda — IRPF organizado em uma planilha

![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![XLSX](https://img.shields.io/badge/.xlsx-workbook-1D6F42?style=flat-square)
![openpyxl](https://img.shields.io/badge/gerado%20com-openpyxl-4B8BBE?style=flat-square)
![DIO](https://img.shields.io/badge/DIO-desafio-8A2BE2?style=flat-square)

Organizador completo da declaração de **Imposto de Renda de Pessoa Física (IRPF)** em Excel: 7 abas com navegação por hiperlinks, validações automáticas de dados, fórmulas de totalização, formatação condicional e consolidação automática de rendimentos, deduções e bens em um resumo final. Gerada programaticamente com **openpyxl** (`gerar_planilha.py`) — a planilha é código versionável, não arquivo morto.

---

## 🇧🇷 Português

### Abas da planilha

| Aba | Finalidade |
|-----|------------|
| **Início** | Capa com menu de navegação (hiperlinks) para todas as seções e instruções de uso. |
| **Dados Pessoais** | Cadastro do contribuinte: nome, CPF, endereço, exercício. |
| **Rendimentos** | 40 linhas de lançamento com validação de tipo (tributável, isento, aplicações) e total automático. |
| **Deduções** | Despesas dedutíveis com lista suspensa de categorias e total automático. |
| **Bens e Direitos** | Patrimônio declarado por grupo de bens, com totalização. |
| **Resumo** | Consolidação automática: total de rendimentos, deduções, bens, base de cálculo e imposto estimado. |
| **Links Úteis** | Acesso rápido a Receita Federal, e-CAC, programa IRPF e materiais de apoio. |

### Funcionalidades

- **Menu de navegação** com hiperlinks entre abas (capa + retorno ao Início em cada aba);
- **Validações automáticas** (`Data Validation`) em listas suspensas para impedir erros de preenchimento;
- **Formatação condicional** destacando rendimentos acima de R$ 10.000,00;
- **Efeito zebrado** nas tabelas de leitura longa e **congelamento de cabeçalhos** nas de lançamento;
- **Fórmulas automáticas** de totalização e cálculo de base/imposto estimado.

### Como usar

1. Abra `Organizador_IRPF.xlsx` no Excel (ou LibreOffice Calc).
2. Navegue pelo menu na aba **Início**.
3. Preencha Dados Pessoais, Rendimentos, Deduções e Bens e Direitos usando as listas suspensas.
4. Consulte o consolidado na aba **Resumo**.

> O cálculo de imposto na aba Resumo é **genérico e educativo** — não substitui o programa oficial da Receita nem um contador.

### Por que gerar com openpyxl

A planilha inteira é descrita em `gerar_planilha.py`: abas, validações, fórmulas e formatação nascem de código. Isso torna o entregável **reprodutível** (rodar o script regenera o arquivo idêntico), versionável no Git (diffs fazem sentido) e auditável — cada escolha de fórmula está no repositório, não escondida em células.

### Autor

**Kelvin Oliveira** — [GitHub](https://github.com/KelvinOliveiraCode) · [LinkedIn](https://www.linkedin.com/in/kelvin-oliveira-code/)

---

## 🇺🇸 English

A complete **Brazilian individual tax return (IRPF)** organizer in Excel: 7 sheets with hyperlink navigation, automatic data validation, totaling formulas, conditional formatting and automatic consolidation of income, deductions and assets into a final summary. Generated programmatically with **openpyxl** (`gerar_planilha.py`) — the spreadsheet is versionable code, not a dead file.

### Sheets

| Sheet | Purpose |
|-------|---------|
| **Início** | Cover page with hyperlink navigation menu to every section. |
| **Dados Pessoais** | Taxpayer registration: name, CPF, address, tax year. |
| **Rendimentos** | 40 entry rows with type validation (taxable, exempt, investments) and automatic totals. |
| **Deduções** | Deductible expenses with category dropdown lists and automatic totals. |
| **Bens e Direitos** | Declared assets grouped by type, with totalization. |
| **Resumo** | Automatic consolidation: income, deductions, assets, taxable base and estimated tax. |
| **Links Úteis** | Quick access to Receita Federal, e-CAC, the official IRPF program and support material. |

### Features

- Hyperlink navigation between sheets (cover + back-to-start on every sheet);
- **Data Validation** dropdown lists preventing entry errors;
- **Conditional formatting** highlighting income above R$ 10,000;
- Zebra striping on long tables and frozen headers on entry tables;
- Automatic totaling and estimated-tax formulas.

> The tax estimate in Resumo is **generic and educational** — it does not replace the official Receita program or an accountant.

### Why generate with openpyxl

The entire workbook is described in `gerar_planilha.py`: sheets, validations, formulas and formatting are born from code. That makes the deliverable **reproducible** (running the script regenerates an identical file), Git-versionable (diffs make sense) and auditable — every formula choice lives in the repository, not hidden inside cells.

### Author

**Kelvin Oliveira** — [GitHub](https://github.com/KelvinOliveiraCode) · [LinkedIn](https://www.linkedin.com/in/kelvin-oliveira-code/)
