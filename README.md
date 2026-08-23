# Organizador de Declaração de Imposto de Renda

Ferramenta completa desenvolvida em **Microsoft Excel** para organizar e reunir as informações essenciais da declaração de Imposto de Renda de Pessoa Física (IRPF).

> Projeto criado para o desafio da [DIO](https://www.dio.me/) — "Construindo um Organizador de Declaração de Imposto de Renda no Excel".
---

## 📁 Estrutura do Projeto

```
Organizador de Declaração de Imposto de Renda/
├── Organizador_IRPF.xlsx     # Planilha principal (entregável)
├── gerar_planilha.py         # Script Python (openpyxl) que gera a planilha
├── README.md                 # Este arquivo
└── images/                   # (opcional) capturas de tela
```

---

## 📊 Abas da Planilha

| Aba | Finalidade |
|-----|------------|
| **Início** | Capa com menu de navegação (hiperlinks) para todas as seções e instruções de uso. |
| **Dados Pessoais** | Cadastro do contribuinte (nome, CPF, endereço, exercício, etc.). |
| **Rendimentos** | Controle de entradas com validação de tipo (tributável, isento, aplicações…), 40 linhas e total automático. |
| **Deduções** | Despesas dedutíveis com lista suspensa de categorias e total automático. |
| **Bens e Direitos** | Patrimônio declarado com grupos de bens e totalização. |
| **Resumo** | Consolidação automática: total de rendimentos, deduções, bens, base de cálculo e imposto estimado (exemplo genérico). |
| **Links Úteis** | Acesso rápido a sites oficiais e materiais de apoio (estrutura extra de links). |

---

## ✨ Funcionalidades

- **Menu de navegação** com hiperlinks entre as abas (na capa e no topo de cada aba, retorno ao Início).
- **Validações automáticas** (`Data Validation`) em listas suspensas para evitar erros de preenchimento.
- **Fórmulas automáticas** de totalização e cálculo de base/imposto estimado.
- **Formatação condicional** que destaca rendimentos acima de R$ 10.000,00.
- **Efeito zebrado** para facilitar a leitura de grandes tabelas.
- **Links rápidos** para Receita Federal, e-CAC, programa IRPF e materiais de apoio DIO/GitHub.
- **Congelamento de cabeçalhos** nas tabelas de lançamento.

---

## 🚀 Como Usar

1. Abra o arquivo `Organizador_IRPF.xlsx` no Excel (ou LibreOffice Calc).
2. Na aba **Início**, clique nos itens do menu para navegar.
3. Preencha **Dados Pessoais**, **Rendimentos**, **Deduções** e **Bens e Direitos**.
4. Utilize as listas suspenses para os campos validados.
5. Acesse **Resumo** para ver o consolidado automático.
6. Use **Links Úteis** para consultar sites oficiais durante o preenchimento.

> ⚠️ O cálculo de imposto na aba Resumo é **genérico e educativo**. Consulte sempre a legislação vigente e um profissional contábil.

---

## 🔧 Como Regenerar a Planilha

O arquivo `.xlsx` foi gerado com Python + [openpyxl](https://openpyxl.readthedocs.io/). Para recriá-lo:

```bash
pip install openpyxl
python gerar_planilha.py
```

---

## 📚 Recursos Úteis

### Documentações Oficiais
- [Receita Federal – Meu Imposto de Renda](https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda)
- [Programa IRPF (download)](https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/download/pgd)
- [Tabelas Práticas da IE](https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/irpf/2025/tabelas)
- [Início Rápido: VM Windows no Azure](https://learn.microsoft.com/pt-br/azure/virtual-machines/windows/quick-create-portal)

### GitHub / Documentação
- [GitHub Quick Start (DIO)](https://github.com/digitalinnovationone/github-quick-start)
- [Documentação do GitHub](https://docs.github.com/pt)
- [GitHub Markdown Guide](https://docs.github.com/pt/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github)
- ✅ Pasta `images/` para capturas de tela (opcional)

**Bons estudos!** 🚀
