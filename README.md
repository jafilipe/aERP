# aERP — Business Intelligence em Excel para PHC

> Dashboards, consultas e comparações anuais sobre os dados do ERP **PHC**, dentro do Excel, sem o utilizador precisar de saber Power Pivot nem SQL.

O **aERP** liga-se à base de dados SQL Server do PHC, carrega os dados de vendas, compras, clientes, fornecedores e artigos num modelo **Power Pivot** e apresenta-os em **dashboards interativos**, **folhas de consulta** e **comparações anuais** ao estilo do PHC. O utilizador configura a ligação uma vez, carrega em **Sincronizar dados** e trabalha só com botões e filtros.

---

## O que faz

- **Liga-se ao PHC** (SQL Server) e traz as tabelas de movimentos e fichas para um modelo Power Pivot, sem o utilizador escrever SQL.
- **7 dashboards** com o mesmo aspeto (Vendas, Vendas por Artigo, Compras, Compras por Artigo, Clientes, Fornecedores, Artigos): KPIs, gráficos por mês, tops 10, margem bruta, distribuição por distrito/concelho/país e por tipo, vendedor ou família.
- **Comparação de anos**: escolhe um ou dois anos e todos os gráficos passam a comparar (barras com duas séries, donuts com dois anéis, títulos "2026 vs 2025"); o ano escolhido é sincronizado entre todas as folhas.
- **Comparação anual ao estilo PHC**: tabela mês a mês com acumulados, desvios em valor e em %, rentabilidade e setas ▲▼, com detalhe por cliente, fornecedor ou artigo.
- **Folhas de consulta**: uma lista completa por tabela, com os mesmos filtros do dashboard, criada só quando é aberta (para a sincronização ser rápida).
- **Filtros fáceis**: segmentações compactas para campos pequenos e botões de lista com pesquisa e escolha múltipla para campos grandes; campos sem dados escondem-se sozinhos.
- **Nome e logótipo da empresa** lidos automaticamente do próprio PHC.
- **Pronto a instalar**: um instalador regista o suplemento, cria o local fidedigno das macros e copia tudo para a pasta do utilizador, sem precisar de administrador.

---

## Descarregar

Os ficheiros para experimentar estão na secção **[Releases](../../releases)**:

- **`aERP_Setup_1.0.0.exe`** — instalador do aERP (livro + suplemento + lista de códigos postais).
- **`Video aERP.mp4`** — vídeo explicativo que percorre as principais funcionalidades.

> Dados sensíveis (servidor, base de dados, credenciais) são de demonstração nos materiais publicados.

---

## Como funciona

```
  SQL Server (PHC)                aERP_Addin.xlam                 aERP.xlsm
  ┌───────────────┐   SELECT     ┌────────────────┐   CSV temp.  ┌────────────────────┐
  │ ft, fi, fo,   │ ───────────▶ │ Configuração   │ ───────────▶ │ Power Query (M)    │
  │ cl, fl, st,   │  por tabela  │ Sincronização  │              │        ▼           │
  │ para1 ...     │  (ADO)       │ Licenças       │              │ Power Pivot (DAX)  │
  └───────────────┘              └───────┬────────┘              │        ▼           │
                                         │                       │ Funções CUBE       │
       Google Sheet (licença/log) ◀──────┘                       │        ▼           │
       via Google Apps Script                                    │ Dashboards/Consultas│
                                                                 └────────────────────┘
```

1. O suplemento executa um `SELECT` por tabela no SQL Server do PHC e grava o resultado num **CSV temporário** (sem o limite de linhas de uma folha).
2. O **Power Query** lê o CSV, aplica os tipos e carrega para o **Power Pivot**.
3. Os dashboards usam funções **CUBE** e tabelas dinâmicas OLAP sobre o modelo.
4. A licença é validada a cada arranque numa Google Sheet (via Google Apps Script), com tolerância em modo sem ligação.

---

## Requisitos

- Windows com **Excel 365 / 2021 de 64 bits** (com Power Pivot).
- Acesso de rede ao **SQL Server do PHC**.
- Ligação à Internet para validar a licença (com tolerância offline).

---

## Documentação

- [`docs/aERP_visao_geral.md`](docs/aERP_visao_geral.md) — visão geral completa: arquitetura, modelo de dados, funcionalidades, suplemento, instalador, tecnologias, limitações e guião do vídeo.

---

## Tecnologias

Power Pivot (modelo, DAX) · Power Query (M) · funções CUBE e segmentações OLAP · VBA (UserForms, ADO, eventos de livro) · SQL Server (PHC) · Google Apps Script + Google Sheets (licenças) · Inno Setup (instalador).

---

## Estado

Projeto em desenvolvimento. O instalador ainda não foi testado num PC de cliente final. Ver as limitações conhecidas e o trabalho futuro na [visão geral](docs/aERP_visao_geral.md#8-limitações-conhecidas-e-trabalho-futuro).

---

_Repositório privado. Todos os direitos reservados._
