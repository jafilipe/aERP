# aERP — Visão geral

## 1. O que é o aERP

O **aERP** é uma solução de *Business Intelligence* em Excel para empresas que usam o ERP **PHC**. Lê os dados de vendas, compras, clientes, fornecedores e artigos diretamente da base de dados do PHC e apresenta-os em **dashboards interativos**, **folhas de consulta** e **comparações anuais** ao estilo do PHC.

É composto por dois ficheiros:

| Ficheiro | Papel |
|---|---|
| **aERP.xlsm** | O livro do utilizador, com os dashboards, as consultas e a folha de resumo. |
| **aERP_Addin.xlam** | O suplemento que trata da ligação ao PHC e da sincronização dos dados. |

A ideia central: **o utilizador não precisa de saber Power Pivot nem SQL**. Configura a ligação uma vez, carrega em **Sincronizar dados** e trabalha só com botões e filtros.

### Para quem
- Gestores e administrativos de PME com PHC que querem indicadores de vendas, margem, compras e stock sem relatórios à medida.
- Parceiros e consultores PHC que querem entregar um pacote de BI pronto a instalar.

---

## 2. Como funciona

1. O utilizador indica o servidor e a base de dados do PHC e carrega em **Guardar**. O nome e o logótipo da empresa são lidos do próprio PHC.
2. Ao carregar em **Sincronizar dados**, o aERP lê do PHC as vendas, compras, clientes, fornecedores e artigos e carrega-os no modelo de dados do Excel.
3. Os dashboards, as consultas e as comparações atualizam-se sozinhos, já com os dois anos mais recentes escolhidos.

---

## 3. Dados disponíveis

| Tabela | Conteúdo |
|---|---|
| **Vendas** | Documentos de venda e respetivas linhas |
| **Compras** | Documentos de compra e respetivas linhas |
| **Clientes** | Tipo, zona, vendedor, país, código postal, estado… |
| **Fornecedores** | Ficha de fornecedores |
| **Artigos** | Família, marca, modelo, linha, grupo, classe, stock… |
| **Calendário** | Datas para as análises por ano e por mês |
| **Códigos postais** | Lista dos CTT, para mostrar clientes e fornecedores por **distrito** e **concelho** |

Sobre estes dados existem cerca de 76 medidas: vendas líquidas, compras líquidas, custo, **margem bruta**, margem %, quantidades, número de clientes, fornecedores e artigos, ticket médio, entre outras.

---

## 4. Funcionalidades

### 4.1 Resumo e navegação
- Página inicial com uma **visão geral** dos gráficos principais de cada dashboard.
- **Cartões de indicadores** do ano mais recente no topo do resumo.
- Barra lateral com **logótipo e nome da empresa**, filtro de **Ano**, botão **Período**, botões para todos os dashboards e consultas, **Configuração** e **Sincronizar dados**.

### 4.2 Dashboards (7)
Todos com o mesmo aspeto: fundo escuro, barra lateral com logótipo e empresa, cartões de indicadores e grelha de gráficos alinhada.

| Dashboard | Destaques |
|---|---|
| **Vendas** | Vendas e **margem bruta por mês**; top 10 clientes por vendas e por margem; por tipo de cliente, vendedor, distrito, concelho e país. |
| **Vendas por Artigo** | Vendas e margem por cliente e por artigo (família, marca, modelo, linha…). |
| **Compras** | Compras por mês, top fornecedores, distrito e concelho. |
| **Compras por Artigo** | Compras e quantidades por artigo, família e marca. |
| **Clientes** | Clientes ativos por zona, tipo, vendedor, país, distrito e concelho; top 10 por saldo. |
| **Fornecedores** | Fornecedores ativos por zona, tipo, distrito e concelho; top 10. |
| **Artigos** | Artigos por família, marca, modelo, linha, tipo, grupo e classe; top 10 por stock. |

Outros pormenores:
- Gráfico circular de **zona fiscal** (nacional ou estrangeiro) em cada dashboard.
- **Campos sem dados escondem-se sozinhos**: se um campo está vazio na base de dados, o gráfico e o filtro desaparecem e o resto reorganiza-se.
- Botões para passar entre dashboards relacionados, por exemplo "Ver detalhe por artigo".

### 4.3 Filtros
- Campos com poucos valores usam **segmentações compactas**.
- Campos com muitos valores usam **listas com escolha múltipla** e caixa **Procurar**.
- Botão **Limpar filtros** em cada folha.

### 4.4 Comparação de anos
- O filtro **Ano** aceita um ou dois anos, e o ano escolhido passa para todas as folhas.
- Com dois anos, **todos os gráficos comparam**: barras com duas séries, circulares com dois anéis e títulos "2026 vs 2025".
- Os tops 10 são ordenados pelo ano mais recente.
- Botão **Período** para analisar o **ano completo**, **até ao mês atual**, **até ao fim do mês**, um **trimestre** ou um **semestre**. As comparações usam o mesmo período nos dois anos e os títulos indicam o período escolhido.

### 4.5 Comparação anual ao estilo PHC
- Tabela mês a mês nos dashboards de Vendas, Vendas por Artigo, Compras e Compras por Artigo.
- Valores dos dois anos, **acumulados**, **desvios em valor e em %** e rentabilidade, com setas ▲ ▼.
- Com um só ano escolhido, compara com o ano anterior.
- Botão **Análise** para escolher o que mostrar, com detalhe por **cliente**, **fornecedor** ou **artigo**.

### 4.6 Consultas
- Uma folha de consulta por tabela, com a lista completa e os mesmos filtros do dashboard correspondente.
- **Ordenar por** qualquer coluna, de A a Z ou de Z a A, com opção para voltar à ordem original.
- Cada consulta é criada só quando é aberta, para a sincronização ser rápida.

### 4.7 Segurança
- Dashboards, resumo e consultas protegidos: o utilizador usa botões e filtros, e gráficos e objetos não se mexem.
- A estrutura do livro está protegida e o código VBA está protegido por senha.
- Se o livro for aberto noutro PC, os dados são limpos automaticamente, para não se distribuírem por engano.

---

## 5. Instalação

O instalador é distribuído no ficheiro `aERP_Setup_1.0.0.zip`: basta extraí-lo e executar o `aERP_Setup_1.0.0.exe`. O instalador não precisa de administrador e:
- copia o livro, o suplemento e a lista de códigos postais para `Documentos\aERP`;
- regista o suplemento no Excel e autoriza as macros da pasta do aERP;
- cria os atalhos;
- verifica se o **Microsoft OLE DB Driver 18 para SQL Server** está instalado e, se faltar, instala-o (x64 em Windows de 64 bits, que inclui também o de 32 bits; x86 em Windows de 32 bits). Só este passo pede autorização de administrador ao Windows. Sem o driver, o teste de ligação dá o erro *"3706 - Não foi encontrado o fornecedor"*.

A desinstalação remove o registo.

### Requisitos
- Windows com **Excel 365 / 2021 de 64 bits** (com Power Pivot).
- Acesso de rede ao **SQL Server do PHC**.
- Ligação à Internet.

---

## 6. Tecnologias

- **Excel**: Power Pivot (modelo de dados, DAX), Power Query (M), funções CUBE, segmentações e tabelas dinâmicas sobre o modelo, matrizes dinâmicas.
- **VBA**: formulários, ligação ao SQL Server (ADO), eventos do livro, gráficos e formas.
- **SQL Server / PHC**: origem dos dados.
- **Inno Setup**: instalador.
