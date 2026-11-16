# aERP — Visão geral do projeto

> Documento de apoio para preparar a documentação de publicação no GitHub (README, wiki) e o guião de um vídeo explicativo.
> Compilado a 14/11/2026 e atualizado a 16/11/2026 a partir das notas, análises e código do projeto. Não contém cadeias de ligação, palavras-passe, nomes de servidores nem dados de clientes.

---

## 1. O que é o aERP

O **aERP** é uma solução de *Business Intelligence* em Excel para empresas que usam o ERP **PHC**. Lê os dados de vendas, compras, clientes, fornecedores e artigos diretamente da base de dados SQL Server do PHC, carrega-os num modelo **Power Pivot** e apresenta-os em **dashboards interativos**, **tabelas de consulta** e **comparações anuais** ao estilo do PHC.

É composto por dois ficheiros:

| Ficheiro | Papel |
|---|---|
| **aERP.xlsm** | O livro do utilizador: modelo de dados (Power Pivot + Power Query), dashboards, folhas de consulta, folha de resumo "#aERP" e macros de interface (filtros, anos, proteção). |
| **aERP_Addin.xlam** | O suplemento: configuração da ligação ao SQL, extração dos dados (sincronização), licenciamento, logótipo e nome da empresa, limpeza de dados. |

A ideia central: **o utilizador não precisa de saber Power Pivot nem SQL**. Configura a ligação uma vez, carrega em "Sincronizar dados" e trabalha só com botões e filtros.

### Para quem
- Gestores e administrativos de PME com PHC que querem indicadores de vendas, margem, compras e stock sem relatórios à medida.
- Parceiros/consultores PHC que querem entregar um pacote de BI pronto a instalar.

---

## 2. Arquitetura

```
┌──────────────────────┐      SELECT por tabela        ┌─────────────────────────┐
│  SQL Server (PHC)    │ ───────────────────────────▶ │  aERP_Addin.xlam        │
│  ft, fi, fo, fn, cl, │   (ADO, credenciais da       │  - frmSQLManager        │
│  fl, st, para1 ...   │    configuração)             │  - ProcessarAgendamento │
└──────────────────────┘                              │  - LimpaDados           │
                                                      │  - Licenças             │
          Google Sheet "aERP"  ◀── licença/log ──────▶│                         │
          (folhas SP, Lic, Log)    (Apps Script)       └───────────┬─────────────┘
                                                                  │ CSV temporário
                                                                  ▼ (%TEMP%\aERP_dados)
                                                      ┌─────────────────────────┐
                                                      │  aERP.xlsm              │
                                                      │  Power Query (Csv.Doc.) │
                                                      │        ▼                │
                                                      │  Power Pivot (modelo)   │
                                                      │        ▼                │
                                                      │  CUBEVALUE / CUBESET    │
                                                      │  Tabelas dinâmicas OLAP │
                                                      │        ▼                │
                                                      │  Dashboards / Consultas │
                                                      └─────────────────────────┘
```

### Fluxo de dados (sincronização)
1. O add-in lê da Google Sheet "aERP" (folha **SP**) a lista de consultas: uma linha por tabela do modelo, com o SQL de extração.
2. Para cada tabela, executa o SELECT no SQL Server do PHC por **ADO**.
3. Grava o resultado num **CSV temporário** (UTF-16, separador próprio), para não ter o limite de 1 048 575 linhas de uma folha Excel.
4. A consulta **Power Query** dessa tabela lê o CSV (`Csv.Document`), aplica os tipos de dados e carrega para o **Power Pivot**.
5. No fim, os CSV são apagados e o "Atualizar tudo" do Excel fica bloqueado nessas ligações (os dados só se atualizam pelo aERP).
6. O livro escolhe automaticamente os **dois anos mais recentes**, atualiza filtros, gráficos e volta à folha #aERP.

### Evolução das origens de dados (decisões técnicas)
- **Início:** consultas Power Query ligadas diretamente ao SQL (`Sql.Database`). Problema: noutro PC aparecia a janela de credenciais do Excel, e as credenciais ficavam no PC.
- **"Dados no livro" (02/11):** o add-in passa a ler por ADO e grava em folhas ocultas; o Power Query lê dessas folhas. Sem janela de credenciais.
- **CSV temporário (04/11):** uma base de cliente tinha 1,8 milhões de linhas numa tabela, acima do limite de uma folha. Passou-se a CSV temporário, apagado no fim.

---

## 3. Modelo de dados

### Tabelas
| Tabela | Conteúdo | Origem PHC (aprox.) |
|---|---|---|
| **MovVendas** | Cabeçalhos dos documentos de venda | faturação |
| **MovVendasLinhas** | Linhas dos documentos de venda | linhas de faturação |
| **MovCompras** | Cabeçalhos dos documentos de compra | compras |
| **MovComprasLinhas** | Linhas dos documentos de compra | linhas de compras |
| **FichasClientes** | Ficha de clientes (tipo, zona, vendedor, país, código postal, estado…) | clientes |
| **FichasFornecedores** | Ficha de fornecedores | fornecedores |
| **FichasArtigos** | Ficha de artigos (família, marca, modelo, linha, grupo, classe, stock…) | artigos |
| **Calendário** | Tabela de datas | gerada |
| **CodigosPostais** | Lista CTT (≈193 mil códigos postais) com distrito e concelho | ficheiro CSV externo |

### Relações principais
- Linhas → cabeçalho do documento; Linhas → FichasArtigos (código do artigo).
- MovVendas → FichasClientes; MovCompras → FichasFornecedores (IDFornecedor).
- FichasClientes / FichasFornecedores → **Distrito e Concelho** a partir do código postal (junção com a lista dos CTT feita no Power Query).

### Medidas
- Cerca de 76 medidas DAX: vendas líquidas, compras líquidas, custo, **margem bruta**, margem %, quantidades, nº de clientes/fornecedores/artigos, ticket médio, etc.

### Códigos postais (CTT)
- Ficheiro **CodigosPostais.csv** (≈9 MB) partilhado pelos livros, instalado junto do aERP. Retirá-lo de dentro do livro reduziu o ficheiro de ~16 MB para ~5,5 MB.

---

## 4. Folhas e funcionalidades do livro

### 4.1 "#aERP" — Resumo e navegação
- Página inicial: **Visão geral** com cópias dos gráficos principais de cada dashboard.
- Barra lateral com **logótipo e nome da empresa**, filtro de **Ano**, botões de navegação para todos os dashboards e consultas, botão **Configuração** (abre o formulário de ligação) e **Sincronizar dados**.
- Versão e contactos no rodapé da barra lateral.

### 4.2 Dashboards (7)
Todos com o mesmo aspeto: fundo escuro, barra lateral com logótipo/empresa, botões no topo do painel (Resumo, Consulta, Limpar filtros), cartões KPI arredondados e grelha de gráficos alinhada.

| Dashboard | Destaques |
|---|---|
| **Vendas** | Vendas por mês e **margem bruta por mês** (12 meses sempre visíveis); top 10 clientes por vendas e por margem; por tipo de cliente, vendedor, distrito, concelho, país; **Comparação anual** ao estilo PHC no fundo. |
| **Vendas Linhas** (vendas por artigo) | Vendas e margem por campos de cliente e de artigo (família, marca, modelo, linha…); filtros Cliente e Artigo; comparação anual por artigo e por cliente. |
| **Compras** | Compras por mês, top fornecedores, distrito, concelho; filtro Fornecedor; comparação anual. |
| **Compras Linhas** (compras por artigo) | Compras e quantidades por artigo/família/marca; comparação anual por artigo e fornecedor. |
| **Clientes** | Clientes ativos por zona, tipo, vendedor, país, distrito, concelho; top 10 por saldo de conta corrente. |
| **Fornecedores** | Fornecedores ativos por zona, tipo, distrito, concelho; top 10. |
| **Artigos** | Artigos por família, marca, modelo, linha, tipo de produto, grupo, classe; top 10 por stock. |

Detalhes que valem a pena mostrar:
- **Donut de zona fiscal** (nacional/estrangeiro) no canto superior de cada dashboard, com a legenda sempre completa (a altura e a largura da legenda ajustam-se ao número de categorias).
- **Nome da empresa e logótipo** lidos do próprio PHC ao configurar a ligação.
- **Campos vazios escondidos automaticamente**: se um campo não tem dados na base (ex.: "zona" vazia em todos os clientes), o gráfico e o filtro desaparecem e o resto reorganiza-se.
- **Botões de navegação** entre dashboards relacionados ("Ver detalhe por artigo ›", "‹ Voltar a Vendas").

### 4.3 Filtros
- **Painel misto**: campos com poucos valores usam segmentações compactas; campos com muitos valores usam **botões de lista com escolha múltipla** (formulário próprio com caixa **Procurar**, sem distinguir acentos, botões Todos/Limpar/OK).
- Botão **Limpar filtros** em cada folha.
- Ao mudar de base de dados, os filtros passam a mostrar os valores da nova base (sem itens antigos).

### 4.4 Comparação de anos
- Segmentação **Ano** permite **um ou dois anos**; mais do que dois fica com os dois mais recentes.
- O ano escolhido num dashboard é **sincronizado** com os outros e com o Resumo.
- Com 2 anos, **todos os gráficos comparam** (barras com duas séries, donuts com dois anéis, legenda) e os títulos ficam "… – 2026 vs 2025".
- As listas de top 10 são ordenadas pelo **ano mais recente** escolhido.
- Rótulos dos gráficos mensais adaptam-se (verticais com 2 anos, horizontais com 1) e os zeros são escondidos.

### 4.5 Comparação anual (ao estilo PHC)
- Tabela mês a mês (Jan…Dez) no fundo dos dashboards de Vendas, Vendas Linhas, Compras e Compras Linhas.
- Linhas de valores dos dois anos, **acumulados**, **desvio em valor e em %**, rentabilidade (margem), com setas ▲ ▼ verdes/vermelhas.
- Com um só ano escolhido, compara automaticamente com o **ano anterior**.
- Botão **Análise** escolhe o que mostrar (vendas, rentabilidade, quantidades, acumulado, desvio); detalhe por **cliente**, **fornecedor** ou **artigo**.
- Validada contra o ecrã "Comparação Anual em Euros" do PHC (valores iguais).

### 4.6 Consultas (7 folhas "Consulta …")
- Uma por tabela do modelo: tabela dinâmica em lista (todas as colunas, sem subtotais) com os mesmos filtros do dashboard correspondente.
- **Criadas só quando se abrem** (para a sincronização ser rápida) e abertas no **ano mais recente**; avisa e corrige se a lista não couber na folha.

### 4.7 Pormenores de interface
- **Seleção escondida**: ao mudar de folha, a célula ativa é colocada por baixo do maior gráfico visível, para não aparecer o retângulo verde de seleção por cima do dashboard (mantém a posição do ecrã; nas Consultas não interfere quando se está dentro da tabela).
- **Logótipo nas proporções originais** em todas as folhas, incluindo as Consultas criadas depois (a caixa, o logótipo e o nome da empresa são copiados do dashboard correspondente).
- **Ecrã de progresso** do add-in por cima do Excel enquanto os dashboards recalculam (em vez de a janela ficar preta ou a piscar).

### 4.8 Proteção
- Dashboards, Resumo e Consultas **sempre protegidos**: o utilizador só usa botões e filtros; gráficos e objetos não se mexem.
- **Estrutura do livro protegida**: não se inserem, movem, ocultam ou apagam folhas.
- Código VBA protegido por senha no pacote cliente; o add-in cliente é gerado a partir do "Mestre" com código **ofuscado** (sem comentários, identificadores renomeados).

---

## 5. O suplemento (aERP_Addin.xlam)

### Formulário de configuração (frmSQLManager)
- Separador **Configuração**: servidor, base de dados (lista), tipo de autenticação, utilizador, tempos limite. Botão **Testar** (valida sem deixar ligações "lixo") e **Guardar**.
- Ao Guardar: lê do PHC o **nome da empresa** e o **logótipo** (imagem guardada no PHC) e aplica-os em todas as folhas.
- Ao abrir, testa a última ligação; se estiver bem, abre logo no separador **Execução**. Só grava (e só limpa dados) se a configuração mudou.
- Separador **Execução**: lista das tabelas a sincronizar, **Executar selecionados** / **Executar todas**, com progresso por tabela num ecrã de espera (splash) que tapa o Excel.

### Sincronização (ProcessarAgendamentoSPs)
- Executa só as tabelas escolhidas; mostra "A carregar X (n de N)".
- Antes de carregar, apaga as tabelas dinâmicas das consultas (que tornavam o processo muito lento) e recria-as no fim, como marcadores leves.
- No fim: escolhe os dois anos mais recentes, atualiza filtros e campos vazios, vai para o Resumo.

### LimpaDados
- Esvazia as tabelas do Power Pivot (mantendo relações e medidas), com progresso.
- Usada ao **mudar de base de dados**, quando a **licença é inválida** e quando o livro é aberto **noutro PC** que não o que sincronizou (para não se distribuírem dados por engano).

### Licenciamento
- Cada arranque valida a licença numa Google Sheet (folhas **Lic** e **Log**) através de um **Google Apps Script** (web app): devolve válida/recusada/erro, validade e dias de tolerância offline.
- Regras: válida → funciona; recusada ou fora do prazo offline → limpa os dados; **modo demonstração** (sem ligação ao SQL) cria licença de experiência automaticamente.
- O **Log** regista data, PC, utilizador, versão e resultado.

---

## 6. Instalação e pacote cliente

- **Instalador Inno Setup** (`aERP_Setup_1.0.0.exe`), sem precisar de administrador:
  - copia aERP.xlsm, aERP_Addin.xlam e CodigosPostais.csv para `Documentos\aERP`;
  - só substitui o livro se o utilizador escolher (o add-in é sempre atualizado);
  - regista o add-in no Excel e cria um **Local Fidedigno** (para as macros não serem bloqueadas);
  - cria atalhos e pede para fechar o Excel; a desinstalação limpa o registo.
- **Geração do pacote**: macro que cria o add-in cliente ofuscado a partir do Mestre, copia o livro, e compila o instalador (ISCC). As senhas VBA são postas à mão pelo autor.
- Requisitos: Windows, **Excel 365/2021 64 bits** com Power Pivot, acesso de rede ao SQL Server do PHC, ligação à Internet para validar a licença (com tolerância offline).

---

## 7. Tecnologias usadas

- **Excel**: Power Pivot (modelo de dados, DAX), Power Query (M), funções CUBE (CUBEVALUE, CUBESET, CUBEMEMBER, CUBERANKEDMEMBER), segmentações OLAP, tabelas dinâmicas sobre o modelo, matrizes dinâmicas (FILTER, XLOOKUP, TAKE, LET).
- **VBA**: UserForms, ADO (SQL Server), eventos de livro (Open, SheetCalculate, ModelChange), `Application.OnTime`, manipulação de gráficos e formas, proteção.
- **SQL Server / PHC**: consultas SELECT por tabela.
- **Google Apps Script + Google Sheets**: licenças e registo de utilização.
- **Inno Setup**: instalador.

### Módulos VBA do livro (principais)
| Módulo | Função |
|---|---|
| modAnoSync | Sincronização dos anos, títulos "2026 vs 2025", gráficos mensais |
| modAnosGraficos | Segunda série/anel em todos os gráficos com 2 anos |
| modFiltroLista + frmFiltroLista | Filtros em lista com escolha múltipla e pesquisa |
| modFiltrosBD | Atualização dos filtros após sincronizar / mudar de BD |
| modConsultas, modConsultasFiltros, modBotoesConsulta | Folhas de consulta (criação sob pedido) |
| modComparacaoMedidas / Clientes / Artigos / Compras | Comparação anual |
| modResumo | Folha #aERP (visão geral e navegação) |
| modVazios | Esconder campos sem dados e reorganizar |
| modProtecao, modProtecaoEstado, modEstrutura | Proteção de folhas, segmentações e estrutura |
| modDados | Mostrar/ocultar as folhas técnicas "Dados …" |
| modComparacaoClientesVL / FornecedoresCL | Filtros de cliente/fornecedor da comparação anual nas folhas de linhas |
| modNavegacao | Botões de navegação entre folhas e "Voltar ao topo" |
| modLogoProporcao | Logótipo nas proporções originais em todas as folhas |
| modLegendasDonut | Legenda completa nos gráficos circulares |
| modSelecaoOculta | Esconde o retângulo da célula ativa debaixo de um gráfico |

### Módulos do add-in (Mestre)
| Módulo | Função |
|---|---|
| frmSQLManager | Formulário de configuração e execução da sincronização |
| modCore_Addin | Arranque do add-in, ecrã de progresso, logótipo, limpeza de dados |
| modCore_Dados | Ligação ao modelo, registo (log), consultas Power Query |
| modCore_Folhas | Extração por ADO para CSV temporário e ligação das consultas ao CSV |
| modCore_CP | Pasta do CodigosPostais.csv em cada PC |
| modGoogleSheets | Licenças (Apps Script), lista de SPs, modo sem ligação |
| modSeguranca | Cifra da configuração e identificação do PC |
| CAppEvents, clsAddinContext, clsWorkbookProps | Eventos do Excel, contexto por livro, propriedades do livro |

### Documentação no código
- Desde 16/11/2026 **todas as rotinas** do livro (216) e do add-in Mestre têm um **cabeçalho de comentário** em português: nome, o que faz, parâmetros e o que devolve; cada módulo tem uma descrição no topo.
- A inserção foi feita por macro e verificada: as linhas de código ficaram exatamente iguais (só se acrescentaram comentários).
- O add-in cliente gerado a partir do Mestre continua **sem comentários** (o gerador retira-os ao ofuscar).

---

## 8. Limitações conhecidas e trabalho futuro

- Instalação ainda **não testada num PC de cliente** (só no PC de desenvolvimento e numa VM). Instalador atual: aERP_Setup_1.0.0.exe de 16/11/2026.
- Algumas medidas de tempo antigas (YTD/LY/MoM) dependiam de uma coluna inexistente no Calendário e de uma relação MovVendas–Calendário em falta; o Calendário termina em 2026 e as consultas tinham datas fixas até 2026 (rever para 2027).
- Linhas sem código de artigo (texto livre, ex.: faturas de imobilizado) eram excluídas das consultas de linhas; linhas de venda com preço "IVA incluído" traziam IVA no valor líquido (corrigido em compras; rever em vendas).
- Medidas duplicadas no modelo (ex.: Total Vendas = Vendas Líquidas) a limpar.
- "Clientes/Fornecedores ativos 12M" dependem da data de hoje.
- Fornecedores sem tipo/zona na base de origem → gráficos vazios (escondidos automaticamente).
- Documentos de linhas sem cabeçalho (períodos diferentes nas consultas) produzem "(em branco)" em alguns filtros.
- As stored procedures definidas na folha SP são recriadas a cada Guardar mas não são usadas (a extração usa SELECT direto) — simplificar.
- O botão "Atualizar tudo" do friso não é desativado (só as ligações ficam bloqueadas).
- Consultas de linhas com milhões de registos só funcionam por ano.

### ⚠️ Antes de publicar no GitHub
- **Não publicar** o vbaProject com chaves ou tokens: versões antigas do add-in tinham credenciais embutidas (conta de serviço Google e um token GitHub). Confirmar que foram **revogadas** e que não existem no código a publicar.
- Retirar propriedades personalizadas do livro (ligação SQL cifrada, chave de API, licença, ID da base de dados, nome/logótipo da empresa) e usar **dados de demonstração**.
- Retirar a senha das folhas (constante no módulo de proteção) do código publicado.
- Publicar o código VBA como ficheiros `.bas/.frm/.cls` exportados (legíveis e versionáveis) em vez de apenas os binários.
- Definir licença do repositório (ex.: proprietária/"source available" ou MIT) conforme o modelo de negócio.

---

## 9. Cronologia (2026)

| Data | Marco |
|---|---|
| 9–11/10 | Construção inicial: modelo Power Pivot alimentado pelo PHC e 7 dashboards. |
| 11/10 | Análise completa do livro (problemas de dados, medidas, Resumo desatualizado). Folhas "Dados …" organizadas. |
| 13/10 | Códigos postais CTT → filtros e gráficos por **Distrito/Concelho**. |
| 15–16/10 | Add-in: progresso visível na limpeza e na sincronização; correções no teste de ligação. |
| 16–17/10 | Donuts de zona, KPIs arredondados, **painel de filtros misto**, sincronização do Ano entre dashboards. |
| 18/10 | **Nome e logótipo da empresa lidos do PHC**; nova barra lateral em todas as folhas. |
| 20/10 | Licença inválida → limpa dados; "#aERP" passa a Resumo com navegação. |
| 23–24/10 | Gráficos com 12 meses, **comparação de 2 anos em todos os gráficos**, tops pelo ano mais recente, novos gráficos de margem, reorganização dos dashboards, campos vazios automáticos. |
| 24–25/10 | **Comparação anual** ao estilo PHC (vendas, compras, por cliente/artigo/fornecedor). |
| 25/10 | **7 folhas de Consulta** com tabelas dinâmicas. |
| 28/10 | Sincronização acelerada (de bloqueio para ~1,5 min). |
| 30/10 | Botões no topo, Visão geral no Resumo, grande limpeza de código e nomes. |
| 31/10 | Análise de licenças, **Apps Script**, add-in Mestre vs cliente ofuscado, **instalador Inno Setup**. |
| 1–2/11 | Modo demonstração; dados sem janela de credenciais ("dados no livro"). |
| 3/11 | Códigos postais externos (livro 16 → 5,5 MB), **proteção das folhas**, consultas mais rápidas e criadas sob pedido. |
| 4/11 | Carga por **CSV temporário** (sem limite de linhas). |
| 5/11 | Ecrã preto ao mudar de ano corrigido; livro passa a chamar-se **aERP.xlsm**; backups arrumados. |
| 6/11 | Proteção sempre ativa; segmentações não se movem. |
| 7/11 | Botão **Sincronizar dados**; configuração só grava se mudar. |
| 8–10/11 | Anos comuns a todas as folhas e sincronização agendada (corrige bloqueios). |
| 11/11 | Filtros ao mudar de BD, pesquisa nos filtros, acumulados e desvios, mensagens no splash em vez de caixas OK, logótipo igual em todas as folhas. |
| 12/11 | **Estrutura do livro protegida**; comparação com o ano anterior; versão e contactos no formulário e no Resumo. |
| 13/11 | Livro aberto noutro PC é limpo automaticamente; instalador 1.0.0 recompilado. |
| 14/11 | Documento de visão geral; **vídeo explicativo automático** (macro que percorre o livro com narração por voz sintética). |
| 15/11 | Vídeo: gravação com OBS, narração revista, sem mudanças de ano durante a gravação. |
| 16/11 | Logótipo nas proporções certas nas Consultas; legenda completa nos circulares; seleção escondida; **cabeçalhos de comentário em todo o código**; pacote cliente e instalador refeitos. |

---

## 10. Guião sugerido para o vídeo explicativo

**Duração alvo:** 6–10 minutos. Usar sempre **dados de demonstração**.

1. **Abertura (30 s)** — O problema: os dados do PHC estão lá, mas tirar indicadores exige relatórios ou exportações manuais. O aERP: BI dentro do Excel, pronto a usar.
2. **Instalação (45 s)** — Correr o instalador, abrir o atalho aERP. Mostrar que não pede administrador e que as macros já estão autorizadas.
3. **Configuração (1 min)** — Botão Configuração → servidor, base de dados, Testar, Guardar. O logótipo e o nome da empresa aparecem automaticamente.
4. **Sincronizar (45 s)** — Botão Sincronizar dados; ecrã de progresso tabela a tabela; chegada ao Resumo.
5. **Resumo e navegação (45 s)** — Visão geral, filtro de Ano, botões para cada dashboard.
6. **Dashboard de Vendas (2 min)** — KPIs, vendas e margem por mês, top 10 clientes, distrito/concelho. Mudar para **2 anos** e mostrar todos os gráficos a comparar. Usar um filtro em lista com **Procurar**. Limpar filtros.
7. **Comparação anual (1 min)** — Tabela ao estilo PHC: acumulado, desvios, setas; botão Análise; detalhe por cliente.
8. **Outros dashboards (1 min)** — Vendas por artigo (navegação "Ver detalhe por artigo"), Compras, Clientes, Fornecedores, Artigos (passagem rápida).
9. **Consultas (45 s)** — Abrir uma Consulta: lista completa com os mesmos filtros; exportar/copiar.
10. **Segurança (30 s)** — Folhas protegidas, licença, dados limpos se o livro for copiado para outro PC.
11. **Fecho (20 s)** — Requisitos, contacto, onde está o projeto (GitHub).

**Planos a gravar:** ecrã inteiro 1920×1080, zoom do Excel a 100%; destacar o cursor; legendas em português; música discreta.

---

### Vídeo já produzido (14–15/11)
- Em vez de gravar a voz do autor, o vídeo é **gerado automaticamente**: um livro auxiliar (`aERP_Video.xlsm`, macro `modVideoAERP`) abre o aERP e percorre-o passo a passo — cartões de título, navegação, filtros, formulário de configuração — enquanto toca a narração de cada passo.
- **Narração** em português de Portugal com voz sintética (33 segmentos, ≈5,5 min, gerados com `edge-tts`); o texto está em `narracao.txt` e cada segmento pode ser regenerado sozinho.
- **Cartões de título** em janelas HTA sem moldura; dados sensíveis do formulário de configuração (servidor, utilizador, base de dados, PC) tapados por sobreposições com valores de demonstração.
- **Gravação** com OBS Studio (perfil "aERP Video", 30 fps, só o som do sistema) — a ferramenta de recorte do Windows desfasava som e imagem; legendas automáticas no Clipchamp.
- Para uma nova versão do vídeo: atualizar `narracao.txt`, regenerar os segmentos alterados, correr `EnsaioSemAudio` e depois `GravarVideoAERP` com o OBS a gravar.

## 11. Estrutura sugerida para o repositório GitHub

```
aERP/
├── README.md                 ← resumo, capturas, requisitos, instalação rápida
├── docs/
│   ├── arquitetura.md        ← secções 2 e 3 deste documento
│   ├── funcionalidades.md    ← secção 4
│   ├── addin.md              ← secção 5
│   ├── instalacao.md         ← secção 6
│   └── imagens/              ← capturas dos dashboards (dados demo)
├── src/
│   ├── livro/                ← módulos VBA exportados do aERP.xlsm (.bas/.cls/.frm)
│   ├── addin/                ← módulos VBA do add-in Mestre (sem segredos)
│   ├── powerquery/           ← código M das consultas
│   ├── dax/                  ← lista de medidas
│   ├── sql/                  ← consultas de extração do PHC
│   ├── apps-script/          ← aERP_Licencas.gs (sem identificadores reais)
│   └── video/                ← modVideoAERP.bas, narracao.txt, gerar_audio.py
├── instalador/
│   └── aERP_Setup.iss
├── CHANGELOG.md              ← secção 9
└── LICENSE
```

**Secções do README:** O que é · Capturas · Funcionalidades principais · Requisitos · Instalação · Primeira configuração · Como funciona (diagrama) · Limitações · Roadmap · Licença · Contacto.
