# Modelo e medidas DAX

Guia de construção do relatório no Power BI Desktop a partir de `dados/Licencas_M365_Base.xlsx`.

## 1. Importação

**Obter dados › Excel** e marque as 8 tabelas (ícone de tabela, não as abas).
No Power Query, confira os tipos: colunas `Data_*` e `Mes` como **Data**, `Preco_Mensal`, `Valor_*` como **Número decimal**, quantidades como **Número inteiro**.

## 2. Tabela calendário

**Modelagem › Nova tabela**:

```DAX
dCalendario =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2025, 1, 1 ), DATE ( 2026, 12, 31 ) ),
    "Ano", YEAR ( [Date] ),
    "Mês", MONTH ( [Date] ),
    "Mês Nome", FORMAT ( [Date], "mmm/yy" ),
    "AnoMês", YEAR ( [Date] ) * 100 + MONTH ( [Date] )
)
```

Marque como **tabela de datas** (coluna `Date`) e classifique `Mês Nome` por `AnoMês`.

## 3. Relacionamentos (1 → *, filtro em direção única)

| De (1) | Para (*) |
|---|---|
| `dUnidade[ID_Unidade]` | `dColaboradores`, `fSnapshotMensal`, `fOrcamento` |
| `dColaboradores[ID_Colaborador]` | `fLicencas` |
| `dSKU[ID_SKU]` | `fLicencas`, `fSnapshotMensal`, `fEstoqueSKU`, `fPedidos` |
| `dCalendario[Date]` | `fSnapshotMensal[Mes]`, `fEstoqueSKU[Mes]`, `fOrcamento[Mes]`, `fPedidos[Data]` |

> `fEstoqueSKU` (contratadas) é do tenant inteiro: não se relaciona com unidade. Use as medidas de contratação em visuais por SKU.

## 4. Medidas

Crie uma tabela vazia `_Medidas` para organizar. A data de referência fixa mantém o case reproduzível; em produção, troque por `TODAY()`.

### Base

```DAX
Data Referência = DATE ( 2026, 10, 31 )
```

### Fotografia mensal (semiaditivas)

Quantidade de licenças é uma **fotografia**: somar janeiro + fevereiro não faz sentido. As medidas pegam o último mês do contexto.

```DAX
Licenças Atribuídas =
VAR UltimoMes = MAX ( fSnapshotMensal[Mes] )
RETURN
    CALCULATE ( SUM ( fSnapshotMensal[Atribuidas] ), fSnapshotMensal[Mes] = UltimoMes )

Licenças Ativas 30d =
VAR UltimoMes = MAX ( fSnapshotMensal[Mes] )
RETURN
    CALCULATE ( SUM ( fSnapshotMensal[Ativas_30d] ), fSnapshotMensal[Mes] = UltimoMes )

Taxa de Atividade = DIVIDE ( [Licenças Ativas 30d], [Licenças Atribuídas] )

Licenças Contratadas =
VAR UltimoMes = MAX ( fEstoqueSKU[Mes] )
RETURN
    CALCULATE ( SUM ( fEstoqueSKU[Contratadas] ), fEstoqueSKU[Mes] = UltimoMes )

Licenças Livres = [Licenças Contratadas] - [Licenças Atribuídas]

Ocupação = DIVIDE ( [Licenças Atribuídas], [Licenças Contratadas] )

Alerta de Ocupação =
SWITCH (
    TRUE (),
    [Ocupação] >= 0.98, "Risco de falta",
    [Ocupação] < 0.85, "Excesso contratado",
    "Ok"
)
```

### Custo, rateio e orçamento (aditivas)

O custo mensal soma entre meses: é o que permite rateio, acumulado do ano e comparação com o orçado.

```DAX
Custo Rateado =
SUMX ( fSnapshotMensal, fSnapshotMensal[Atribuidas] * RELATED ( dSKU[Preco_Mensal] ) )

Custo Contratado =
SUMX ( fEstoqueSKU, fEstoqueSKU[Contratadas] * RELATED ( dSKU[Preco_Mensal] ) )

Orçado = SUM ( fOrcamento[Valor_Orcado] )

Desvio Orçamentário = [Custo Rateado] - [Orçado]

% Execução = DIVIDE ( [Custo Rateado], [Orçado] )

Custo YTD = TOTALYTD ( [Custo Rateado], dCalendario[Date] )

Orçado YTD = TOTALYTD ( [Orçado], dCalendario[Date] )

Custo Ano Anterior = CALCULATE ( [Custo Rateado], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )

Variação YoY % = DIVIDE ( [Custo Rateado] - [Custo Ano Anterior], [Custo Ano Anterior] )
```

### Otimização (estado atual, por usuário)

```DAX
Atribuições Ativas =
CALCULATE ( COUNTROWS ( fLicencas ), fLicencas[Status] = "Ativa" )

Licenças Ociosas =
VAR Ref = [Data Referência]
RETURN
    CALCULATE ( [Atribuições Ativas], fLicencas[Data_Ultima_Atividade] < Ref - 30 )

% Ociosas = DIVIDE ( [Licenças Ociosas], [Atribuições Ativas] )

Custo Ocioso Mensal =
VAR Ref = [Data Referência]
RETURN
    SUMX (
        FILTER (
            fLicencas,
            fLicencas[Status] = "Ativa"
                && fLicencas[Data_Ultima_Atividade] < Ref - 30
        ),
        RELATED ( dSKU[Preco_Mensal] )
    )

Licenças de Desligados =
CALCULATE ( [Atribuições Ativas], dColaboradores[Status] = "Desligado" )

Usuários com E3 e E5 =
COUNTROWS (
    FILTER (
        VALUES ( dColaboradores[ID_Colaborador] ),
        CALCULATE (
            DISTINCTCOUNT ( fLicencas[ID_SKU] ),
            fLicencas[Status] = "Ativa",
            fLicencas[ID_SKU] IN { "E3", "E5" }
        ) = 2
    )
)

Linha de Frente com E3 =
CALCULATE (
    [Atribuições Ativas],
    fLicencas[ID_SKU] = "E3",
    dColaboradores[Persona] = "Linha de frente"
)

Economia Potencial Mensal =
VAR Ref = [Data Referência]
VAR PrecoE3 = LOOKUPVALUE ( dSKU[Preco_Mensal], dSKU[ID_SKU], "E3" )
VAR PrecoF3 = LOOKUPVALUE ( dSKU[Preco_Mensal], dSKU[ID_SKU], "F3" )
VAR Remover =
    SUMX (
        FILTER (
            fLicencas,
            fLicencas[Status] = "Ativa"
                && ( fLicencas[Data_Ultima_Atividade] < Ref - 30
                    || RELATED ( dColaboradores[Status] ) = "Desligado" )
        ),
        RELATED ( dSKU[Preco_Mensal] )
    )
VAR Rebaixar =
    COUNTROWS (
        FILTER (
            fLicencas,
            fLicencas[Status] = "Ativa"
                && fLicencas[ID_SKU] = "E3"
                && fLicencas[Data_Ultima_Atividade] >= Ref - 30
                && RELATED ( dColaboradores[Persona] ) = "Linha de frente"
                && RELATED ( dColaboradores[Status] ) = "Ativo"
        )
    ) * ( PrecoE3 - PrecoF3 )
RETURN
    Remover + Rebaixar

Economia Potencial Anual = [Economia Potencial Mensal] * 12
```

> A economia não conta duas vezes: quem é desligado **e** ocioso entra uma vez em "Remover"; "Rebaixar" considera só E3 em uso por ativos da linha de frente.

### Copilot e pedidos

```DAX
Copilot Atribuídas = CALCULATE ( [Licenças Atribuídas], dSKU[ID_SKU] = "COP" )

Adoção Copilot = CALCULATE ( [Taxa de Atividade], dSKU[ID_SKU] = "COP" )

Valor Anual Pedidos = SUM ( fPedidos[Valor_Anual] )

Lead Time Médio (dias) = AVERAGE ( fPedidos[Lead_Time_Dias] )
```

## 5. Páginas do relatório

| Página | Pergunta que responde | Visuais |
|---|---|---|
| **Visão Executiva** | Quanto custa, quanto dá para economizar, estamos no orçamento? | Cartões (Custo Rateado, Economia Potencial Anual, % Execução, Taxa de Atividade), linha Custo × Orçado por mês, barras por diretoria |
| **Ocupação por SKU** | Onde falta e onde sobra licença? | Matriz SKU × Contratadas/Atribuídas/Livres/Ocupação/Atividade com formatação condicional pelo Alerta |
| **Otimização** | O que remover ou rebaixar, e de quem? | Cartões das 4 frentes, barras por unidade, tabela de desligados com licença (UPN, SKU, data) |
| **Copilot** | O piloto está sendo usado? | Linha de atribuídas × ativas por mês, adoção por unidade |
| **Rateio e Orçamento** | Quanto cada unidade consome e quanto fugiu do plano? | Matriz unidade × mês, desvio, YTD e YoY |
| **Pedidos** | Como o estoque foi ajustado? | Linha do tempo de aumentos e reduções, lead time |

Use a mesma ordem de visuais em todas as páginas: cartões no topo, tendência no meio e detalhe embaixo.
