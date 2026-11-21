"""
Gerador da base fictícia "Gestão de Licenças Microsoft 365" — Grupo Aurora Mídia (empresa fictícia).
Gera Licencas_M365_Base.xlsx com tabelas dimensão (d*) e fato (f*) em esquema estrela, prontas para o Power BI.
Todos os dados, nomes e preços são fictícios. Semente fixa: o resultado é sempre o mesmo.
Uso: python gerar_base.py   (requer pandas, openpyxl, python-dateutil)
"""
import random, math, datetime as dt
from dateutil.relativedelta import relativedelta
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

random.seed(2026)
D = dt.date
REF = D(2026, 10, 31)                       # data de referência dos indicadores
MESES = [D(2025, 1, 1) + relativedelta(months=i) for i in range(22)]   # jan/2025 a out/2026
def fim_mes(m): return m + relativedelta(months=1) - dt.timedelta(days=1)
def rdate(a, b): return a + dt.timedelta(days=random.randint(0, max(0, (b - a).days)))

# ---------------- dSKU (preços mensais de referência, fictícios)
skus = [
 ("E3",  "Microsoft 365 E3",              "Licença base",   190, 0.03),
 ("E5",  "Microsoft 365 E5",              "Licença base",   300, 0.04),
 ("F3",  "Microsoft 365 F3",              "Licença base",    45, 0.04),
 ("EXK", "Exchange Online Kiosk",         "Licença base",    12, 0.10),
 ("COP", "Microsoft 365 Copilot",         "Add-on IA",      165, 0.02),
 ("PBI", "Power BI Pro",                  "Analytics",       75, 0.08),
 ("PPU", "Power BI Premium por usuário",  "Analytics",      125, 0.10),
 ("VIS", "Visio Plano 2",                 "Produtividade",   80, 0.25),
 ("PJ3", "Project Plano 3",               "Produtividade",  160, 0.20),
 ("PJ5", "Project Plano 5",               "Produtividade",  280, 0.15),
 ("TPR", "Teams Premium",                 "Comunicação",     55, 0.10),
 ("TPH", "Teams Phone Standard",          "Comunicação",     45, 0.06),
 ("PAP", "Power Apps Premium",            "Low-code",       105, 0.15),
 ("PAU", "Power Automate Premium",        "Low-code",        80, 0.15)]
dSKU = pd.DataFrame([dict(ID_SKU=a, Produto=b, Categoria=c, Preco_Mensal=p, Folga_Alvo=f) for a, b, c, p, f in skus])
PRECO = dict(zip(dSKU.ID_SKU, dSKU.Preco_Mensal))

# ---------------- dUnidade  (headcount, % executivo, % linha de frente, add-ons)
unid = [
 ("U01", "Jornalismo",                 2200, .02, .12, {"PBI": .04, "TPR": .05, "VIS": .01}),
 ("U02", "Produção e Estúdios",        3500, .01, .70, {"PJ3": .02, "PBI": .01}),
 ("U03", "Streaming e Digital",        1800, .02, .05, {"PBI": .20, "PAP": .07, "PAU": .06, "VIS": .05, "PPU": .02}),
 ("U04", "Comercial e Publicidade",    1200, .03, .05, {"PBI": .25, "TPH": .30, "TPR": .08}),
 ("U05", "TV Aberta e Programação",    1100, .02, .45, {"PJ3": .04, "PBI": .03}),
 ("U06", "Tecnologia",                  900, .03, .02, {"VIS": .30, "PJ3": .10, "PJ5": .04, "PAP": .20, "PAU": .22, "PBI": .30, "PPU": .05}),
 ("U07", "Corporativo (Fin, RH, Jur)",  800, .06, .03, {"PBI": .30, "PJ3": .05, "PPU": .03, "TPH": .10}),
 ("U08", "Rádio",                       500, .02, .50, {"TPH": .05})]
diretoria = {"U01": "Conteúdo", "U02": "Conteúdo", "U05": "Conteúdo", "U08": "Conteúdo",
             "U03": "Digital", "U04": "Negócios", "U06": "Tecnologia", "U07": "Corporativo"}
dUni = pd.DataFrame([dict(ID_Unidade=u, Unidade=n, Diretoria=diretoria[u], Headcount_Planejado=h) for u, n, h, *_ in unid])

# ---------------- dColaboradores
ufs, upeso = ["RJ", "SP", "DF", "MG", "PE", "RS", "BA"], [52, 28, 6, 5, 4, 3, 2]
tipos, tpeso = ["CLT", "Terceiro", "Estagiário"], [80, 15, 5]
cols = []; n = 1
for u, nome, h, pexe, pfr, addons in unid:
    for _ in range(h):
        r = random.random()
        persona = "Executivo" if r < pexe else ("Linha de frente" if r < pexe + pfr else "Conhecimento")
        adm = rdate(D(2008, 1, 1), D(2026, 9, 30))
        status, desl = "Ativo", None
        if random.random() < 0.075 and adm < D(2026, 5, 1):
            desl = rdate(max(adm + dt.timedelta(days=90), D(2025, 1, 1)), D(2026, 10, 25)); status = "Desligado"
        cols.append(dict(ID_Colaborador=f"C{n:05d}", UPN=f"u{n:05d}@auroramidia.example", ID_Unidade=u, Persona=persona,
                         Vinculo=random.choices(tipos, tpeso)[0], UF=random.choices(ufs, upeso)[0],
                         Data_Admissao=adm, Status=status, Data_Desligamento=desl)); n += 1
# contas de sala/serviço (caixas compartilhadas)
for k in range(160):
    cols.append(dict(ID_Colaborador=f"C{n:05d}", UPN=f"sala{k+1:03d}@auroramidia.example", ID_Unidade=random.choice(["U01", "U02", "U05", "U07"]),
                     Persona="Conta de serviço", Vinculo="Conta de serviço", UF="RJ", Data_Admissao=rdate(D(2015, 1, 1), D(2025, 6, 1)),
                     Status="Ativo", Data_Desligamento=None)); n += 1
dCol = pd.DataFrame(cols)

# ---------------- fLicencas (uma linha por atribuição)
ADD = {u: a for u, *_, a in unid}
lic = []; j = 1
def atrib(c, sku, ini, risco_ocioso):
    global j
    status, rem = "Ativa", None
    if c["Status"] == "Desligado":
        if random.random() < 0.86:
            status, rem = "Removida", c["Data_Desligamento"] + dt.timedelta(days=random.choice([0, 0, 1, 2, 5, 9, 15, 30]))
    elif random.random() < 0.03 and sku not in ("E3", "E5", "F3", "EXK"):
        status, rem = "Removida", rdate(ini, REF)
    if rem and rem > REF: status, rem = "Ativa", None
    if ini > (rem or REF): ini = (rem or REF) - dt.timedelta(days=20)
    if status == "Removida": ua = rem - dt.timedelta(days=random.randint(0, 25))
    elif c["Status"] == "Desligado": ua = c["Data_Desligamento"] - dt.timedelta(days=random.randint(0, 3))
    else:
        r = random.random()
        ua = REF - dt.timedelta(days=random.randint(0, 12) if r > risco_ocioso else random.randint(31, 240))
    ua = max(ua, ini)
    lic.append(dict(ID_Atribuicao=f"L{j:06d}", ID_Colaborador=c["ID_Colaborador"], ID_SKU=sku, Data_Atribuicao=ini,
                    Data_Remocao=rem, Status=status, Data_Ultima_Atividade=ua,
                    Origem=random.choices(["Grupo dinâmico", "Atribuição direta"], [70, 30])[0])); j += 1

for c in cols:
    base_ini = max(c["Data_Admissao"], D(2022, 1, 1)) if c["Data_Admissao"] < D(2022, 1, 1) else c["Data_Admissao"]
    p = c["Persona"]
    if p == "Conta de serviço":
        atrib(c, "EXK", base_ini, 0.10); continue
    if p == "Executivo": base, risco = "E5", 0.02
    elif p == "Conhecimento": base, risco = ("E5" if random.random() < 0.04 else "E3"), 0.06
    else:  # linha de frente: legado com E3 em parte do público
        if random.random() < 0.24: base, risco = "E3", 0.20
        else: base, risco = "F3", 0.07
    atrib(c, base, base_ini, risco)
    # sobreposição E3 + E5 (upgrade sem remover a E3)
    if base == "E5" and random.random() < 0.09:
        atrib(c, "E3", base_ini, 0.95)
    if p == "Linha de frente": continue
    for sku, pr in ADD[c["ID_Unidade"]].items():
        if random.random() < pr:
            atrib(c, sku, rdate(max(base_ini, D(2023, 1, 1)), D(2026, 10, 15)), 0.12)
    # piloto de Copilot: começa em mar/2025 e cresce
    pc = .55 if p == "Executivo" else {"U06": .14, "U07": .12, "U03": .10, "U04": .09, "U01": .07}.get(c["ID_Unidade"], .02)
    if random.random() < pc:
        d0 = D(2025, 3, 1) + dt.timedelta(days=int(600 * random.random() ** 0.7))
        atrib(c, "COP", max(d0, base_ini), 0.30)
fLic = pd.DataFrame(lic)

# ---------------- fSnapshotMensal (atribuídas e ativas nos últimos 30 dias, por mês/SKU/unidade)
uni_col = dict(zip(dCol.ID_Colaborador, dCol.ID_Unidade))
L = fLic.assign(ID_Unidade=fLic.ID_Colaborador.map(uni_col))
L["ini"] = pd.to_datetime(L.Data_Atribuicao); L["rem"] = pd.to_datetime(L.Data_Remocao)
atual = L[L.Status == "Ativa"].copy()
atual["ativa30"] = (pd.Timestamp(REF) - pd.to_datetime(atual.Data_Ultima_Atividade)).dt.days <= 30
taxa_final = atual.groupby("ID_SKU").ativa30.mean().to_dict()
snap = []
for m in MESES:
    fm = pd.Timestamp(fim_mes(m))
    vig = L[(L.ini <= fm) & (L.rem.isna() | (L.rem > fm))]
    g = vig.groupby(["ID_SKU", "ID_Unidade"]).size()
    for (sku, u), qt in g.items():
        if m == MESES[-1]:
            at = int(atual[(atual.ID_SKU == sku) & (atual.ID_Unidade == u)].ativa30.sum())
        else:
            tx = taxa_final.get(sku, .85) + (0.04 if sku in ("E3", "E5", "F3") else 0) * random.uniform(-1, 1)
            if sku == "COP": tx = min(.95, tx + 0.15 * (1 - MESES.index(m) / 21))
            at = min(qt, int(round(qt * max(0, min(1, tx)))))
        snap.append(dict(Mes=m, ID_SKU=sku, ID_Unidade=u, Atribuidas=int(qt), Ativas_30d=at))
fSnap = pd.DataFrame(snap)

# ---------------- fEstoqueSKU e fPedidos (CSP/NCE: aumento a qualquer momento, redução só no aniversário em março)
tot = fSnap.groupby(["Mes", "ID_SKU"]).Atribuidas.sum()
folga = dict(zip(dSKU.ID_SKU, dSKU.Folga_Alvo))
estoque, pedidos = [], []
cont = {}; nsc = {}
def novo_sc(data):
    nsc[data.year] = nsc.get(data.year, 0) + 1
    return f"SC-{data.year}-{nsc[data.year]:04d}"
def arred(q): return int(10 * math.ceil(q / 10)) if q >= 10 else int(q)
for m in MESES:
    for sku in dSKU.ID_SKU:
        need = int(tot.get((m, sku), 0))
        alvo = int(math.ceil(need * (1 + folga[sku])))
        atual_c = cont.get(sku)
        if atual_c is None:
            cont[sku] = arred(alvo)
        elif need > atual_c:
            q = arred(alvo - atual_c); data = m + dt.timedelta(days=random.randint(2, 20))
            pedidos.append(dict(Num_SC=novo_sc(data), Data=data, ID_SKU=sku, Tipo="Aumento", Qtde=q,
                                Valor_Mensal=q * PRECO[sku], Valor_Anual=q * PRECO[sku] * 12, Lead_Time_Dias=random.randint(1, 7)))
            cont[sku] = atual_c + q
        elif m.month == 3 and atual_c > alvo * 1.02 and random.random() < 0.6:   # ajuste na renovação
            q = arred(atual_c - alvo); data = m + dt.timedelta(days=random.randint(1, 10))
            if q > 0:
                pedidos.append(dict(Num_SC=novo_sc(data), Data=data, ID_SKU=sku, Tipo="Redução na renovação", Qtde=-q,
                                    Valor_Mensal=-q * PRECO[sku], Valor_Anual=-q * PRECO[sku] * 12, Lead_Time_Dias=random.randint(1, 5)))
                cont[sku] = atual_c - q
        estoque.append(dict(Mes=m, ID_SKU=sku, Contratadas=cont[sku]))
fEst = pd.DataFrame(estoque)
fPed = pd.DataFrame(pedidos)

# ---------------- fOrcamento (orçado mensal de licenças por unidade: 2025 sobre o histórico, 2026 com reajuste fixo)
custo_u = fSnap.assign(v=fSnap.Atribuidas * fSnap.ID_SKU.map(PRECO)).groupby(["Mes", "ID_Unidade"]).v.sum()
orc = []
for m in MESES + [D(2026, 11, 1), D(2026, 12, 1)]:
    for u in dUni.ID_Unidade:
        if m.year == 2025: v = custo_u.get((D(2025, 1, 1), u), 0) * 1.03
        else: v = custo_u.get((D(2025, 12, 1), u), 0) * 1.04
        orc.append(dict(Mes=m, ID_Unidade=u, Valor_Orcado=round(v, -2)))
fOrc = pd.DataFrame(orc)

tabelas = [("dUnidade", dUni), ("dSKU", dSKU.drop(columns="Folga_Alvo")), ("dColaboradores", dCol), ("fLicencas", fLic),
           ("fSnapshotMensal", fSnap), ("fEstoqueSKU", fEst), ("fPedidos", fPed), ("fOrcamento", fOrc)]

# ---------------- Excel
desc = {"dUnidade": "Unidades de negócio, diretoria e headcount planejado",
        "dSKU": "Catálogo de licenças Microsoft com preço mensal de referência (fictício, R$)",
        "dColaboradores": "Colaboradores e contas de serviço fictícios: unidade, persona, vínculo, UF e status",
        "fLicencas": "Atribuições de licença por usuário, com remoção e data da última atividade",
        "fSnapshotMensal": "Fotografia mensal: licenças atribuídas e ativas (30 dias) por SKU e unidade",
        "fEstoqueSKU": "Quantidade contratada por SKU em cada mês",
        "fPedidos": "Pedidos de licença (aumentos e reduções na renovação anual)",
        "fOrcamento": "Orçamento mensal de licenças por unidade (2025–2026)"}
wb = Workbook(); ws = wb.active; ws.title = "LEIA-ME"
F = Font(name="Arial", size=10)
linhas = [["Gestão de Licenças Microsoft 365 · Grupo Aurora Mídia (empresa fictícia)"], [""],
          ["Base 100% fictícia, gerada por script (gerar_base.py). Nomes, números e preços não representam nenhuma empresa real."],
          ["Período: jan/2025 a out/2026 (orçamento até dez/2026). Data de referência dos indicadores: 31/10/2026."],
          ["Cada aba tem uma Tabela do Excel com o mesmo nome. No Power BI: Obter dados > Excel > marque as tabelas."], [""],
          ["Tabela", "Linhas", "Conteúdo"]]
for r in linhas: ws.append(r)
for nme, t in tabelas: ws.append([nme, len(t), desc[nme]])
ws.append([""]); ws.append(["Relacionamentos (1 → *)"])
for r in ["dUnidade[ID_Unidade] → dColaboradores, fSnapshotMensal, fOrcamento",
          "dColaboradores[ID_Colaborador] → fLicencas",
          "dSKU[ID_SKU] → fLicencas, fSnapshotMensal, fEstoqueSKU, fPedidos",
          "dCalendario[Date] → fSnapshotMensal[Mes], fEstoqueSKU[Mes], fOrcamento[Mes], fPedidos[Data]  (dCalendario criada em DAX)"]:
    ws.append([r])
for row in ws.iter_rows():
    for c in row: c.font = F
ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E79")
for c in ws[7]: c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E79")
ws.cell(row=9 + len(tabelas), column=1).font = Font(name="Arial", size=10, bold=True)
ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 10; ws.column_dimensions["C"].width = 100
for nme, t in tabelas:
    s = wb.create_sheet(nme); s.append(list(t.columns))
    for rec in t.itertuples(index=False):
        s.append([None if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT else (v.date() if isinstance(v, pd.Timestamp) else v) for v in rec])
    tab = Table(displayName=nme, ref=f"A1:{get_column_letter(len(t.columns))}{len(t) + 1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True); s.add_table(tab)
    for k, col in enumerate(t.columns, 1):
        Lc = get_column_letter(k)
        fmt = '"R$" #,##0.00' if ("Valor" in col or "Preco" in col) else ("dd/mm/yyyy" if col.startswith("Data") or col == "Mes" else None)
        s.column_dimensions[Lc].width = min(max(len(col) + 2, *(len(str(x)) + 2 for x in t[col].head(200)), 10), 42)
        for c in s[Lc]:
            c.font = F
            if fmt and c.row > 1: c.number_format = fmt
        s[f"{Lc}1"].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    s.freeze_panes = "A2"
wb.save("Licencas_M365_Base.xlsx")
for nme, t in tabelas: print(f"{nme}: {len(t)} linhas")
