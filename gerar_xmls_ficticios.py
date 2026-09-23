"""
Gera um lote de XMLs de NF-e 100% FICTÍCIOS (layout simplificado da NF-e 4.00)
para a Atlas Distribuidora de Suprimentos Industriais Ltda (empresa inventada).

Cria compras (entradas), vendas (saídas) e devoluções de venda de jan a dez/2025,
além de alguns problemas de propósito (nota duplicada, nota não autorizada,
chave com dígito verificador errado) para o analisador detectar.

Uso:  python gerar_xmls_ficticios.py  [pasta_saida]
"""
import os, sys, random
from datetime import datetime, timedelta
from xml.sax.saxutils import escape

random.seed(42)
OUT = sys.argv[1] if len(sys.argv) > 1 else "xmls"
os.makedirs(OUT, exist_ok=True)

EMPRESA = {"cnpj": "11222333000181", "nome": "Atlas Distribuidora de Suprimentos Industriais Ltda", "uf": "SP"}
FORNECEDORES = [
    {"cnpj": "22333444000172", "nome": "Fornecedor Alfa EPI Ltda", "uf": "SP"},
    {"cnpj": "33444555000163", "nome": "Fornecedor Beta Abrasivos Ltda", "uf": "MG"},
    {"cnpj": "44555666000154", "nome": "Fornecedor Gama Quimicos Ltda", "uf": "PR"},
]
CLIENTES = [{"cnpj": f"{55000000000100 + i*1011:014d}", "nome": f"Cliente Industrial {i:02d} Ltda",
             "uf": random.choice(["SP"]*6 + ["MG", "PR", "GO", "RJ"])} for i in range(1, 31)]

# código, descrição, NCM, unidade, custo, preço de venda, fornecedor, volume médio/mês
PRODUTOS = [
    ("P001", "Luva nitrilica cx 100un", "40151900", "CX", 38.0, 59.9, 0, 220),
    ("P002", "Oculos de protecao incolor", "90049020", "UN", 6.2, 12.9, 0, 400),
    ("P003", "Bota de seguranca bico PVC", "64029990", "PAR", 72.0, 109.0, 0, 90),
    ("P004", "Protetor auricular plug", "39269090", "UN", 0.9, 2.1, 0, 1500),
    ("P005", "Capacete de seguranca aba frontal", "65061000", "UN", 18.5, 27.9, 0, 120),
    ("P006", "Disco de corte 7 pol", "68042211", "UN", 7.8, 11.5, 1, 800),
    ("P007", "Lixa d'agua grao 220", "68052000", "FL", 1.1, 1.9, 1, 2500),
    ("P008", "Disco flap 4.1/2 pol", "68042211", "UN", 9.4, 10.9, 1, 600),   # margem apertada de propósito
    ("P009", "Lubrificante spray 300ml", "34031900", "UN", 14.0, 23.5, 2, 300),
    ("P010", "Desengraxante industrial 5L", "34022000", "GL", 46.0, 52.0, 2, 110),  # margem baixa
]
PIS, COFINS = 0.0065, 0.03       # Lucro Presumido, regime cumulativo
SAZONAL = [0.85, 0.9, 1.0, 1.05, 1.1, 1.0, 0.95, 1.05, 1.1, 1.15, 1.1, 0.8]

def dv_chave(chave43):
    pesos, soma = [2, 3, 4, 5, 6, 7, 8, 9], 0
    for i, d in enumerate(reversed(chave43)):
        soma += int(d) * pesos[i % 8]
    r = soma % 11
    return "0" if r < 2 else str(11 - r)

def chave(cuf, dt, cnpj, serie, nnf):
    c43 = f"{cuf}{dt:%y%m}{cnpj}55{serie:03d}{nnf:09d}1{random.randint(0, 99999999):08d}"
    return c43 + dv_chave(c43)

CUF = {"SP": "35", "MG": "31", "PR": "41", "GO": "52", "RJ": "33"}
def aliq_icms(uf_orig, uf_dest): return 0.18 if uf_orig == uf_dest else 0.12
def f2(v): return f"{v:.2f}"

def nota_xml(ch, dt, nnf, emit, dest, tpnf, finnfe, natop, itens, cstat="100"):
    dets, tot = [], {"vprod": 0, "vdesc": 0, "vbc": 0, "vicms": 0, "vpis": 0, "vcofins": 0}
    for n, it in enumerate(itens, 1):
        p = it["prod"]; vprod = round(it["qtd"] * it["vun"], 2); vdesc = round(vprod * it.get("desc", 0), 2)
        bc = vprod - vdesc; ali = it["icms"]; vicms = round(bc * ali, 2)
        bc_pc = bc - vicms                        # exclusão do ICMS da base de PIS/COFINS (Tema 69 STF)
        vpis = round(bc_pc * PIS, 2) if it["pc"] else 0.0
        vcof = round(bc_pc * COFINS, 2) if it["pc"] else 0.0
        for k, v in (("vprod", vprod), ("vdesc", vdesc), ("vbc", bc), ("vicms", vicms), ("vpis", vpis), ("vcofins", vcof)):
            tot[k] += v
        desc_tag = f"<vDesc>{f2(vdesc)}</vDesc>" if vdesc else ""
        dets.append(
            f'<det nItem="{n}"><prod><cProd>{p[0]}</cProd><xProd>{escape(p[1])}</xProd><NCM>{p[2]}</NCM>'
            f'<CFOP>{it["cfop"]}</CFOP><uCom>{p[3]}</uCom><qCom>{it["qtd"]:.4f}</qCom><vUnCom>{it["vun"]:.4f}</vUnCom>'
            f'<vProd>{f2(vprod)}</vProd>{desc_tag}</prod><imposto>'
            f'<ICMS><ICMS00><orig>0</orig><CST>00</CST><modBC>3</modBC><vBC>{f2(bc)}</vBC><pICMS>{ali*100:.2f}</pICMS><vICMS>{f2(vicms)}</vICMS></ICMS00></ICMS>'
            f'<PIS><PISAliq><CST>01</CST><vBC>{f2(bc_pc)}</vBC><pPIS>{PIS*100:.2f}</pPIS><vPIS>{f2(vpis)}</vPIS></PISAliq></PIS>'
            f'<COFINS><COFINSAliq><CST>01</CST><vBC>{f2(bc_pc)}</vBC><pCOFINS>{COFINS*100:.2f}</pCOFINS><vCOFINS>{f2(vcof)}</vCOFINS></COFINSAliq></COFINS>'
            f'</imposto></det>')
    vnf = tot["vprod"] - tot["vdesc"]
    return (f'<?xml version="1.0" encoding="UTF-8"?><nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00"><NFe>'
            f'<infNFe Id="NFe{ch}" versao="4.00"><ide><cUF>{CUF[emit["uf"]]}</cUF><natOp>{natop}</natOp><mod>55</mod>'
            f'<serie>1</serie><nNF>{nnf}</nNF><dhEmi>{dt:%Y-%m-%dT%H:%M:%S}-03:00</dhEmi><tpNF>{tpnf}</tpNF>'
            f'<finNFe>{finnfe}</finNFe></ide>'
            f'<emit><CNPJ>{emit["cnpj"]}</CNPJ><xNome>{escape(emit["nome"])}</xNome><enderEmit><UF>{emit["uf"]}</UF></enderEmit><CRT>3</CRT></emit>'
            f'<dest><CNPJ>{dest["cnpj"]}</CNPJ><xNome>{escape(dest["nome"])}</xNome><enderDest><UF>{dest["uf"]}</UF></enderDest></dest>'
            + "".join(dets) +
            f'<total><ICMSTot><vBC>{f2(tot["vbc"])}</vBC><vICMS>{f2(tot["vicms"])}</vICMS><vProd>{f2(tot["vprod"])}</vProd>'
            f'<vDesc>{f2(tot["vdesc"])}</vDesc><vPIS>{f2(tot["vpis"])}</vPIS><vCOFINS>{f2(tot["vcofins"])}</vCOFINS>'
            f'<vNF>{f2(vnf)}</vNF></ICMSTot></total></infNFe></NFe>'
            f'<protNFe versao="4.00"><infProt><chNFe>{ch}</chNFe><cStat>{cstat}</cStat></infProt></protNFe></nfeProc>')

def salvar(ch, xml):
    with open(os.path.join(OUT, f"{ch}-nfe.xml"), "w", encoding="utf-8") as f: f.write(xml)

nnf_emp, nnf_forn = 1000, {f["cnpj"]: 5000 for f in FORNECEDORES}
vendas_emitidas, arquivos = [], 0
for mes in range(1, 13):
    base = datetime(2025, mes, 1)
    # compras: 2 pedidos por fornecedor por mês
    for fi, forn in enumerate(FORNECEDORES):
        for pedido in range(2):
            dt = base + timedelta(days=3 + pedido * 14, hours=9)
            itens = []
            for p in PRODUTOS:
                if p[6] != fi: continue
                qtd = round(p[7] * SAZONAL[mes-1] * random.uniform(0.45, 0.6))
                custo = p[4] * (1 + 0.004 * (mes - 1)) * random.uniform(0.98, 1.02)   # leve inflação no ano
                cfop = "5102" if forn["uf"] == "SP" else "6102"
                itens.append({"prod": p, "qtd": qtd, "vun": round(custo, 4), "cfop": cfop,
                              "icms": aliq_icms(forn["uf"], "SP"), "pc": True})
            nnf_forn[forn["cnpj"]] += 1
            ch = chave(CUF[forn["uf"]], dt, forn["cnpj"], 1, nnf_forn[forn["cnpj"]])
            salvar(ch, nota_xml(ch, dt, nnf_forn[forn["cnpj"]], forn, EMPRESA, 1, 1, "Venda de mercadoria", itens)); arquivos += 1
    # vendas: ~35 notas por mês
    for _ in range(int(35 * SAZONAL[mes-1])):
        cli = random.choice(CLIENTES); dt = base + timedelta(days=random.randint(0, 27), hours=random.randint(8, 17))
        itens = []
        for p in random.sample(PRODUTOS, random.randint(2, 5)):
            qtd = max(1, round(p[7] * SAZONAL[mes-1] / 35 * random.uniform(1.2, 3.0)))
            cfop = "5102" if cli["uf"] == "SP" else "6102"
            itens.append({"prod": p, "qtd": qtd, "vun": p[5] * (1 + 0.003 * (mes - 1)), "cfop": cfop,
                          "icms": aliq_icms("SP", cli["uf"]), "pc": True, "desc": random.choice([0, 0, 0, 0.03, 0.05])})
        nnf_emp += 1
        ch = chave("35", dt, EMPRESA["cnpj"], 1, nnf_emp)
        salvar(ch, nota_xml(ch, dt, nnf_emp, EMPRESA, cli, 1, 1, "Venda de mercadoria", itens)); arquivos += 1
        vendas_emitidas.append((ch, dt, cli, itens))
    # devoluções de venda: 2 por mês, parciais
    for ch_orig, dt_orig, cli, itens in random.sample([v for v in vendas_emitidas if v[1].month == mes], 2):
        it = dict(random.choice(itens)); it["qtd"] = max(1, round(it["qtd"] * 0.3))
        it["cfop"] = "1202" if cli["uf"] == "SP" else "2202"
        dt = dt_orig + timedelta(days=2); nnf_emp += 1
        ch = chave("35", dt, EMPRESA["cnpj"], 1, nnf_emp)
        salvar(ch, nota_xml(ch, dt, nnf_emp, EMPRESA, cli, 0, 4, "Devolucao de venda", [it])); arquivos += 1

# --- problemas plantados para o analisador encontrar ---
dup_ch = vendas_emitidas[10][0]
with open(os.path.join(OUT, f"{dup_ch}-nfe.xml"), encoding="utf-8") as f: conteudo = f.read()
with open(os.path.join(OUT, f"{dup_ch}-nfe (copia).xml"), "w", encoding="utf-8") as f: f.write(conteudo)
ch_d, dt_d, cli_d, it_d = vendas_emitidas[25]; nnf_emp += 1
ch = chave("35", dt_d, EMPRESA["cnpj"], 1, nnf_emp)
salvar(ch, nota_xml(ch, dt_d, nnf_emp, EMPRESA, cli_d, 1, 1, "Venda de mercadoria", it_d, cstat="110"))  # denegada
ch_bad = vendas_emitidas[40][0]
bad = ch_bad[:-1] + str((int(ch_bad[-1]) + 1) % 10)
with open(os.path.join(OUT, f"{ch_bad}-nfe.xml"), encoding="utf-8") as f: conteudo = f.read()
os.remove(os.path.join(OUT, f"{ch_bad}-nfe.xml"))
with open(os.path.join(OUT, f"{bad}-nfe.xml"), "w", encoding="utf-8") as f: f.write(conteudo.replace(ch_bad, bad))
print(f"{arquivos + 2} XMLs fictícios gerados em {OUT}")
