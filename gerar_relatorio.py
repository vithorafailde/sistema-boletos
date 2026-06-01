"""
Sistema de Conferência de Boletos - Abril 2026
Combina dados do Excel + PDFs de Condomínio + Boletos de referência
"""
import re, os
import openpyxl
import pdfplumber

BASE = "C:/Users/vitho/OneDrive/Área de Trabalho/ABRIL 2026 - Copia"

# ─── 1. DADOS DE CONDOMÍNIO extraídos dos PDFs por visão ─────────────────────
# Formato: chave = nome do arquivo sem .pdf
# Valores: {cota, energia, tx_leitura, agua, gas, vaga_moto, tag, outros, fundo_reserva, total_boleto, tenant_key}
# OBS: fundo_reserva NÃO é repassado ao locatário
# OBS: correio NÃO é repassado ao locatário

COND_DATA = {
    # ── VM = Condomínio Edifício Clube Jardim, Rua Itaúna 1050 (BBZ/Superlógica)
    "vm_claudemir_92a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-92A",
        "proprietario_cond": "Claudemir Izildo Frutuoso",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 69.46, "gas": 1.05, "fundo_reserva": 31.31,
        "total_boleto_cond": 776.62,
        "tenant_key": "MARIAJOSEOLIVEIRA",
    },
    "vm_luciaria_163a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-163A",
        "proprietario_cond": "Osmar Oliveira Silva",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 166.55, "gas": 181.35, "fundo_reserva": 31.31,
        "total_boleto_cond": 1054.01,
        "tenant_key": "EDGARDGARCIA",
    },
    "vm_cleubio_73b": {
        "edificio": "Clube Jardim (VM)", "unidade": "02-73B",
        "proprietario_cond": "Luciara Maria de Oliveira",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 64.91, "gas": 47.34, "fundo_reserva": 31.31,
        "total_boleto_cond": 818.36,
        "tenant_key": "ROBERTODASILVAPEDROSO",
    },
    "vm_helio_56a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-56A",
        "proprietario_cond": "Helio Teixeira Romeiro",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 54.18, "gas": 50.66, "fundo_reserva": 31.31,
        "total_boleto_cond": 810.95,
        "tenant_key": "RODRIGOELIZIARIOLIMADAC",
    },
    "vm_marceandro_7c": {
        "edificio": "Clube Jardim (VM)", "unidade": "03-07C",
        "proprietario_cond": "Marceandro da Costa Cruz",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 116.31, "gas": 164.26, "fundo_reserva": 31.31,
        "total_boleto_cond": 986.68,
        "tenant_key": "LURDESSELIMADUTRA",
    },
    "vm_maria_53c": {
        "edificio": "Clube Jardim (VM)", "unidade": "03-53C",
        "proprietario_cond": "Maria Dias Santiago",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 0.00, "gas": 59.46, "fundo_reserva": 31.31,
        "total_boleto_cond": 765.57,
        "tenant_key": "ARLENEMARIAPASSOS",
    },
    "vm_mario_63b": {
        "edificio": "Clube Jardim (VM)", "unidade": "02-63B",
        "proprietario_cond": "Mário André Caveiro",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 101.20, "gas": 122.50, "fundo_reserva": 31.31,
        "total_boleto_cond": 929.81,
        "tenant_key": "JACICARLASOUZADASILVA",
    },
    "vm_mary_63a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-63A",
        "proprietario_cond": "Mary Michel Bacha",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 184.71, "gas": 0.00, "fundo_reserva": 31.31,
        "total_boleto_cond": 890.82,
        "tenant_key": "CLAUDIOBERNARDINO",
    },
    "vm_rita_91a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-91A",
        "proprietario_cond": "Dagoberto Pires Santana",
        "cota": 472.32, "energia": 30.75, "tx_leitura": 7.75,
        "agua": 0.78, "gas": 12.21, "fundo_reserva": 23.62,
        "total_boleto_cond": 547.43,
        "tenant_key": None,  # sem locatário identificado neste lote
    },
    "vm_rosana_145c": {
        "edificio": "Clube Jardim (VM)", "unidade": "03-145C",
        "proprietario_cond": "Rosana Alves",
        "cota": 473.40, "energia": 30.82, "tx_leitura": 7.75,
        "agua": 89.75, "gas": 89.63, "fundo_reserva": 23.67,
        "total_boleto_cond": 715.02,
        "tenant_key": "THIAGORIBEIROARTACHO",
    },
    "vm_sheila_34a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-34A",
        "proprietario_cond": "Sheila Gomes dos Santos Silva",
        "cota": 473.40, "energia": 30.82, "tx_leitura": 7.75,
        "agua": 190.19, "gas": 170.97, "fundo_reserva": 23.67,
        "total_boleto_cond": 896.80,
        "tenant_key": "PRISCILALOPESGARUTTI",
    },
    "vm_silvia_louro_66b": {
        "edificio": "Clube Jardim (VM)", "unidade": "02-66B",
        "proprietario_cond": "Sílvia Regina Louro",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 192.59, "gas": 237.41, "fundo_reserva": 31.31,
        "total_boleto_cond": 1136.11,
        "tenant_key": "MARCELOMARCOSLOPES",
    },
    "vm_tania_43a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-43A",
        "proprietario_cond": "Hélio Teixeira Romeiro (A/C Ericson)",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 240.62, "gas": 313.61, "fundo_reserva": 31.31,
        "total_boleto_cond": 1260.34,
        "tenant_key": "ERICSONDJALMADASILVACOS",
    },
    "vm_vanessa_35b": {
        "edificio": "Clube Jardim (VM)", "unidade": "02-35B",
        "proprietario_cond": "Vanessa da Silva Paes",
        "cota": 473.40, "energia": 30.82, "tx_leitura": 7.75,
        "agua": 102.01, "gas": 154.76, "fundo_reserva": 23.67,
        "total_boleto_cond": 792.41,
        "tenant_key": "NATHALIANASCIMENTOOLIVE",
    },
    "vm_waldemar_32b": {
        "edificio": "Clube Jardim (VM)", "unidade": "02-32B",
        "proprietario_cond": "Waldemar Roberto Perillo",
        "cota": 626.28, "energia": 40.77, "tx_leitura": 7.75,
        "agua": 109.09, "gas": 86.14, "fundo_reserva": 31.31,
        "total_boleto_cond": 901.34,
        "tenant_key": "RENATAPINTOANDRADE",
    },
    "vm_ze_luiz_104a": {
        "edificio": "Clube Jardim (VM)", "unidade": "01-104A",
        "proprietario_cond": "José Luiz Ramos",
        "cota": 473.40, "energia": 30.82, "tx_leitura": 7.75,
        "agua": 79.84, "gas": 63.30, "fundo_reserva": 23.67,
        "total_boleto_cond": 678.78,
        "tenant_key": "GABRIELADEARAUJOSANTOS",
    },
    # ── CV = Condomínio Vila Jardim Casa Verde, Av. Mandaqui 189 (MANAGER)
    "cv_carolina_44b": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "B-044",
        "proprietario_cond": "Carolina Manfrin Silverio",
        "cota": 711.25, "energia": 73.21, "agua": 145.57, "gas": 102.30,
        "vaga_moto": 43.06, "fundo_reserva": 35.56,
        "total_boleto_cond": 1110.95,
        "tenant_key": "MILENASALVADORBONAFE",
    },
    "cv_chang_102a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-102",
        "proprietario_cond": "Chen Fengling",
        "cota": 700.98, "energia": 72.15, "agua": 143.47, "gas": 100.82,
        "fundo_reserva": 35.05,
        "total_boleto_cond": 1052.47,
        "tenant_key": "WILSONNEPRIBEIRO",
    },
    "cv_elvira_115a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-115",
        "proprietario_cond": "Elvira Mitsuko Shimizu Frugis (via Sergio Frugis)",
        "cota": 530.18, "energia": 54.57, "agua": 108.52, "gas": 76.26,
        "fundo_reserva": 26.51, "correio": 3.70,
        "total_boleto_cond": 799.74,
        "tenant_key": "DANILOARAUJODESOUZA",
    },
    "cv_hp_64b": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "B-064",
        "proprietario_cond": "HP Adm. de Imóveis Próprios Ltda",
        "cota": 711.25, "energia": 73.21, "agua": 145.57, "gas": 102.30,
        "fundo_reserva": 35.56,
        "total_boleto_cond": 1067.89,
        "tenant_key": "LIVIARODRIGUESALEXANDRI",
    },
    "cv_joao_113b": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "B-113",
        "proprietario_cond": "João Azevedo da Silva",
        "cota": 549.44, "energia": 56.55, "agua": 112.46, "gas": 79.03,
        "fundo_reserva": 27.47,
        "total_boleto_cond": 824.95,
        "tenant_key": "ANAPAULADESANTANASILVER",
    },
    "cv_juliana_41a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-041",
        "proprietario_cond": "Juliana Stefani de Carvalho",
        "cota": 530.18, "energia": 54.57, "agua": 108.52, "gas": 76.26,
        "fundo_reserva": 26.51, "correio": 3.70,
        "total_boleto_cond": 799.74,
        "tenant_key": "HENRIQUEMOREIRA",
    },
    "cv_leandro_94a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-094",
        "proprietario_cond": "Leandro Caetano Mariano",
        "cota": 530.18, "energia": 54.57, "agua": 108.52, "gas": 76.26,
        "fundo_reserva": 26.51,
        "total_boleto_cond": 796.04,
        "tenant_key": "MATHEUSTHEODORODOSSANTO",
    },
    "cv_pablo_25a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-025",
        "proprietario_cond": "Pablo Ezequiel López",
        "cota": 530.18, "energia": 54.57, "agua": 108.52, "gas": 76.26,
        "fundo_reserva": 26.51, "correio": 3.70,
        "total_boleto_cond": 799.74,
        "tenant_key": "RODNEIVENTURADEOLIVEIRA",
    },
    "cv_rafael_103b": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "B-103",
        "proprietario_cond": "Rafael de Souza Andrade Lopes",
        "cota": 549.44, "energia": 56.55, "agua": 112.46, "gas": 79.03,
        "fundo_reserva": 27.47, "correio": 3.70,
        "total_boleto_cond": 828.65,
        "tenant_key": "RICARDOANTONIOSRUSSOJR",
    },
    "cv_raphael_38a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-038",
        "proprietario_cond": "Ana Lúcia Levy",
        "cota": 530.18, "energia": 54.57, "agua": 108.52, "gas": 76.26,
        "fundo_reserva": 26.51,
        "total_boleto_cond": 796.04,
        "tenant_key": "ADRIANADOSSANTOSREISCOC",
    },
    "cv_silvia_teixeira_56a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-056",
        "proprietario_cond": "Espólio de Sílvia Regina Teixeira",
        "cota": 700.98, "energia": 72.15, "agua": 143.47, "gas": 100.82,
        "fundo_reserva": 35.05,
        "total_boleto_cond": 1052.47,
        "tenant_key": "FERNANDOBARBARARAMOS",
    },
    "cv_walter_93a": {
        "edificio": "Vila Jardim Casa Verde (CV)", "unidade": "A-093",
        "proprietario_cond": "Walter dos Santos Filho",
        "cota": 700.98, "energia": 72.15, "agua": 143.47, "gas": 100.82,
        "tag": 15.00, "fundo_reserva": 35.05,
        "total_boleto_cond": 1067.47,
        "tenant_key": "CARINASATIKOSASAKEMATSU",
    },
    # ── VG = Condomínio Ed. Parque Jardim, Rua Henrique Felipe da Costa 681 (NEOCOND)
    "vg_adriana_44": {
        "edificio": "Ed. Parque Jardim (VG)", "unidade": "2-AU44",
        "proprietario_cond": "Adriana Santana de Aguiar",
        "cota": 498.79, "energia": 34.37, "agua": 66.69, "gas": 88.58,
        "fundo_reserva": 24.94,
        "total_boleto_cond": 713.37,
        "tenant_key": "ALANDERSONMACHADOCESAR",
    },
    "vg_leticia_34": {
        "edificio": "Ed. Parque Jardim (VG)", "unidade": "1-AU34",
        "proprietario_cond": "Leticia Neto Leme",
        "cota": 498.79, "energia": 34.37, "agua": 66.69, "gas": 156.12,
        "fundo_reserva": 24.94, "correio": 3.70,
        "total_boleto_cond": 784.61,
        "tenant_key": "VICTORFERREIRACOTTA",
    },
    "vg_sergio_53": {
        "edificio": "Ed. Parque Jardim (VG)", "unidade": "2-BU53",
        "proprietario_cond": "Sergio Luiz Miqueleti",
        "cota": 498.79, "energia": 34.37, "agua": 66.69, "gas": 109.70,
        "fundo_reserva": 24.94, "correio": 3.70,
        "total_boleto_cond": 738.19,
        "tenant_key": None,  # contrato novo (Fernando Hideyoshi), sem boleto neste lote
    },
    # ── Altavilla = Residencial Altavilla (OMNI)
    "altavilla_joao_37": {
        "edificio": "Residencial Altavilla", "unidade": "Apto 37",
        "proprietario_cond": "João Azevedo da Silva",
        "nota": "Demonstrativo complexo (OMNI). Total do boleto de condomínio: R$ 1.163,84. "
                "Itens repassados ao locatário apurados via boleto de referência.",
        "total_boleto_cond": 1163.84,
        "tenant_key": "DENISWILLIAMCOSMOPINTOD",
    },
    # ── Lamelas = Cond. Edifício Crist Lamelas (Itaú 341-7)
    "lamelas_francisco": {
        "edificio": "Ed. Crist Lamelas", "unidade": "Apto 05",
        "proprietario_cond": "Francisco Carlos de Oliveira",
        "cota": 470.00, "agua": 76.55, "gas": 24.62,
        "total_boleto_cond": 571.17,  # aprox (imagem parcialmente legível)
        "tenant_key": "ROGERIOPEREIRADEALMEIDA",
    },
    # ── Paisagem = Paisagem Vila Maria Alta (Lello)
    "paisagem_aline_134": {
        "edificio": "Paisagem Vila Maria Alta", "unidade": "Apto 134",
        "proprietario_cond": "Aline Gomes de Moraes Reis",
        "cota": 883.95, "vaga_moto": 100.00, "gas": 20.27, "agua": 84.03,
        "tx_leitura": 7.44, "melhorias": 260.94, "fundo_reserva": 44.20,
        "total_boleto_cond": 1390.80,
        "tenant_key": None,  # locatário não identificado no cadastro
        "nota": "Endereço: Rua Tapirai 62. Não encontrado locatário correspondente na planilha.",
    },
    # ── Victoria = Condomínio Edifício Victoria, Av. Prestes Maia 321 (Santander)
    "victoria_mauricio": {
        "edificio": "Ed. Victoria", "unidade": "Loja 13",
        "proprietario_cond": "Maurício Sampaio",
        "total_boleto_cond": 329.39,
        "nota": "Demonstrativo global do condomínio. Valor total do boleto da loja: R$ 329,39.",
        "tenant_key": "DONGMINGCHEN",
    },
    # ── Way = Cond. Edifício Way Vila Guilherme (Umuarama/Itaú)
    "way_vitor_34": {
        "edificio": "Cond. Way Vila Guilherme", "unidade": "1-000034",
        "proprietario_cond": "Victor Hugo Venancio da Silva",
        "cota": 633.61, "agua": 115.13, "gas": 177.21, "fundo_reserva": 31.68,
        "total_boleto_cond": 957.63,
        "tenant_key": "JEFFERSONSOARESDELIMA",
    },
}

# Chave: campos repassáveis (excluindo fundo_reserva e correio)
REPASSE_FIELDS = ["cota", "energia", "tx_leitura", "agua", "gas", "vaga_moto", "tag", "melhorias", "outros"]


def parse_brl(s):
    """Converte string BR para float."""
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0


def normalizar_nome(nome):
    """Remove espaços/acentos para chave de busca."""
    import unicodedata
    n = unicodedata.normalize("NFKD", nome.upper())
    n = "".join(c for c in n if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", n)


def ler_excel():
    """Lê contratos_locacao_editado.xlsx e retorna dict por chave do locatário."""
    wb = openpyxl.load_workbook(f"{BASE}/contratos_locacao_editado.xlsx")
    ws = wb.active
    contratos = {}
    for row in ws.iter_rows(values_only=True):
        if not row[4]:  # sem locatário
            continue
        locatario = str(row[4]).strip()
        aluguel = row[5]
        if not isinstance(aluguel, (int, float)):
            continue
        chave = normalizar_nome(locatario)
        contratos[chave] = {
            "locatario": locatario,
            "proprietario": str(row[3]).strip() if row[3] else "",
            "endereco": str(row[2]).strip() if row[2] else "",
            "tipo": str(row[1]).strip() if row[1] else "",
            "aluguel": float(aluguel),
        }
    return contratos


def ler_boletos_conferir():
    """Lê todos os boletos de referência (texto) e extrai total + itens."""
    pasta = f"{BASE}/BOLETOS Conferir"
    boletos = {}
    for fn in sorted(os.listdir(pasta)):
        if not fn.endswith(".pdf"):
            continue
        # Usa o boleto ALTERADO quando disponível (substitui o original)
        path = os.path.join(pasta, fn)
        with pdfplumber.open(path) as pdf:
            text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        if not text.strip():
            continue

        # Total
        valores = re.findall(r"R.{0,2}\s*([\d\.]+,\d{2})", text)
        total = parse_brl(valores[0]) if valores else None

        # Linhas de descrição (entre "10,00%" e "(=) Valor")
        desc_match = re.search(r"10,00%.*?\n(.*?)\(=\) Valor", text, re.DOTALL)
        items_raw = desc_match.group(1).strip() if desc_match else ""

        # Pagador
        pag = re.search(r"Pagador:\s*(.+?)(?:\s+CNPJ|$)", text, re.MULTILINE)
        pagador = pag.group(1).strip() if pag else ""

        # Vencimento
        venc = re.search(r"Vencimento\s*(\d{2}/\d{2}/\d{4})", text)
        vencimento = venc.group(1) if venc else ""

        chave_raw = re.sub(r"\s+ALTERADO.*", "", fn.replace("Boleto_", "").replace(".pdf", "").strip())
        chave = normalizar_nome(chave_raw)

        # Só grava se ainda não tem ou se este é ALTERADO (tem prioridade)
        is_alterado = "ALTERADO" in fn.upper() or "alterado" in fn
        if chave not in boletos or is_alterado:
            boletos[chave] = {
                "arquivo": fn,
                "pagador": pagador,
                "total": total,
                "items_raw": items_raw.replace("\n", " | "),
                "vencimento": vencimento,
            }
    return boletos


def calcular_cond_repasse(cond):
    """Soma apenas os campos repassáveis do condomínio."""
    total = 0.0
    itens = {}
    for f in REPASSE_FIELDS:
        v = cond.get(f, 0.0)
        if v:
            itens[f] = v
            total += v
    return total, itens


def montar_dados():
    contratos = ler_excel()
    boletos_ref = ler_boletos_conferir()

    # Mapeia tenant_key → cond_key
    tenant_to_cond = {}
    for ck, cd in COND_DATA.items():
        tk = cd.get("tenant_key")
        if tk:
            tenant_to_cond[tk] = ck

    registros = []

    # ─ Percorre todos os contratos do Excel
    for chave, c in sorted(contratos.items(), key=lambda x: x[1]["locatario"]):
        cond_key = tenant_to_cond.get(chave)
        cond = COND_DATA.get(cond_key, {}) if cond_key else {}
        ref = boletos_ref.get(chave, {})

        cond_repasse, cond_itens = calcular_cond_repasse(cond)

        total_calculavel = c["aluguel"] + cond_repasse
        total_ref = ref.get("total")
        diff = round(total_ref - total_calculavel, 2) if total_ref else None

        registros.append({
            "chave": chave,
            "locatario": c["locatario"],
            "proprietario": c["proprietario"],
            "endereco": c["endereco"],
            "tipo": c["tipo"],
            "aluguel": c["aluguel"],
            "cond_edificio": cond.get("edificio", ""),
            "cond_unidade": cond.get("unidade", ""),
            "cond_proprietario": cond.get("proprietario_cond", ""),
            "cond_itens": cond_itens,
            "cond_repasse": round(cond_repasse, 2),
            "cond_nota": cond.get("nota", ""),
            "total_calculavel": round(total_calculavel, 2),
            "total_ref": total_ref,
            "diff": diff,
            "items_ref": ref.get("items_raw", ""),
            "vencimento": ref.get("vencimento", ""),
            "tem_cond": bool(cond),
            "tem_ref": bool(total_ref),
        })

    # ─ Condomínios sem locatário identificado
    extras = []
    for ck, cd in COND_DATA.items():
        if cd.get("tenant_key") is None:
            extras.append({
                "cond_key": ck,
                "edificio": cd.get("edificio", ""),
                "unidade": cd.get("unidade", ""),
                "proprietario_cond": cd.get("proprietario_cond", ""),
                "total_boleto_cond": cd.get("total_boleto_cond", 0),
                "nota": cd.get("nota", ""),
            })

    return registros, extras


# ─── 3. GERADOR HTML ─────────────────────────────────────────────────────────

def fmt(v):
    if v is None:
        return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def status_badge(diff, tem_cond, tem_ref):
    if not tem_ref:
        return '<span class="badge badge-gray">Sem ref.</span>'
    if not tem_cond and diff:
        return f'<span class="badge badge-orange">Diff: {fmt(diff)}</span>'
    if diff is None:
        return '<span class="badge badge-gray">—</span>'
    if abs(diff) < 0.10:
        return '<span class="badge badge-green">✓ OK</span>'
    if abs(diff) < 10:
        return f'<span class="badge badge-yellow">≈ Diff: {fmt(diff)}</span>'
    return f'<span class="badge badge-red">✗ Diff: {fmt(diff)}</span>'


def gerar_html(registros, extras):
    linhas = []
    for r in registros:
        badge = status_badge(r["diff"], r["tem_cond"], r["tem_ref"])

        # Itens de condomínio
        cond_html = ""
        if r["cond_itens"]:
            itens_str = []
            nomes = {
                "cota": "Cota/Taxa Cond.", "energia": "Energia Comum",
                "tx_leitura": "Taxa Leitura", "agua": "Consumo Água",
                "gas": "Consumo Gás", "vaga_moto": "Vaga Moto",
                "tag": "TAG", "melhorias": "Melhorias",
            }
            for k, v in r["cond_itens"].items():
                itens_str.append(f"<span class='item-tag'>{nomes.get(k,k)}: {fmt(v)}</span>")
            cond_html = "".join(itens_str)

        # Linha de referência
        ref_html = ""
        if r["items_ref"]:
            ref_html = f'<div class="ref-items">{r["items_ref"]}</div>'

        diff_class = ""
        if r["diff"] is not None:
            if abs(r["diff"]) < 0.10:
                diff_class = "diff-ok"
            elif abs(r["diff"]) < 10:
                diff_class = "diff-small"
            else:
                diff_class = "diff-large"

        linhas.append(f"""
        <tr class="tenant-row" data-edificio="{r['cond_edificio']}">
          <td>
            <div class="tenant-name">{r['locatario']}</div>
            <div class="tenant-sub">{r['proprietario']} · {r['tipo']}</div>
            <div class="tenant-sub">{r['endereco']}</div>
          </td>
          <td class="val-col">{fmt(r['aluguel'])}<div class="source-tag">Excel</div></td>
          <td>
            {f'<div class="cond-name">{r["cond_edificio"]} {r["cond_unidade"]}</div>' if r['cond_edificio'] else '<span class="no-data">Sem condomínio</span>'}
            {cond_html}
            {f'<div class="cond-total">Total repasse: {fmt(r["cond_repasse"])}</div>' if r["cond_repasse"] else ''}
            {f'<div class="nota">{r["cond_nota"]}</div>' if r["cond_nota"] else ''}
          </td>
          <td class="val-col">{fmt(r['total_calculavel'])}<div class="source-tag">Calc.</div></td>
          <td class="val-col">
            {fmt(r['total_ref'])}
            {f'<div class="source-tag">Ref. ({r["vencimento"]})</div>' if r["total_ref"] else ''}
            {ref_html}
          </td>
          <td class="{diff_class}">{badge}</td>
        </tr>""")

    # Condomínios sem locatário
    extras_html = ""
    if extras:
        extras_html = "<h2>⚠️ Condomínios sem locatário identificado</h2><table class='table'><tr><th>Condomínio</th><th>Unidade</th><th>Proprietário</th><th>Total Boleto</th><th>Observação</th></tr>"
        for e in extras:
            extras_html += f"<tr><td>{e['edificio']}</td><td>{e['unidade']}</td><td>{e['proprietario_cond']}</td><td>{fmt(e['total_boleto_cond'])}</td><td>{e['nota']}</td></tr>"
        extras_html += "</table>"

    total_aluguel = sum(r["aluguel"] for r in registros)
    total_ref_sum = sum(r["total_ref"] for r in registros if r["total_ref"])
    ok_count = sum(1 for r in registros if r["diff"] is not None and abs(r["diff"]) < 0.10)
    approx_count = sum(1 for r in registros if r["diff"] is not None and 0.10 <= abs(r["diff"]) < 10)
    diff_count = sum(1 for r in registros if r["diff"] is not None and abs(r["diff"]) >= 10)
    sem_ref = sum(1 for r in registros if not r["tem_ref"])

    edificios = sorted(set(r["cond_edificio"] for r in registros if r["cond_edificio"]))
    filtros = '<button class="filter-btn active" onclick="filtrar(\'todos\', this)">Todos</button>'
    filtros += '<button class="filter-btn" onclick="filtrar(\'sem_cond\', this)">Sem Cond.</button>'
    filtros += '<button class="filter-btn" onclick="filtrar(\'sem_ref\', this)">Sem Ref.</button>'
    filtros += '<button class="filter-btn" onclick="filtrar(\'diff\', this)">Com Diferença</button>'
    for ed in edificios:
        filtros += f'<button class="filter-btn" onclick="filtrar(\'{ed}\', this)">{ed}</button>'

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Conferência de Boletos - Abril 2026</title>
<style>
  :root {{
    --bg: #f4f6f9; --card: #fff; --border: #e2e8f0;
    --text: #1e293b; --muted: #64748b; --accent: #3b82f6;
    --green: #16a34a; --yellow: #ca8a04; --red: #dc2626; --orange: #ea580c;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }}
  .header {{ background: #1e3a5f; color: #fff; padding: 18px 28px; }}
  .header h1 {{ font-size: 20px; }}
  .header p {{ color: #94a3b8; margin-top: 4px; font-size: 12px; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
  .stats {{ display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }}
  .stat {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 12px 18px; min-width: 140px; }}
  .stat .num {{ font-size: 22px; font-weight: 700; }}
  .stat .lbl {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}
  .filters {{ display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px; }}
  .filter-btn {{ padding: 5px 14px; border: 1px solid var(--border); border-radius: 20px; cursor: pointer; background: var(--card); font-size: 12px; transition: all .15s; }}
  .filter-btn:hover {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
  .filter-btn.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
  .table-wrap {{ overflow-x: auto; }}
  table.table {{ width: 100%; border-collapse: collapse; background: var(--card); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.07); }}
  th {{ background: #1e3a5f; color: #fff; padding: 10px 12px; text-align: left; font-size: 12px; font-weight: 600; position: sticky; top: 0; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: #f8fafc; }}
  .tenant-name {{ font-weight: 600; font-size: 13px; }}
  .tenant-sub {{ color: var(--muted); font-size: 11px; margin-top: 2px; }}
  .val-col {{ text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 600; }}
  .source-tag {{ font-size: 10px; color: var(--muted); font-weight: normal; }}
  .cond-name {{ font-size: 11px; color: var(--accent); font-weight: 600; margin-bottom: 4px; }}
  .item-tag {{ display: inline-block; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 4px; padding: 1px 6px; font-size: 11px; margin: 1px 2px 1px 0; color: #1d4ed8; }}
  .cond-total {{ font-size: 11px; font-weight: 600; color: var(--muted); margin-top: 4px; }}
  .no-data {{ color: #94a3b8; font-size: 11px; font-style: italic; }}
  .nota {{ font-size: 10px; color: var(--orange); margin-top: 3px; font-style: italic; }}
  .ref-items {{ font-size: 10px; color: var(--muted); margin-top: 4px; max-width: 280px; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; white-space: nowrap; }}
  .badge-green {{ background: #dcfce7; color: var(--green); }}
  .badge-yellow {{ background: #fef9c3; color: var(--yellow); }}
  .badge-red {{ background: #fee2e2; color: var(--red); }}
  .badge-orange {{ background: #ffedd5; color: var(--orange); }}
  .badge-gray {{ background: #f1f5f9; color: var(--muted); }}
  .diff-ok td {{ }}
  .diff-small {{ color: var(--yellow); font-weight: 600; }}
  .diff-large {{ color: var(--red); font-weight: 600; }}
  .legend {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; font-size: 11px; color: var(--muted); }}
  .legend span {{ display: flex; align-items: center; gap: 4px; }}
  h2 {{ font-size: 16px; margin: 28px 0 12px; }}
  .info-box {{ background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; padding: 14px 18px; margin-bottom: 20px; font-size: 12px; }}
  .info-box h3 {{ font-size: 13px; margin-bottom: 6px; color: #92400e; }}
  .info-box ul {{ padding-left: 16px; line-height: 1.8; }}
</style>
</head>
<body>
<div class="header">
  <h1>📋 Conferência de Boletos — Abril 2026</h1>
  <p>Sistema gerado automaticamente · Dados: Excel de contratos + PDFs de condomínio + Boletos de referência</p>
</div>
<div class="container">

  <div class="info-box">
    <h3>⚙️ Como os valores são compostos</h3>
    <ul>
      <li><strong>Aluguel</strong> → planilha <em>contratos_locacao_editado.xlsx</em></li>
      <li><strong>Cota Cond. / Água / Gás / Energia / Vagas</strong> → PDFs da pasta <em>CONDOMINIOS</em> (excluindo Fundo de Reserva e Correio, que são de responsabilidade do proprietário)</li>
      <li><strong>IPTU, Seguros, Abonos, Controle de Acesso, Retenção IR</strong> → <em>não extraíveis automaticamente</em> dos PDFs disponíveis (aparecem como diferença entre "Calculável" e "Referência")</li>
      <li><strong>Total de referência</strong> → boletos da pasta <em>BOLETOS Conferir</em></li>
    </ul>
  </div>

  <div class="stats">
    <div class="stat"><div class="num">{len(registros)}</div><div class="lbl">Locatários</div></div>
    <div class="stat"><div class="num" style="color:var(--green)">{ok_count}</div><div class="lbl">✓ Bateu (diff &lt; R$0,10)</div></div>
    <div class="stat"><div class="num" style="color:var(--yellow)">{approx_count}</div><div class="lbl">≈ Dif. pequena (&lt;R$10)</div></div>
    <div class="stat"><div class="num" style="color:var(--red)">{diff_count}</div><div class="lbl">✗ Diferença relevante</div></div>
    <div class="stat"><div class="num" style="color:var(--muted)">{sem_ref}</div><div class="lbl">Sem boleto de ref.</div></div>
    <div class="stat"><div class="num">{fmt(total_aluguel)}</div><div class="lbl">Total aluguéis (Excel)</div></div>
    <div class="stat"><div class="num">{fmt(total_ref_sum)}</div><div class="lbl">Total cobrado (refs)</div></div>
  </div>

  <div class="filters">{filtros}</div>

  <div class="table-wrap">
  <table class="table" id="mainTable">
    <tr>
      <th style="min-width:220px">Locatário</th>
      <th style="min-width:110px">Aluguel</th>
      <th style="min-width:320px">Condomínio (itens repassados)</th>
      <th style="min-width:110px">Calculável</th>
      <th style="min-width:200px">Referência</th>
      <th style="min-width:120px">Status</th>
    </tr>
    {''.join(linhas)}
  </table>
  </div>

  {extras_html}

</div>

<script>
function filtrar(tipo, btn) {{
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const rows = document.querySelectorAll('.tenant-row');
  rows.forEach(r => {{
    const edificio = r.dataset.edificio || '';
    const badge = r.querySelector('.badge');
    const badgeText = badge ? badge.textContent : '';
    const semCond = r.querySelector('.no-data');
    const semRef = badgeText.includes('Sem ref');
    const temDiff = badgeText.includes('Diff') || badgeText.includes('✗');
    let show = true;
    if (tipo === 'sem_cond') show = !!semCond;
    else if (tipo === 'sem_ref') show = semRef;
    else if (tipo === 'diff') show = temDiff;
    else if (tipo !== 'todos') show = edificio === tipo;
    r.style.display = show ? '' : 'none';
  }});
}}
</script>
</body>
</html>"""


if __name__ == "__main__":
    print("Carregando dados...")
    registros, extras = montar_dados()
    html = gerar_html(registros, extras)
    out = f"{BASE}/sistema_boletos/relatorio_abril2026.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Relatorio gerado: {out}")
    print(f"   {len(registros)} locatarios processados")
    ok = sum(1 for r in registros if r["diff"] is not None and abs(r["diff"]) < 0.10)
    diff = sum(1 for r in registros if r["diff"] is not None and abs(r["diff"]) >= 10)
    print(f"   OK: {ok}  |  Diferenca relevante: {diff}")
    # Print details for large diffs
    for r in sorted(registros, key=lambda x: abs(x["diff"] or 0), reverse=True):
        if r["diff"] is not None and abs(r["diff"]) >= 1:
            print(f"   {r['locatario'][:40]:<40} calc={r['total_calculavel']:>10.2f}  ref={r['total_ref']:>10.2f}  diff={r['diff']:>+8.2f}")
