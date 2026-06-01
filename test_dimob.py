"""
Script de teste DIMOB — roda fora do Flask para testar a lógica completa.
Executa: python test_dimob.py
"""
import sys, json, traceback
from pathlib import Path
from datetime import date

# Adiciona o diretório do app ao path
BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

# Importa funções do app
from app import (
    ler_excel_dimob, calcular_meses_dimob, aplicar_reajustes_excel,
    ler_dimob_historico, salvar_dimob_historico, norm,
    DIMOB_HISTORICO_FILE, UPLOAD_DIR, safe_float
)

PLANILHA = UPLOAD_DIR / "contratos.xlsx"
OK  = "[OK]"
ERR = "[ERRO]"
WARN= "[AVISO]"

erros = []

def check(desc, cond, detalhe=""):
    if cond:
        print(f"  {OK} {desc}")
    else:
        print(f"  {ERR} {desc}" + (f": {detalhe}" if detalhe else ""))
        erros.append(desc)

# ─── 1. Leitura da planilha ────────────────────────────────────────────────────
print("\n=== 1. Leitura da planilha contratos.xlsx ===")
try:
    contratos = ler_excel_dimob(PLANILHA)
    check("Planilha lida sem exceção", True)
    check("Pelo menos 1 contrato lido", len(contratos) > 0, f"encontrou {len(contratos)}")
    print(f"     Total de contratos: {len(contratos)}")

    # Valida campos obrigatórios
    for i, c in enumerate(contratos):
        for campo in ['locatario', 'proprietario', 'endereco', 'aluguel', 'num_linha']:
            check(f"Contrato {i+1} tem campo '{campo}'", campo in c and c[campo] is not None,
                  f"valor={c.get(campo)}")

    # Mostra primeiros 3
    print("\n  Primeiros contratos:")
    for c in contratos[:3]:
        print(f"    [{c['num_linha']}] {c['locatario']} | aluguel={c['aluguel']} "
              f"| mes_rej={c['mes_reajuste']} | mes_aplic={c['mes_aplicacao']}")

except Exception as e:
    print(f"  {ERR} Erro ao ler planilha: {e}")
    traceback.print_exc()
    erros.append("Leitura planilha")
    sys.exit(1)

# ─── 2. Cálculo de meses sem histórico ────────────────────────────────────────
print("\n=== 2. Cálculo de meses sem histórico ===")
try:
    hist_vazio = {}
    c0 = contratos[0]
    meses, tem_hist = calcular_meses_dimob(c0, 2025, hist_vazio)
    check("calcular_meses_dimob retorna lista de 12", len(meses) == 12, f"len={len(meses)}")
    check("sem histórico → tem_historico=False", not tem_hist)
    ativos = [v for v in meses if v is not None]
    check("Pelo menos 1 mês ativo", len(ativos) > 0)
    check("Todos os meses ativos usam aluguel atual", all(v == c0['aluguel'] for v in ativos),
          f"valores únicos: {set(ativos)}")
    print(f"     Meses: {meses}")
except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("calcular_meses sem histórico")

# ─── 3. Aplicação de reajuste fictício ────────────────────────────────────────
print("\n=== 3. Aplicar reajuste fictício no 1º contrato ===")
try:
    c0 = contratos[0]
    aluguel_original = c0['aluguel']
    aluguel_novo     = round(aluguel_original * 1.05, 2)  # +5% fictício
    mes_rej          = c0['mes_reajuste'] or 4
    mes_aplic        = (mes_rej % 12 + 1)

    print(f"     Contrato: {c0['locatario']}")
    print(f"     Aluguel original: {aluguel_original} → novo: {aluguel_novo}")
    print(f"     Mês reajuste: {mes_rej}, mês aplicação: {mes_aplic}")

    contrato_fict = [{
        'num_linha':    c0['num_linha'],
        'novo_aluguel': aluguel_novo,
        'locatario':    c0['locatario'],
        'mes_reajuste': mes_rej,
    }]

    n, erros_aplic = aplicar_reajustes_excel(PLANILHA, contrato_fict)
    check("aplicar_reajustes_excel executou sem exceção", True)
    check("1 contrato atualizado", n == 1, f"n={n}")
    check("Sem erros de aplicação", len(erros_aplic) == 0, str(erros_aplic))

    # Verifica se escreveu na planilha
    contratos_depois = ler_excel_dimob(PLANILHA)
    c0_depois = next((c for c in contratos_depois if c['num_linha'] == c0['num_linha']), None)
    check("Aluguel foi atualizado na planilha", c0_depois and abs(c0_depois['aluguel'] - aluguel_novo) < 0.01,
          f"valor na planilha: {c0_depois['aluguel'] if c0_depois else 'não encontrado'}")

except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("aplicar_reajuste")

# ─── 4. Verificar dimob_historico.json ────────────────────────────────────────
print("\n=== 4. Verificar dimob_historico.json ===")
try:
    check("dimob_historico.json foi criado", DIMOB_HISTORICO_FILE.exists())
    hist = ler_dimob_historico()
    ano  = str(date.today().year)
    check(f"Ano {ano} presente no histórico", ano in hist, f"chaves: {list(hist.keys())}")

    chave = norm(contratos[0]['locatario'])
    check("Contrato reajustado está no histórico", chave in hist.get(ano, {}),
          f"chave={chave}, disponíveis={list(hist.get(ano,{}).keys())[:5]}")

    reg = hist.get(ano, {}).get(chave, {})
    check("aluguel_antigo salvo corretamente",
          reg.get('aluguel_antigo') is not None and reg['aluguel_antigo'] > 0,
          f"valor={reg.get('aluguel_antigo')}")
    check("aluguel_novo salvo corretamente",
          reg.get('aluguel_novo') is not None and reg['aluguel_novo'] > 0,
          f"valor={reg.get('aluguel_novo')}")
    check("mes_aplicacao salvo corretamente",
          reg.get('mes_aplicacao') is not None,
          f"valor={reg.get('mes_aplicacao')}")
    print(f"     Registro: {reg}")

except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("dimob_historico")

# ─── 5. Cálculo de meses COM histórico ────────────────────────────────────────
print("\n=== 5. Cálculo de meses COM histórico ===")
try:
    hist = ler_dimob_historico()
    contratos_atuais = ler_excel_dimob(PLANILHA)
    c0 = contratos_atuais[0]
    meses, tem_hist = calcular_meses_dimob(c0, date.today().year, hist)

    check("tem_historico=True após salvar", tem_hist)
    check("Retorna 12 meses", len(meses) == 12)

    mes_aplic = c0['mes_aplicacao']
    if mes_aplic:
        antes  = [meses[i] for i in range(mes_aplic - 1) if meses[i] is not None]
        depois = [meses[i] for i in range(mes_aplic - 1, 12) if meses[i] is not None]
        alug_antigo = hist[str(date.today().year)][norm(c0['locatario'])]['aluguel_antigo']
        alug_novo   = c0['aluguel']
        check("Meses antes do reajuste usam aluguel antigo",
              all(abs(v - alug_antigo) < 0.01 for v in antes) if antes else True,
              f"antes={antes}, esperado={alug_antigo}")
        check("Meses depois do reajuste usam aluguel novo",
              all(abs(v - alug_novo) < 0.01 for v in depois) if depois else True,
              f"depois={depois}, esperado={alug_novo}")
        print(f"     mes_aplicacao={mes_aplic}")
        print(f"     Antes ({len(antes)} meses): {alug_antigo} | Depois ({len(depois)} meses): {alug_novo}")
        print(f"     Meses: {meses}")

    total = round(sum(v for v in meses if v is not None), 2)
    check("Total > 0", total > 0, f"total={total}")

except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("calcular_meses com histórico")

# ─── 6. Reverter planilha ao valor original ────────────────────────────────────
print("\n=== 6. Revertendo planilha ao valor original ===")
try:
    c0_orig = contratos[0]
    reverter = [{
        'num_linha':    c0_orig['num_linha'],
        'novo_aluguel': c0_orig['aluguel'],  # aluguel original
        'locatario':    c0_orig['locatario'],
        'mes_reajuste': c0_orig['mes_reajuste'] or 4,
    }]
    n, _ = aplicar_reajustes_excel(PLANILHA, reverter)
    # Limpa a entrada fictícia do histórico para não poluir dados reais
    hist = ler_dimob_historico()
    chave = norm(c0_orig['locatario'])
    ano   = str(date.today().year)
    if chave in hist.get(ano, {}):
        del hist[ano][chave]
        salvar_dimob_historico(hist)
    check("Planilha revertida ao valor original", n == 1)
    print("     (entrada fictícia removida do histórico)")
except Exception as e:
    print(f"  {ERR} Erro ao reverter: {e}")
    erros.append("reversão")

# ─── 7. Teste de contratos com data_inicio/data_fim ───────────────────────────
print("\n=== 7. Contratos com data_inicio / data_fim ===")
try:
    from datetime import date as dt
    c_test = {
        'locatario': 'TESTE', 'aluguel': 2000.0, 'mes_aplicacao': 7,
        'data_inicio': dt(2025, 3, 1), 'data_fim': dt(2025, 10, 31),
    }
    hist_test = {'2025': {norm('TESTE'): {'aluguel_antigo': 1800.0, 'aluguel_novo': 2000.0, 'mes_aplicacao': 7}}}
    meses, _ = calcular_meses_dimob(c_test, 2025, hist_test)
    check("Meses antes do início são None", meses[0] is None and meses[1] is None,
          f"jan={meses[0]}, fev={meses[1]}")
    check("Meses depois do fim são None", meses[10] is None and meses[11] is None,
          f"nov={meses[10]}, dez={meses[11]}")
    check("Meses dentro do período não são None", meses[2] is not None and meses[9] is not None,
          f"mar={meses[2]}, out={meses[9]}")
    check("Meses antes do reajuste (mar-jun) usam aluguel antigo",
          all(meses[i] == 1800.0 for i in range(2, 6)),
          f"mar-jun: {meses[2:6]}")
    check("Meses do reajuste em diante (jul-out) usam aluguel novo",
          all(meses[i] == 2000.0 for i in range(6, 10)),
          f"jul-out: {meses[6:10]}")
except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("data_inicio/data_fim")

# ─── 8. Teste edge cases ──────────────────────────────────────────────────────
print("\n=== 8. Edge cases ===")
try:
    # Mês 12 → mes_aplicacao deve ser 1 (janeiro do ano seguinte)
    c_dez = {'locatario': 'X', 'aluguel': 1000.0, 'mes_aplicacao': 1,
              'data_inicio': None, 'data_fim': None}
    hist_dez = {'2025': {norm('X'): {'aluguel_antigo': 900.0, 'aluguel_novo': 1000.0, 'mes_aplicacao': 1}}}
    meses_dez, _ = calcular_meses_dimob(c_dez, 2025, hist_dez)
    check("mes_aplicacao=1 (aniv dez) → todos os meses usam novo aluguel",
          all(v == 1000.0 for v in meses_dez if v is not None),
          f"meses={meses_dez}")

    # Contrato sem data_reajuste (mes_aplicacao=None) → usa aluguel atual p/ tudo
    c_sem = {'locatario': 'Y', 'aluguel': 1500.0, 'mes_aplicacao': None,
             'data_inicio': None, 'data_fim': None}
    meses_sem, tem = calcular_meses_dimob(c_sem, 2025, {})
    check("Sem mes_aplicacao → 12 meses com aluguel atual",
          all(v == 1500.0 for v in meses_sem), f"meses={meses_sem}")
    check("Sem mes_aplicacao → tem_historico=False", not tem)

except Exception as e:
    print(f"  {ERR} Erro: {e}")
    traceback.print_exc()
    erros.append("edge cases")

# ─── Resultado final ───────────────────────────────────────────────────────────
print("\n" + "="*50)
if erros:
    print(f"\033[91m  FALHOU — {len(erros)} erro(s):\033[0m")
    for e in erros:
        print(f"    - {e}")
else:
    print("\033[92m  TODOS OS TESTES PASSARAM\033[0m")
print("="*50 + "\n")
