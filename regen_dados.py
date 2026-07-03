"""
regen_dados.py — versão LIGHT do gerar_dados.py que roda no GitHub Actions.

Reprocessa o dados.json usando SOMENTE arquivos JSON que já estão no repo:
- dados.json anterior (usa como base — mantém produtos PSI, catálogo, saldos)
- pedidos_kanban.json (sell_in + prox_chegada)
- produto_overrides.json (forecast/PV/PC/marca/tipo por produto)
- produto_status.json (não usa aqui — cliente aplica em runtime)

NÃO lê Excel (PSI_integrado, estoque consolidado, VENDAS B2B, catálogo Sênior).
Pra atualizar essas fontes, Douglas roda `gerar_dados.py` local (versão completa).

O que este script faz:
1. Recalcula sell_in semanal a partir do kanban (match por código > nome)
2. Atualiza prox_chegada, in_transit, entradas
3. Aplica overrides (marca, tipo, PV, PC, fcst_mensal, pv_mensal, pc_mensal)
4. Propaga fcst_mensal (override) pra forecast semanal
5. Cascateia saldo (backward + forward) usando realizado quando > 0, forecast senão
6. Promove produtos do catálogo com override.fcst_mensal → produtos (Curva A virtual)

Uso: python regen_dados.py [--dry-run]
"""
import json
import os
import sys
import calendar
import datetime
from copy import deepcopy

ROOT = os.path.dirname(os.path.abspath(__file__))
DADOS_PATH     = os.path.join(ROOT, "dados.json")
PEDIDOS_PATH   = os.path.join(ROOT, "pedidos_kanban.json")
OVERRIDES_PATH = os.path.join(ROOT, "produto_overrides.json")

BASE_DATE = datetime.date(2025, 12, 29)
DRY_RUN = "--dry-run" in sys.argv

# Mapa semana → mês (idêntico ao V3 e gerar_dados.py)
WEEK_TO_MONTH = {
    1: "JAN/26", 5: "FEV/26", 9: "MAR/26", 14: "ABR/26", 18: "MAI/26",
    23: "JUN/26", 27: "JUL/26", 31: "AGO/26", 36: "SET/26", 40: "OUT/26",
    44: "NOV/26", 49: "DEZ/26", 53: "JAN/27", 58: "FEV/27", 62: "MAR/27",
    66: "ABR/27", 70: "MAI/27",
}

# Semana atual e da foto do Sênior (mesma coisa hoje — foto é feita quinzenalmente)
def _semana_hoje():
    return max(1, (datetime.date.today() - BASE_DATE).days // 7 + 1)

SEMANA_ATUAL   = _semana_hoje()
SEMANA_ESTOQUE = SEMANA_ATUAL  # backend não altera foto real (mantida no dados.json)

# Realizado manual (semanas fechadas sem base B2B). Douglas atualiza toda semana.
# Fonte da verdade fica no gerar_dados.py (não replicar aqui). Este script só
# reusa o realizado que JÁ está em dados.json.produtos[].psi.realizado.


def _fase_para_status(fase):
    fase = (fase or "").lower()
    if fase in ("transito", "porto", "numerario"): return "EM_TRANSITO"
    if fase in ("recebido", "estoque"):            return "RECEBIDO"
    return "PEDIDO"


def _month_week_range(mes):
    """Retorna (start, end) das semanas do mês. Ex: JUN/26 → (23, 26)."""
    entries = sorted(WEEK_TO_MONTH.items())
    for i, (w, m) in enumerate(entries):
        if m != mes: continue
        start = w
        end = entries[i+1][0] - 1 if i + 1 < len(entries) else start + 3
        return start, end, end - start + 1
    return None


_MM_ABR = {'JAN':1,'FEV':2,'MAR':3,'ABR':4,'MAI':5,'JUN':6,
           'JUL':7,'AGO':8,'SET':9,'OUT':10,'NOV':11,'DEZ':12}


def _mes_start_end(mes: str):
    """('JUL/26') → (date(2026,7,1), date(2026,7,31), 31)"""
    m_abr, y_suf = mes.split('/')
    year, month = 2000 + int(y_suf), _MM_ABR[m_abr]
    dias = calendar.monthrange(year, month)[1]
    return datetime.date(year, month, 1), datetime.date(year, month, dias), dias


def _dias_semana_no_mes(semana: int, ini_mes: datetime.date, fim_mes: datetime.date) -> int:
    """Quantos dias da semana W caem no intervalo do mês (0..7)."""
    ini_sem = BASE_DATE + datetime.timedelta(days=(semana - 1) * 7)
    fim_sem = ini_sem + datetime.timedelta(days=6)
    inicio = max(ini_sem, ini_mes)
    fim    = min(fim_sem, fim_mes)
    return max(0, (fim - inicio).days + 1)


def propagar_fcst_mensal(psi, fcst_mensal):
    """Distribui fcst_mensal por semana, PROPORCIONAL aos dias que caem em cada mês.

    Regra (Opção C, 03/07/2026):
    - Cada dia do mês vale `total_mensal / dias_do_mes`.
    - Cada semana ganha a SOMA dos valores dos 7 dias dela (que podem estar em
      meses diferentes, no caso de semanas-fronteira).
    - Exemplo: W27 (29/jun-5/jul) com JUN/26=950 (30 dias) e JUL/26=1400 (31 dias):
        forecast[W27] = 2 × (950/30) + 5 × (1400/31) = 63.3 + 225.8 = 289 un.
    - Semanas 100% passadas (fim antes da SEMANA_ATUAL) NÃO são tocadas.
    """
    if not psi or not isinstance(psi.get("forecast"), list) or not isinstance(psi.get("semanas"), list):
        return
    semanas  = psi["semanas"]
    forecast = psi["forecast"]

    # Pré-calcula ranges dos meses do fcst_mensal (uma vez só)
    ranges = {}
    for mes in fcst_mensal:
        try:
            ini, fim, dias = _mes_start_end(mes)
            total = int(fcst_mensal[mes]) if isinstance(fcst_mensal[mes], (int, float)) else 0
            ranges[mes] = (ini, fim, dias, total)
        except (KeyError, ValueError):
            continue

    # Descobre quais semanas serão tocadas (têm dias em algum dos meses do fcst_mensal).
    # Inclui semanas passadas — o cascade usa `realizado` quando existe (não afeta saldo),
    # mas o forecast semanal precisa refletir o mensal pra o card do mês bater.
    semanas_tocadas = set()
    for mes, (ini, fim, dias, total) in ranges.items():
        for i, w in enumerate(semanas):
            if _dias_semana_no_mes(w, ini, fim) > 0:
                semanas_tocadas.add(i)

    # Zera as tocadas antes de somar as contribuições
    for i in semanas_tocadas:
        forecast[i] = 0.0

    # Soma as contribuições proporcionais de cada mês
    for mes, (ini, fim, dias, total) in ranges.items():
        if dias <= 0 or total < 0: continue
        val_por_dia = total / dias if dias else 0
        for i, w in enumerate(semanas):
            if i not in semanas_tocadas: continue
            d_sem = _dias_semana_no_mes(w, ini, fim)
            if d_sem > 0:
                forecast[i] += val_por_dia * d_sem

    # Arredonda pra inteiro nas semanas tocadas
    for i in semanas_tocadas:
        forecast[i] = int(round(forecast[i]))


def aplicar_kanban_sell_in(produtos, kanban_ativos):
    """Recalcula psi.sell_in de cada produto a partir do kanban.
    Match: código exato > nome upper-strip.
    Também atualiza entradas, prox_chegada, in_transit.
    """
    por_codigo = {}   # cod (int) → {semana: qtd}
    por_nome   = {}   # nome_upper → {semana: qtd}
    ent_codigo = {}   # cod → [entradas]
    ent_nome   = {}   # nome_upper → [entradas]

    for p in kanban_ativos:
        nome = (p.get("nome") or "").strip().upper()
        sem  = p.get("semana")
        qtd  = p.get("qtd") or 0
        cod  = p.get("codigo")
        if not nome or not isinstance(sem, (int, float)) or not qtd: continue
        sem = int(sem); qtd = int(qtd)
        entrada = {
            "semana": sem, "quantidade": qtd,
            "status": _fase_para_status(p.get("fase")),
            "mes": p.get("mes", ""),
            "codigo_pedido": p.get("codigo_pedido", ""),
            "fase": p.get("fase", ""),
        }
        por_nome.setdefault(nome, {})
        por_nome[nome][sem] = por_nome[nome].get(sem, 0) + qtd
        ent_nome.setdefault(nome, []).append(entrada)
        if cod is not None:
            try: cod_i = int(cod)
            except (ValueError, TypeError): continue
            por_codigo.setdefault(cod_i, {})
            por_codigo[cod_i][sem] = por_codigo[cod_i].get(sem, 0) + qtd
            ent_codigo.setdefault(cod_i, []).append(entrada)

    for p in produtos:
        psi = p.get("psi")
        if not psi: continue
        cod = p.get("codigo")
        nome_u = (p.get("nome") or "").strip().upper()
        # Match: código > nome
        sin_map = (por_codigo.get(cod) if cod is not None else None) or por_nome.get(nome_u, {})
        ent_list = (ent_codigo.get(cod) if cod is not None else None) or ent_nome.get(nome_u, [])
        # Aplica sell_in em cada semana do PSI
        semanas = psi.get("semanas", [])
        psi["sell_in"] = [int(sin_map.get(w, 0)) for w in semanas]
        # Entradas futuras (>= SEMANA_ATUAL)
        ent_futuras = sorted([e for e in ent_list if e.get("semana", 0) >= SEMANA_ATUAL],
                             key=lambda e: e["semana"])
        p["entradas"] = ent_futuras
        p["prox_chegada"] = ent_futuras[0] if ent_futuras else None
        p["in_transit"] = sum(e["quantidade"] for e in ent_futuras if e.get("status") == "EM_TRANSITO")


def aplicar_overrides(produtos, overrides):
    """Aplica overrides.marca/tipo/pv/pc/fcst_mensal/pv_mensal/pc_mensal sobre os produtos."""
    for p in produtos:
        ov = overrides.get(str(p.get("codigo") or ""))
        if not ov: continue
        if ov.get("marca") is not None: p["marca"] = ov["marca"]
        if ov.get("tipo")  is not None: p["tipo"]  = ov["tipo"]
        if ov.get("pv")    is not None: p["pv"]    = ov["pv"]
        if ov.get("pc")    is not None: p["pc"]    = ov["pc"]
        if ov.get("fcst_mensal"):
            # Inclui zeros no merge (v é 0 legítimo pra meses zerados de propósito).
            # Filtrar por `if v` estava mantendo o valor antigo em meses zerados.
            p["fcst_mensal"] = {**(p.get("fcst_mensal") or {}), **ov["fcst_mensal"]}
        # Sempre re-propaga TODOS os meses do fcst_mensal (não só os do override).
        # Se override sumiu — cliente deleta chaves iguais ao dados.json na migração
        # fantasma — as semanas continuavam com distribuição arredondada antiga.
        # Propagar em cima do fcst_mensal completo garante que a soma semanal
        # sempre bate exato com o mensal.
        if p.get("fcst_mensal"):
            propagar_fcst_mensal(p.get("psi"), p["fcst_mensal"])
        if ov.get("pv_mensal"):
            p["pv_mensal"] = {**(p.get("pv_mensal") or {}), **ov["pv_mensal"]}
        if ov.get("pc_mensal"):
            p["pc_mensal"] = {**(p.get("pc_mensal") or {}), **ov["pc_mensal"]}


def cascatear_saldo(produtos):
    """Recalcula psi.saldo APENAS forward — semanas passadas ficam como estavam.
    - Semana < SEMANA_ATUAL com realizado > 0: usa realizado
    - Senão: usa forecast
    - Cascade FORWARD apenas (backward foi removido em 02/07/2026 — causava
      padrão dente-de-serra nas semanas passadas).
    """
    for p in produtos:
        psi = p.get("psi")
        if not psi: continue
        semanas = psi.get("semanas", [])
        saldo   = psi.get("saldo", [])
        forecast = psi.get("forecast", [])
        realizado = psi.get("realizado", []) or [0]*len(semanas)
        sell_in = psi.get("sell_in", [])
        if not semanas or SEMANA_ESTOQUE not in semanas: continue

        idx = semanas.index(SEMANA_ESTOQUE)
        # Saldo inicial da SEMANA_ESTOQUE = saldo_atual do produto (foto Sênior)
        saldo_foto = int(p.get("saldo_atual") or 0)

        def _consumo(i):
            w = semanas[i]
            fc = forecast[i]
            rl = realizado[i]
            return rl if (w <= SEMANA_ATUAL and rl > 0) else fc

        # FORWARD apenas — semanas passadas mantêm o valor que estava em psi.saldo
        # (que foi calculado pelo gerar_dados.py com histórico congelado).
        saldo[idx] = max(0, round(saldo_foto - _consumo(idx) + sell_in[idx]))
        for i in range(idx + 1, len(semanas)):
            saldo[i] = max(0, round(saldo[i-1] + sell_in[i] - _consumo(i)))


def promover_do_catalogo(dados, overrides, kanban_ativos):
    """Adiciona entries virtuais em dados.produtos pra produtos do catálogo com override.fcst_mensal.
    Se produto já está em produtos, ignora (override normal já vai aplicar).
    """
    codigos_em_produtos = {p.get("codigo") for p in dados.get("produtos", []) if p.get("codigo") is not None}
    catalogo_por_codigo = {c.get("codigo"): c for c in dados.get("catalogo", []) if c.get("codigo") is not None}

    # Precisa das semanas padrão pra construir PSI sintético
    _semanas_padrao = []
    for p in dados.get("produtos", []):
        s = (p.get("psi") or {}).get("semanas")
        if s: _semanas_padrao = list(s); break
    if not _semanas_padrao:
        _semanas_padrao = list(range(1, 54))

    # Kanban indexado por código
    kanban_sin_cod = {}
    kanban_ent_cod = {}
    for p in kanban_ativos:
        cod = p.get("codigo")
        sem = p.get("semana"); qtd = p.get("qtd") or 0
        if not cod or not sem or not qtd: continue
        try: cod = int(cod)
        except (ValueError, TypeError): continue
        kanban_sin_cod.setdefault(cod, {})
        kanban_sin_cod[cod][int(sem)] = kanban_sin_cod[cod].get(int(sem), 0) + int(qtd)
        kanban_ent_cod.setdefault(cod, []).append({
            "semana": int(sem), "quantidade": int(qtd),
            "status": _fase_para_status(p.get("fase")),
            "mes": p.get("mes", ""),
            "codigo_pedido": p.get("codigo_pedido", ""),
            "fase": p.get("fase", ""),
        })

    promovidos = 0
    for cod_str, ov in overrides.items():
        fcst = ov.get("fcst_mensal") or {}
        if not any(v for v in fcst.values() if v): continue
        try: cod = int(cod_str)
        except (ValueError, TypeError): continue
        if cod in codigos_em_produtos: continue
        cat = catalogo_por_codigo.get(cod)
        if cat is None: continue

        # Sell_in via kanban (por código)
        sin_map = kanban_sin_cod.get(cod, {})
        ent_list = kanban_ent_cod.get(cod, [])
        ent_futuras = sorted([e for e in ent_list if e["semana"] >= SEMANA_ATUAL], key=lambda e: e["semana"])
        prox = ent_futuras[0] if ent_futuras else None
        in_transit = sum(e["quantidade"] for e in ent_futuras if e["status"] == "EM_TRANSITO")

        # PSI sintético
        forecast_arr = [0]*len(_semanas_padrao)
        sellin_arr   = [int(sin_map.get(w, 0)) for w in _semanas_padrao]
        realizado_arr = [0]*len(_semanas_padrao)
        saldo_arr    = [0]*len(_semanas_padrao)
        saldo_cat    = int(cat.get("saldo_atual", 0) or 0)

        # fc_sem estimado: total 3 primeiros meses com forecast / semanas
        MESES = ['MAI/26','JUN/26','JUL/26','AGO/26','SET/26','OUT/26',
                 'NOV/26','DEZ/26','JAN/27','FEV/27','MAR/27','ABR/27','MAI/27']
        meses_com_fc = [m for m in MESES if fcst.get(m, 0) > 0][:3]
        if meses_com_fc:
            r_ini = _month_week_range(meses_com_fc[0])
            r_fim = _month_week_range(meses_com_fc[-1])
            total_sem = (r_fim[1] - r_ini[0] + 1) if r_ini and r_fim else 12
            fc_sem = round(sum(fcst.get(m, 0) for m in meses_com_fc) / max(1, total_sem))
        else:
            fc_sem = 0

        dados["produtos"].append({
            "nome":            cat.get("nome", ""),
            "marca":           ov.get("marca") or cat.get("marca", ""),
            "codigo":          cod,
            "familia":         cat.get("familia", "") or "",
            "tipo":            ov.get("tipo") or "Nacional",
            "pv":              float(ov.get("pv", 0) or 0),
            "pc":              float(ov.get("pc", 0) or 0),
            "saldo_atual":     saldo_cat,
            "fc_sem":          fc_sem,
            "in_transit":      in_transit,
            "prox_chegada":    prox,
            "entradas":        ent_futuras,
            "fcst_mensal":     fcst,
            "venda_mensal":    {},
            "realizado_split": {},
            "psi": {
                "semanas": _semanas_padrao,
                "forecast": forecast_arr,
                "sell_in": sellin_arr,
                "realizado": realizado_arr,
                "saldo": saldo_arr,
            },
            "_origem": "catalogo_promovido",
        })
        promovidos += 1

    if promovidos:
        print(f"Promovidos do catálogo: {promovidos}")
    # Remove promovidos do catálogo pra não duplicar
    codigos_promovidos = {p["codigo"] for p in dados["produtos"] if p.get("_origem") == "catalogo_promovido"}
    dados["catalogo"] = [c for c in dados.get("catalogo", []) if c.get("codigo") not in codigos_promovidos]


def main():
    print(f"regen_dados.py — SEMANA_ATUAL=W{SEMANA_ATUAL}")

    with open(DADOS_PATH, encoding="utf-8") as f: dados = json.load(f)
    with open(PEDIDOS_PATH, encoding="utf-8") as f: pedidos = json.load(f)
    with open(OVERRIDES_PATH, encoding="utf-8") as f: overrides = json.load(f)

    kanban_ativos = [p for p in pedidos if not p.get("cancelado")]
    print(f"  Produtos PSI: {len(dados.get('produtos', []))}  Catálogo: {len(dados.get('catalogo', []))}")
    print(f"  Pedidos kanban ativos: {len(kanban_ativos)}")
    print(f"  Overrides: {len(overrides)}")

    # Snapshot pra comparar
    dados_orig = deepcopy(dados)

    # 1. Promove catálogo → PSI (se override tem fcst_mensal)
    promover_do_catalogo(dados, overrides, kanban_ativos)

    # 2. Aplica kanban → sell_in / entradas / prox_chegada
    aplicar_kanban_sell_in(dados["produtos"], kanban_ativos)

    # 3. Aplica overrides (marca/tipo/PV/PC/fcst_mensal/etc)
    aplicar_overrides(dados["produtos"], overrides)

    # 4. Cascade saldo
    cascatear_saldo(dados["produtos"])

    # Metadata
    dados["semana_atual"]   = SEMANA_ATUAL
    dados["semana_estoque"] = SEMANA_ESTOQUE
    dados["gerado_em"]      = datetime.datetime.now().isoformat(timespec="seconds")
    dados["week_to_month"]  = {str(k): v for k, v in WEEK_TO_MONTH.items()}

    if DRY_RUN:
        print("[DRY-RUN] Não sobrescreveu dados.json")
        return

    # Grava só se houve mudança relevante (evita commit vazio no CI)
    novo = json.dumps(dados, ensure_ascii=False, indent=2, default=str)
    antigo = json.dumps(dados_orig, ensure_ascii=False, indent=2, default=str)
    if novo == antigo:
        print("Sem mudança relevante — dados.json inalterado")
        return

    with open(DADOS_PATH, "w", encoding="utf-8") as f: f.write(novo)
    print(f"dados.json atualizado ({len(dados['produtos'])} produtos)")


if __name__ == "__main__":
    main()
