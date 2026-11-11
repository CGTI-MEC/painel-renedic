#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PAINEL RENEDIC — SCRIPT DE PROCESSAMENTO
Uso: py -3.11 processar.py   (ou dê duplo clique no atualizar.bat)

O que faz:
  1. Pega o relatório MAIS RECENTE em brutos/ (Enec_Status_Adesao_Detalhado_*)
     — aceita o "falso Excel" (HTML) e o Excel de verdade.
  2. Descarta a coluna Responsável (dados pessoais nunca saem do computador).
  3. Remove Brasília/DF e Fernando de Noronha/PE das redes municipais.
  4. Cruza com a base do Censo (referencia/) para os filtros da equipe.
  5. Atualiza o histórico de envios ao MEC (curva de evolução).
  6. Gera saida/redes.json, saida/envios.json, saida/config.json e saida/meta.json.
"""

import json, re, sys, hashlib, unicodedata, warnings
from datetime import datetime
from io import StringIO
from pathlib import Path
import pandas as pd

warnings.filterwarnings("ignore")

BASE_DIR   = Path(__file__).parent
BRUTOS_DIR = BASE_DIR / "brutos"
REF_DIR    = BASE_DIR / "referencia"
SAIDA_DIR  = BASE_DIR / "saida"
for _d in (BRUTOS_DIR, REF_DIR, SAIDA_DIR):
    _d.mkdir(exist_ok=True)

PREFIXO_RELATORIO = "Enec_Status_Adesao_Detalhado"
CENSO_NOMES = ["Enec_escolas_2026_base_censo25.xlsx", "escolas_censo_25.xlsx"]
MODULO = "RENEDIC"

# Status do SIMEC → código usado no painel.
# Se aparecer um status novo no relatório, acrescente aqui (texto sem acento, minúsculo).
STATUS_MAP = {
    "nao iniciado":     "ni",
    "termo nao aceito": "tna",
    "termo aceito":     "ta",
    "em preenchimento": "ep",
    "em analise":       "ea",
}

# Redes municipais que não existem (DF não tem rede municipal; Noronha é distrito estadual)
SEM_REDE_MUNICIPAL = {"5300108", "2605459"}

UFS_IBGE = {"11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP","17":"TO",
            "21":"MA","22":"PI","23":"CE","24":"RN","25":"PB","26":"PE","27":"AL","28":"SE","29":"BA",
            "31":"MG","32":"ES","33":"RJ","35":"SP","41":"PR","42":"SC","43":"RS",
            "50":"MS","51":"MT","52":"GO","53":"DF"}


# =============================================================================
# UTILITÁRIOS
# =============================================================================

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)

def erro(msg):
    log("❌ " + msg)
    sys.exit(1)

def sem_acento(s):
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()

def achar_relatorio():
    cands = [p for p in BRUTOS_DIR.iterdir()
             if p.is_file() and p.name.startswith(PREFIXO_RELATORIO) and not p.name.startswith("~$")]
    if not cands:
        erro(f"Nenhum arquivo '{PREFIXO_RELATORIO}...' encontrado em brutos/")
    # O mais recente pela data/hora do nome; se o nome não tiver data, pela data do arquivo
    return max(cands, key=lambda p: (data_do_nome(p) or datetime.fromtimestamp(p.stat().st_mtime)))

def data_do_nome(path):
    """Enec_Status_Adesao_Detalhado_09-11-2026_13_06.xls → datetime(2026,11,9,13,6)"""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})[_ ](\d{2})[_h:](\d{2})", path.name)
    if not m:
        return None
    d, mo, a, h, mi = map(int, m.groups())
    try:
        return datetime(a, mo, d, h, mi)
    except ValueError:
        return None

def ler_tabela(path):
    """Lê pelo CONTEÚDO, não pela extensão: Excel (.xlsx), Excel antigo (.xls) ou HTML disfarçado."""
    cabeca = path.read_bytes()[:8]
    if cabeca.startswith(b"PK"):
        log("  Formato detectado: Excel (.xlsx)")
        return pd.read_excel(path, dtype=str, engine="openpyxl")
    if cabeca.startswith(b"\xd0\xcf\x11\xe0"):
        log("  Formato detectado: Excel antigo (.xls)")
        return pd.read_excel(path, dtype=str, engine="xlrd")
    log("  Formato detectado: HTML (falso Excel)")
    bruto = path.read_bytes()
    texto = None
    m = re.search(rb'charset=["\']?([\w-]+)', bruto[:3000], re.I)
    encs = ([m.group(1).decode()] if m else []) + ["utf-8", "latin1"]
    for enc in encs:
        try:
            texto = bruto.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    tabelas = pd.read_html(StringIO(texto), thousands=None)
    df = max(tabelas, key=len)
    # Algumas exportações trazem o cabeçalho como primeira linha de dados
    if not any("status" in sem_acento(c) for c in df.columns.astype(str)):
        for i in range(min(10, len(df))):
            if any("status" in sem_acento(v) for v in df.iloc[i].astype(str)):
                df.columns = df.iloc[i].astype(str)
                df = df.iloc[i + 1:].reset_index(drop=True)
                break
    return df.astype(str).replace({"nan": None, "None": None, "": None})

def coluna(df, *chaves, obrigatoria=True, exceto=()):
    for c in df.columns:
        n = sem_acento(c)
        if any(k in n for k in chaves) and not any(e in n for e in exceto):
            return c
    if obrigatoria:
        erro(f"Coluna não encontrada (procurei por: {', '.join(chaves)}). Colunas do arquivo: {list(df.columns)}")
    return None

def para_data(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "nat", "none", "-"):
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:19], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%Y-%m-%d")
    except Exception:
        return ""

def para_ibge(v):
    try:
        return str(int(float(str(v).strip())))
    except Exception:
        return ""


# =============================================================================
# 1. RELATÓRIO DE ADESÃO
# =============================================================================

def processar_relatorio():
    path = achar_relatorio()
    log(f"Relatório: {path.name}")
    df = ler_tabela(path)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.astype(object).where(pd.notna(df), None)
    df = df.dropna(how="all")

    # Dados pessoais: descartados imediatamente
    col_resp = coluna(df, "responsav", obrigatoria=False)
    if col_resp:
        df = df.drop(columns=[col_resp])
        log("  Coluna 'Responsável' descartada (não vai para o painel)")

    c_mod    = coluna(df, "modulo", obrigatoria=False)
    c_ibge   = coluna(df, "ibge")
    c_uf     = coluna(df, "estado", "uf")
    c_mun    = coluna(df, "municipio")
    c_esfera = coluna(df, "esfera", "dependencia")
    c_data   = coluna(df, "aceite", obrigatoria=False)
    c_status = coluna(df, "status")

    if c_mod:
        antes = len(df)
        df = df[df[c_mod].astype(str).str.strip().str.upper() == MODULO]
        if len(df) < antes:
            log(f"  Mantidas só as linhas do módulo {MODULO}: {len(df)} de {antes}")

    desconhecidos = set()
    redes = []
    for _, r in df.iterrows():
        if not r[c_status]:
            continue
        st_txt = sem_acento(r[c_status])
        st = STATUS_MAP.get(st_txt)
        if st is None:
            desconhecidos.add(str(r[c_status]))
            continue
        esfera = sem_acento(r[c_esfera] or "")
        esf = "E" if esfera.startswith("estad") else ("M" if esfera.startswith("munic") else "")
        if not esf:
            continue
        ibge = para_ibge(r[c_ibge])
        uf = str(r[c_uf] or "").strip().upper()
        if len(uf) != 2:
            uf = UFS_IBGE.get(ibge[:2], uf)
        if esf == "M" and ibge in SEM_REDE_MUNICIPAL:
            continue
        mun = str(r[c_mun]).strip() if esf == "M" and r[c_mun] not in (None, "nan") else ""
        redes.append({
            "e": esf, "uf": uf, "m": mun, "i": ibge if esf == "M" else "",
            "s": st, "a": para_data(r[c_data]) if c_data else "",
        })

    if desconhecidos:
        erro("Status desconhecido(s) no relatório: " + ", ".join(sorted(desconhecidos)) +
             "\n   Acrescente no STATUS_MAP do processar.py e rode de novo.")

    n_m = sum(1 for r in redes if r["e"] == "M")
    n_e = sum(1 for r in redes if r["e"] == "E")
    log(f"  Redes: {n_m} municipais | {n_e} estaduais")
    if n_m and abs(n_m - 5569) > 5:
        log(f"  ⚠️  Esperado ~5.569 redes municipais e vieram {n_m}. Confira se o relatório está completo.")
    for esf, nome in (("M", "Municipais"), ("E", "Estaduais")):
        cont = {}
        for r in redes:
            if r["e"] == esf:
                cont[r["s"]] = cont.get(r["s"], 0) + 1
        log(f"    {nome}: " + " | ".join(f"{k}={v}" for k, v in sorted(cont.items())))

    data_rel = data_do_nome(path) or datetime.fromtimestamp(path.stat().st_mtime)
    return redes, data_rel, path.name


# =============================================================================
# 2. CENSO — filtros da equipe (por rede municipal)
# =============================================================================

def processar_censo():
    path = next((REF_DIR / n for n in CENSO_NOMES if (REF_DIR / n).exists()), None)
    if path is None:
        log("⚠️  Base do Censo não encontrada em referencia/ — filtros da equipe ficarão vazios")
        return {}
    cache = REF_DIR / "_censo_cache.json"
    assinatura = f"{path.name}|{path.stat().st_size}|{int(path.stat().st_mtime)}"
    if cache.exists():
        try:
            c = json.loads(cache.read_text(encoding="utf-8"))
            if c.get("assinatura") == assinatura:
                log(f"Censo: {path.name} (já processado antes, usando cache)")
                return c["flags"]
        except Exception:
            pass

    log(f"Censo: {path.name} (lendo, pode levar alguns segundos)")
    df = pd.read_excel(path, usecols=lambda c: c in {
        "CO_MUNICIPIO", "TP_DEPENDENCIA", "TP_LOCALIZACAO", "TP_LOCALIZACAO_DIFERENCIADA",
        "ESCOLA_INDIGENA", "ESCOLA_QUILOMBOLA", "ESCOLA_DO_CAMPO"})
    df = df[df["TP_DEPENDENCIA"].astype(str).str.strip() == "Municipal"].copy()
    df["co"] = df["CO_MUNICIPIO"].apply(para_ibge)
    df["rural"] = (df["TP_LOCALIZACAO"].astype(str).str.strip() == "Rural").astype(int)
    dif = df["TP_LOCALIZACAO_DIFERENCIADA"].astype(str)
    df["assent"] = dif.str.contains("assentamento", case=False).astype(int)
    df["tradic"] = dif.str.contains("tradicionais", case=False).astype(int)
    g = df.groupby("co")
    tab = pd.DataFrame({
        "q":  g["ESCOLA_QUILOMBOLA"].max(),
        "i":  g["ESCOLA_INDIGENA"].max(),
        "c":  g["ESCOLA_DO_CAMPO"].max(),
        "r":  g["rural"].max(),
        "mr": (g["rural"].mean() > 0.5).astype(int),
        "as": g["assent"].max(),
        "tr": g["tradic"].max(),
    }).fillna(0)
    flags = {co: [k for k in tab.columns if int(v[k]) == 1] for co, v in tab.iterrows()}
    flags = {k: v for k, v in flags.items() if v}
    cache.write_text(json.dumps({"assinatura": assinatura, "flags": flags}), encoding="utf-8")
    log(f"  {len(tab)} redes municipais no Censo")
    return flags


# =============================================================================
# 3. HISTÓRICO DE ENVIOS (curva de evolução)
# =============================================================================

def atualizar_envios(redes, data_rel):
    """
    Guarda, para cada rede hoje 'Em análise', o dia em que ela apareceu assim pela 1ª vez.
    - Primeiro processamento: usa a data do termo aceito (para não perder o histórico).
    - Depois: usa a data do relatório em que a rede apareceu 'Em análise'.
    - Se a rede sair de 'Em análise', sai da curva.
    Várias atualizações no mesmo dia contam como um único dia.
    """
    arq = SAIDA_DIR / "envios.json"
    primeira = not arq.exists()
    envios = {} if primeira else json.loads(arq.read_text(encoding="utf-8"))
    hoje = data_rel.strftime("%Y-%m-%d")
    novo = {}
    novos = 0
    for r in redes:
        if r["s"] != "ea":
            continue
        chave = r["e"] + (r["i"] or r["uf"])
        if chave in envios:
            novo[chave] = envios[chave]
        else:
            novo[chave] = (r["a"] or hoje) if primeira else hoje
            novos += 1
    if primeira:
        log(f"  Histórico iniciado com {len(novo)} redes (datas do termo aceito)")
    else:
        saidas = len(set(envios) - set(novo))
        log(f"  Envios: +{novos} novas redes em análise" + (f" | {saidas} saíram de 'Em análise'" if saidas else ""))
    arq.write_text(json.dumps(novo, separators=(",", ":"), sort_keys=True), encoding="utf-8")
    return novo


# =============================================================================
# 4. CONFIGURAÇÃO (relatorio.txt)
# =============================================================================

MODELO_CONFIG = """##CONFIG
MENSAGEM: Atenção, articuladores: acompanhem as redes que ainda não iniciaram a adesão e as que aceitaram o termo mas ainda não indicaram o Ponto Focal.
PRAZO_MUNICIPAL: 13/11/2026
PRAZO_ESTADUAL: 29/10/2026
CHAVE_ADMIN: troque-esta-chave
##FIMCONFIG
"""

def processar_config():
    src = BASE_DIR / "relatorio.txt"
    if not src.exists():
        log("⚠️  relatorio.txt não encontrado — criando modelo (edite e rode de novo)")
        src.write_text(MODELO_CONFIG, encoding="utf-8")
    cfg = {}
    for linha in src.read_text(encoding="utf-8-sig").splitlines():
        if ":" in linha and not linha.startswith("##"):
            k, v = linha.split(":", 1)
            cfg[k.strip().upper()] = v.strip()

    def iso(d):
        try:
            return datetime.strptime(d, "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            return ""

    chave = cfg.get("CHAVE_ADMIN", "")
    if not chave or chave == "troque-esta-chave":
        log("⚠️  Defina uma CHAVE_ADMIN própria no relatorio.txt (a chave modelo não é segura)")
    out = {
        "mensagem": cfg.get("MENSAGEM", ""),
        "prazo_municipal": iso(cfg.get("PRAZO_MUNICIPAL", "")),
        "prazo_estadual": iso(cfg.get("PRAZO_ESTADUAL", "")),
        # Só o "resumo" (hash) da chave vai para o site; a chave em si fica no seu computador
        "admin_hash": hashlib.sha256(chave.encode("utf-8")).hexdigest() if chave else "",
    }
    (SAIDA_DIR / "config.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"  Prazos: municipal {cfg.get('PRAZO_MUNICIPAL','—')} | estadual {cfg.get('PRAZO_ESTADUAL','—')}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    log("=" * 60)
    log("PAINEL RENEDIC — PROCESSAMENTO")
    log("=" * 60)

    meta_arq = SAIDA_DIR / "meta.json"
    meta_ant = json.loads(meta_arq.read_text(encoding="utf-8")) if meta_arq.exists() else {}

    log("\n▶ Relatório de adesão...")
    redes, data_rel, nome_arq = processar_relatorio()

    ant = meta_ant.get("relatorio_iso")
    if ant and data_rel.strftime("%Y-%m-%dT%H:%M") < ant:
        erro(f"O relatório em brutos/ ({data_rel:%d/%m/%Y %H:%M}) é MAIS ANTIGO que o já publicado "
             f"({meta_ant.get('atualizado_em')}). Nada foi alterado.")

    log("\n▶ Base do Censo (filtros da equipe)...")
    flags = processar_censo()
    com_flag = 0
    for r in redes:
        if r["e"] == "M" and r["i"] in flags:
            r["f"] = flags[r["i"]]
            com_flag += 1
    log(f"  {com_flag} redes municipais com marcação do Censo")

    log("\n▶ Curva de evolução...")
    envios = atualizar_envios(redes, data_rel)
    for r in redes:
        chave = r["e"] + (r["i"] or r["uf"])
        if chave in envios:
            r["d"] = envios[chave]

    (SAIDA_DIR / "redes.json").write_text(
        json.dumps(redes, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log("  ✅ redes.json salvo")

    log("\n▶ Configuração (relatorio.txt)...")
    processar_config()

    meta = {
        "atualizado_em": data_rel.strftime("%d/%m/%Y às %H:%M"),
        "relatorio_iso": data_rel.strftime("%Y-%m-%dT%H:%M"),
        "arquivo": nome_arq,
        "processado_em": datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }
    meta_arq.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")

    log("\n" + "=" * 60)
    log(f"CONCLUÍDO — dados de {meta['atualizado_em']}")
    log("Arquivos em saida/: redes.json | envios.json | config.json | meta.json")
    log("=" * 60)


if __name__ == "__main__":
    main()
