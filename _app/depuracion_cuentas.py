"""
_app/depuracion_cuentas.py — Motor del módulo DEPURACIÓN DE CUENTAS.

Lee el "Auxiliar de cuentas" exportado del ERP (hoja Datos, encabezados en la fila 12)
y hace la depuración de cualquier cuenta (proveedores, clientes, cuentas por cobrar, deudores…):

  1. Integridad aritmética (saldo corrido y saldo final por cuenta).
  2. Relación cargo-abono: cada pago (EA) y nota de crédito se liga a su factura (PRO)
     por folio / UUID / importe; cada factura queda liquidada, parcial o sin pago.
  3. Hallazgos: saldos deudores, NC sin factura, pagos duplicados o en exceso,
     egresos manuales, NC altas vs facturado, catálogo vs CFDI, pagos antes de factura,
     saldos iniciales sin liquidar, antigüedad, concentración y observaciones menores.
  4. Salidas Excel: auxiliar original marcado (amarillo = relacionado, azul = sin relación)
     y papel de trabajo de auditoría.
"""
from __future__ import annotations

import io
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_ROW = 12            # fila de encabezados en la hoja Datos
FIRST_DATA_ROW = HEADER_ROW + 1
COLS_REQ = ["Sucursal", "Periodo", "Tipo póliza", "Num. Póliza", "Fecha", "Cta.", "Cta. Desc.",
            "Desc. Detalle", "Desc. Póliza", "Cargo", "Abono", "Saldo", "Ren."]

AMARILLO = "FFFF00"
AZUL = "9DC3E6"
MON = '#,##0.00;[Red]-#,##0.00;-'
PCT = '0.0%'
DT = 'dd/mm/yyyy'
F = "Arial"

_RFC = r'[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{3}'


def _fmt(v: float) -> str:
    return f"${v:,.2f}"


def _norm(s) -> str:
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'[^A-Z0-9]', '', s)


# ═══════════════════════════════════════════════════════════════════════════════
# LECTURA
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass
class Resultado:
    meta: dict
    cuentas: pd.DataFrame          # una fila por cuenta
    mov: pd.DataFrame              # movimientos con columnas derivadas
    relacion: pd.DataFrame         # row, color (Y/B), motivo
    hallazgos: list = field(default_factory=list)
    tablas: dict = field(default_factory=dict)
    errores_aritmeticos: int = 0


def _leer_meta(wb) -> dict:
    ws = wb["Datos"] if "Datos" in wb.sheetnames else wb.worksheets[0]
    meta = {}
    for r in range(1, HEADER_ROW):
        k, v = ws.cell(r, 1).value, ws.cell(r, 2).value
        if k:
            k = str(k).strip().rstrip(':')
            if k in meta:      # el ERP repite "Cuenta inicial" para la cuenta final
                k = k.replace("inicial", "final") + " (rotulada 'Cuenta inicial')"
            meta[k] = v
    return meta


def leer_auxiliar(data: bytes) -> tuple[pd.DataFrame, dict]:
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    meta = _leer_meta(wb)
    hoja = "Datos" if "Datos" in wb.sheetnames else wb.sheetnames[0]
    wb.close()
    df = pd.read_excel(io.BytesIO(data), sheet_name=hoja, header=HEADER_ROW - 1)
    falt = [c for c in COLS_REQ if c not in df.columns]
    if falt:
        raise ValueError("El archivo no tiene el formato de 'Auxiliar de cuentas'. "
                         f"Faltan columnas: {', '.join(falt)}")
    df["row"] = df.index + FIRST_DATA_ROW
    for c in ("Cargo", "Abono", "Saldo"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    return df, meta


# ═══════════════════════════════════════════════════════════════════════════════
# ANÁLISIS
# ═══════════════════════════════════════════════════════════════════════════════
def analizar(data: bytes) -> Resultado:
    df, meta = leer_auxiliar(data)

    # ── bloques por cuenta + integridad ──────────────────────────────────────
    accts: dict = {}
    for i, r in df.iterrows():
        c = r["Cta."]
        if pd.isna(c):
            continue
        if pd.isna(r["Sucursal"]):
            dd = str(r["Desc. Detalle"])
            if dd.startswith("Saldo final") and c in accts:
                accts[c]["fin"] = (r["Cargo"], r["Abono"], r["Saldo"], int(r["row"]))
            elif dd != "nan" and dd.startswith(str(c)):
                accts.setdefault(c, {"name": str(r["Cta. Desc."]).strip(), "ini": float(r["Saldo"] or 0), "mov": []})
        else:
            accts.setdefault(c, {"name": str(r["Cta. Desc."]).strip(), "ini": 0.0, "mov": []})
            accts[c]["mov"].append(i)

    def _corrido(a, signo):
        """Errores del saldo corrido con signo +1 = acreedora (saldo + abono − cargo), −1 = deudora."""
        s, err = a["ini"], 0
        for i in a["mov"]:
            s = s + signo * ((df.at[i, "Abono"] or 0) - (df.at[i, "Cargo"] or 0))
            if abs(s - df.at[i, "Saldo"]) > 0.01:
                err += 1
                s = df.at[i, "Saldo"]
        return err

    errores = 0
    filas = []
    for c, a in accts.items():
        e_acr, e_deu = _corrido(a, 1), _corrido(a, -1)
        if e_acr != e_deu:
            nat = "Acreedora" if e_acr < e_deu else "Deudora"
        else:      # sin movimientos que lo definan: por el rubro de la cuenta
            nat = "Deudora" if str(c)[:1] in ("1", "5", "6", "7") else "Acreedora"
        signo = 1 if nat == "Acreedora" else -1
        errores += min(e_acr, e_deu)
        mm = df.loc[a["mov"]]
        fc, fa, fs, fr = a.get("fin", (mm.Cargo.sum(), mm.Abono.sum(), None, None))
        if fs is None:
            fs = a["ini"] + signo * (mm.Abono.sum() - mm.Cargo.sum())
        if abs(a["ini"] + signo * (mm.Abono.sum() - mm.Cargo.sum()) - fs) > 0.01:
            errores += 1
        filas.append(dict(cta=c, nombre=a["name"], ini=a["ini"], cargos=fc, abonos=fa, fin=fs,
                          frow=fr, n=len(a["mov"]), naturaleza=nat))
    cuentas = pd.DataFrame(filas).set_index("cta")

    # ── movimientos y campos derivados ───────────────────────────────────────
    m = df[df["Sucursal"].notna()].copy()
    # Vista interna "acreedora": Abono = aumento del saldo, Cargo = disminución.
    # En cuentas deudoras (clientes, deudores, anticipos…) se intercambian; los originales quedan en *_aux.
    m["cargo_aux"], m["abono_aux"] = m["Cargo"], m["Abono"]
    m["naturaleza"] = m["Cta."].map(cuentas.naturaleza)
    deu = m.naturaleza == "Deudora"
    m.loc[deu, ["Cargo", "Abono"]] = m.loc[deu, ["abono_aux", "cargo_aux"]].values
    d = m["Desc. Detalle"].fillna("").astype(str)
    m["isnc"] = d.str.contains("NT.CR", regex=False)
    m["folio"] = d.str.extract(r'PRVS(?:\.NT\.CR)?:\s*(.*?)\s+PROV:')[0].str.strip()
    m["rfc"] = d.str.extract(r'PROV:\s*\[?(' + _RFC + ')')[0]
    m["pname"] = d.str.extract(r'PROV:\s*\[?' + _RFC + r'-(.*?)\]?\s+ABONO PROV')[0]
    m["amt"] = m["Cargo"].fillna(0) + m["Abono"].fillna(0)
    m["xml"] = d.str.startswith("[P:")
    m["pago"] = d.str.extract(r'\[P:(\d+)\]')[0]
    m["uuid5"] = d.str.extract(r'UUID\):\s*(\w+)')[0].str.upper()
    m["flist"] = (m["Desc. Póliza"].fillna("").astype(str).str.extract(r'UUID\):\s*(.*)$')[0].fillna("")
                  .apply(lambda s: [x.strip() for x in s.split(",") if x.strip()]))

    rel = _relacionar(m, cuentas, meta)
    res = Resultado(meta=meta, cuentas=cuentas, mov=m, relacion=rel, errores_aritmeticos=errores)
    _hallazgos(res)
    return res


def _relacionar(m: pd.DataFrame, cuentas: pd.DataFrame, meta: dict) -> pd.DataFrame:
    corte = meta.get("Fecha final")
    corte_txt = corte.strftime("%d-%m-%Y") if hasattr(corte, "strftime") else "la fecha de corte"
    es_pro = (m["Tipo póliza"] == "PRO") & m.folio.notna()
    inv = m[es_pro & ~m.isnc & (m.Abono > 0)]
    allea = m[(m["Tipo póliza"] == "EA") & m.xml]
    invamt = inv.groupby(["Cta.", "folio"]).Abono.apply(list)

    # Folio de cada renglón de pago
    fol, u2f = {}, {}
    for _, r in allea[allea.flist.str.len() == 1].iterrows():
        u2f[(r["Cta."], r.uuid5)] = r.flist[0]
    for (c, p), g in allea.groupby(["Cta.", "pago"]):
        g = g.sort_values("Ren.")
        fl = g.flist.iloc[0]
        if len(g) == len(fl):
            for k, (i, r) in enumerate(g.iterrows()):
                fol[i] = fl[k]
        else:
            used = set()
            for i, r in g.iterrows():
                ff = u2f.get((c, r.uuid5))
                if ff is None and pd.notna(r["Ren."]) and 0 < int(r["Ren."]) <= len(fl):
                    fr = fl[int(r["Ren."]) - 1]     # el renglón de la póliza indica la posición del folio pagado
                    if fr not in used and (c, fr) in invamt.index:
                        ff = fr
                if ff is None:
                    cands = [x for x in fl if x not in used and (c, x) in invamt.index
                             and any(abs(a - r.Cargo) < 0.01 for a in invamt[(c, x)])]
                    ff = cands[0] if cands else None
                fol[i] = ff
                if ff:
                    used.add(ff)
    m["lf"] = pd.Series(fol)

    ea = m[(m["Tipo póliza"] == "EA") & (m.Cargo > 0) & m.xml]
    neg = m[(m["Tipo póliza"] == "EA") & (m.Cargo < 0)]
    # Movimientos sin formato XML de proveedores (clientes, deudores, otras cuentas, capturas manuales):
    # aumentos = documentos; disminuciones se relacionan por importe y referencia
    gen_doc = m[~(es_pro | m.isnc) & ~((m["Tipo póliza"] == "EA") & m.xml) & (m.Abono > 0)]
    gen_dis = m[~m.isnc & ~((m["Tipo póliza"] == "EA") & m.xml) & ~es_pro & (m.Cargo > 0)]
    st: dict = {}
    # importes por renglón: [cargo conciliado vs facturas, cargo vs saldo inicial, cargo pendiente,
    #                        abono conciliado, abono pendiente]
    imp: dict = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, 0.0])
    # grupos de conciliación (unión de renglones que se cancelan entre sí)
    padre: dict = {}

    def _raiz(x):
        padre.setdefault(x, x)
        while padre[x] != x:
            padre[x] = padre[padre[x]]
            x = padre[x]
        return x

    def _unir(a, b):
        ra, rb = _raiz(a), _raiz(b)
        if ra != rb:
            padre[max(ra, rb)] = min(ra, rb)
    grupo_si: dict = {}

    for r in m[(m.Cargo.fillna(0) == 0) & (m.Abono.fillna(0) == 0)].row:
        st[r] = ("B", "Renglón con importe cero: no se relaciona con ningún movimiento")
    for _, r in m[~m.isnc & (m.Abono < 0)].iterrows():
        st[r.row] = ("B", "Movimiento negativo que disminuye documentos: revisar cancelación")
        imp[r.row][4] = r.Abono
    for _, r in m[~m.isnc & (m["Tipo póliza"] != "EA") & (m.Cargo < 0)].iterrows():
        st[r.row] = ("B", "Movimiento negativo: revisar cancelación o reclasificación")
        imp[r.row][2] = r.Cargo

    invd: dict = {}
    for _, r in inv.sort_values("Fecha").iterrows():
        invd.setdefault((r["Cta."], r.folio), []).append(
            dict(row=r.row, amt=r.Abono, paid=0.0, links=[], date=r.Fecha))
    gdocs: dict = defaultdict(list)          # documentos genéricos por cuenta
    for _, r in gen_doc.sort_values(["Fecha", "row"]).iterrows():
        gdocs[r["Cta."]].append(dict(row=r.row, amt=r.Abono, paid=0.0, links=[], date=r.Fecha,
                                     toks=_tokens(r["Desc. Detalle"], r["Desc. Póliza"], r["Num. Póliza"])))

    ini_used = defaultdict(float)
    for _, r in ea[ea.xml].sort_values(["Fecha", "row"]).iterrows():
        k = (r["Cta."], r.lf)
        if k in invd:
            lst = invd[k]
            ex = [x for x in lst if abs((x["amt"] - x["paid"]) - r.Cargo) < 0.01]
            left, touched = r.Cargo, []
            for x in ex + [x for x in lst if x not in ex]:
                rem = x["amt"] - x["paid"]
                if rem <= 0.005 or left <= 0.005:
                    continue
                t = min(rem, left)
                x["paid"] += t
                left -= t
                x["links"].append(r.row)
                touched.append(x["row"])
            imp[r.row][0] = r.Cargo - left
            imp[r.row][2] = left
            if not touched:
                st[r.row] = ("B", f"Pago en exceso: la factura {r.lf} ya estaba liquidada (pago duplicado)")
            elif left > 0.01:
                st[r.row] = ("B", f"Pago mayor al saldo pendiente de la factura {r.lf}; exceso {_fmt(left)} sin abono relacionado")
            else:
                extra = f" (diferencia de redondeo {_fmt(left)})" if left > 0.004 else ""
                st[r.row] = ("Y", f"Pago relacionado con factura {r.lf} (renglón {', '.join(map(str, touched))}){extra}")
        else:
            c = r["Cta."]
            cand = [(fo, x) for (cc, fo), lst in invd.items() if cc == c and str(fo).endswith("-0")
                    for x in lst if abs((x["amt"] - x["paid"]) - r.Cargo) < 0.01]
            if cand:
                fo, x = min(cand, key=lambda t: abs((t[1]["date"] - r.Fecha).days))
                x["paid"] += r.Cargo
                x["links"].append(r.row)
                st[r.row] = ("Y", f"Pago relacionado por UUID e importe con factura sin folio (renglón {x['row']})")
                imp[r.row][0] = r.Cargo
                continue
            ini = max(cuentas.loc[c, "ini"], 0) if c in cuentas.index else 0
            if ini_used[c] + r.Cargo <= ini + 0.01:
                ini_used[c] += r.Cargo
                grupo_si[r.row] = c
                imp[r.row][1] = r.Cargo
                st[r.row] = ("Y", f"Pago de factura {r.lf} del saldo inicial (ejercicio anterior)")
            else:
                imp[r.row][2] = r.Cargo
                st[r.row] = ("B", f"Pago de factura {r.lf} que no está en el auxiliar y excede el saldo inicial: sin abono relacionado")

    # Disminuciones genéricas: 1) mismo importe pendiente y referencia común, 2) mismo importe,
    # 3) aplicadas al saldo inicial, 4) sin contrapartida
    for _, r in gen_dis.sort_values(["Fecha", "row"]).iterrows():
        c = r["Cta."]
        tk = _tokens(r["Desc. Detalle"], r["Desc. Póliza"], r["Num. Póliza"])
        cands = [x for x in gdocs.get(c, []) if abs((x["amt"] - x["paid"]) - r.Cargo) < 0.01]
        if cands:
            x = min(cands, key=lambda x: (0 if tk & x["toks"] else 1, 0 if x["date"] <= r.Fecha else 1,
                                          abs((x["date"] - r.Fecha).days)))
            x["paid"] += r.Cargo
            x["links"].append(r.row)
            imp[r.row][0] = r.Cargo
            ref = " y referencia" if tk & x["toks"] else ""
            st[r.row] = ("Y", f"Relacionado por importe{ref} con el renglón {x['row']}")
            continue
        ini = max(cuentas.loc[c, "ini"], 0) if c in cuentas.index else 0
        if ini_used[c] + r.Cargo <= ini + 0.01:
            ini_used[c] += r.Cargo
            grupo_si[r.row] = c
            imp[r.row][1] = r.Cargo
            st[r.row] = ("Y", "Aplicado al saldo inicial (ejercicio anterior)")
            continue
        imp[r.row][2] = r.Cargo
        if r["Tipo póliza"] == "EA":
            st[r.row] = ("B", "Egreso capturado a mano sin referencia a factura (CFDI)")
        else:
            st[r.row] = ("B", "Sin contrapartida: no coincide con ningún movimiento pendiente de la cuenta")

    # NC aplicadas dentro del pago (cargo negativo por el mismo importe)
    ncrows = m[m.isnc & (m.Cargo > 0)].sort_values("Fecha")
    usedneg, paired = set(), set()
    for _, r in ncrows.iterrows():
        c = neg[(neg["Cta."] == r["Cta."]) & ((neg.Cargo + r.Cargo).abs() < 0.01) & ~neg.row.isin(usedneg)]
        if len(c):
            c = c.assign(dd=(c.Fecha - r.Fecha).abs()).sort_values("dd").iloc[0]
            usedneg.add(c.row)
            _unir(r.row, c.row)
            paired.add(r.row)
            st[r.row] = ("Y", f"Nota de crédito aplicada en el pago (renglón {c.row}, cargo negativo por el mismo importe)")
            st[c.row] = ("Y", f"Aplicación de la nota de crédito del renglón {r.row} dentro del pago")
            imp[r.row][0], imp[c.row][0] = r.Cargo, c.Cargo      # se cancelan entre sí (suman cero)
    for _, r in neg[~neg.row.isin(usedneg)].iterrows():
        st[r.row] = ("B", "Cargo negativo sin nota de crédito relacionada")
        imp[r.row][2] = r.Cargo
    for _, r in ncrows[~ncrows.row.isin(paired)].iterrows():
        best = None
        for (cc, fo), lst in invd.items():
            if cc != r["Cta."]:
                continue
            for x in lst:
                rem = x["amt"] - x["paid"]
                if rem > 0.005 and abs(rem - r.Cargo) < 0.01 and x["date"] <= r.Fecha \
                        and (best is None or x["date"] > best[1]["date"]):
                    best = (fo, x)
        if best:
            best[1]["paid"] += r.Cargo
            best[1]["links"].append(r.row)
            imp[r.row][0] = r.Cargo
            st[r.row] = ("Y", f"Nota de crédito aplicada por el saldo exacto de la factura {best[0]} (renglón {best[1]['row']}); "
                              "la NC no trae la referencia, se relacionó por importe")
        else:
            imp[r.row][2] = r.Cargo
            st[r.row] = ("B", "Nota de crédito sin factura relacionada: no coincide con el saldo de ninguna factura del proveedor")

    for (c, fo), lst in invd.items():
        for x in lst:
            for lk in x["links"]:
                _unir(x["row"], lk)
            rem = x["amt"] - x["paid"]
            imp[x["row"]][3] = x["paid"]
            imp[x["row"]][4] = rem
            if rem <= 0.01:
                extra = f" (diferencia de redondeo {_fmt(rem)})" if rem > 0.004 else ""
                st[x["row"]] = ("Y", f"Factura liquidada con {len(x['links'])} cargo(s) relacionado(s){extra}")
            elif x["paid"] > 0:
                st[x["row"]] = ("B", f"Factura con pago parcial: pagado {_fmt(x['paid'])}, pendiente {_fmt(rem)} al {corte_txt}")
            else:
                st[x["row"]] = ("B", f"Factura sin ningún pago ni nota de crédito relacionados (pendiente al {corte_txt})")
    for c, lst in gdocs.items():
        for x in lst:
            for lk in x["links"]:
                _unir(x["row"], lk)
            rem = x["amt"] - x["paid"]
            imp[x["row"]][3] = x["paid"]
            imp[x["row"]][4] = rem
            if rem <= 0.01:
                st[x["row"]] = ("Y", f"Saldo liquidado con {len(x['links'])} movimiento(s) relacionado(s)")
            elif x["paid"] > 0:
                st[x["row"]] = ("B", f"Aplicación parcial: aplicado {_fmt(x['paid'])}, pendiente {_fmt(rem)} al {corte_txt}")
            else:
                st[x["row"]] = ("B", f"Sin contrapartida relacionada (pendiente al {corte_txt})")
    for r in m.row:
        st.setdefault(r, ("B", "Movimiento sin relación identificable"))

    def _gkey(r):
        if r in grupo_si:
            return f"SI|{grupo_si[r]}"
        return f"R|{_raiz(r)}"
    nat = dict(zip(m.row, m.naturaleza))
    filas = []
    for r, (c, t) in st.items():
        dis_c, dis_si, dis_p, aum_c, aum_p = imp[r]
        if nat.get(r) == "Deudora":      # regresar a los lados reales: aumento = cargo, disminución = abono
            filas.append((r, c, t, aum_c, 0.0, aum_p, dis_c, dis_si, dis_p, _gkey(r)))
        else:
            filas.append((r, c, t, dis_c, dis_si, dis_p, aum_c, 0.0, aum_p, _gkey(r)))
    rel = pd.DataFrame(filas, columns=["row", "color", "motivo", "cargo_conciliado", "cargo_saldo_inicial",
                                       "cargo_pendiente", "abono_conciliado", "abono_saldo_inicial",
                                       "abono_pendiente", "grupo_key"])
    rel = rel.merge(m[["row", "Cta.", "Cta. Desc.", "Tipo póliza", "Num. Póliza", "Fecha", "cargo_aux", "abono_aux", "lf",
                       "naturaleza"]].rename(columns={"cargo_aux": "Cargo", "abono_aux": "Abono"}), on="row")
    rel["categoria"] = rel.motivo.map(_categoria)
    return rel.sort_values("row").reset_index(drop=True)


CATEGORIAS = [
    ("Y", "Pago relacionado con factura"), ("Y", "Pago relacionado por UUID"), ("Y", "Pago de factura del saldo inicial"),
    ("Y", "Factura liquidada"), ("Y", "Nota de crédito aplicada"), ("Y", "Aplicación de nota de crédito en pago"),
    ("B", "Nota de crédito sin factura relacionada"), ("B", "Pago sin factura en el auxiliar"),
    ("B", "Pago en exceso / duplicado"), ("B", "Pago mayor al saldo de la factura"), ("B", "Egreso capturado a mano"),
    ("B", "Cargo negativo sin NC"), ("B", "Abono negativo"), ("B", "Factura sin pago"),
    ("B", "Factura con pago parcial"), ("B", "Importe cero"),
    ("Y", "Relacionado por importe"), ("Y", "Aplicado al saldo inicial"), ("Y", "Movimiento liquidado"),
    ("B", "Movimiento sin contrapartida"), ("B", "Movimiento con aplicación parcial"), ("B", "Movimiento negativo"),
    ("B", "Sin relación identificable"),
]
_CAT_PATS = [
    ("Pago relacionado con factura", "Pago relacionado con factura"), ("Pago relacionado por UUID", "Pago relacionado por UUID"),
    ("del saldo inicial", "Pago de factura del saldo inicial"), ("Factura liquidada", "Factura liquidada"),
    ("Nota de crédito aplicada", "Nota de crédito aplicada"), ("Aplicación de la nota", "Aplicación de nota de crédito en pago"),
    ("Nota de crédito sin", "Nota de crédito sin factura relacionada"), ("que no está en el auxiliar", "Pago sin factura en el auxiliar"),
    ("Pago en exceso", "Pago en exceso / duplicado"), ("Pago mayor al saldo", "Pago mayor al saldo de la factura"),
    ("Egreso capturado a mano", "Egreso capturado a mano"), ("Cargo negativo sin", "Cargo negativo sin NC"),
    ("Abono negativo", "Abono negativo"), ("Factura sin ningún pago", "Factura sin pago"),
    ("Factura con pago parcial", "Factura con pago parcial"), ("importe cero", "Importe cero"),
    ("Relacionado por importe", "Relacionado por importe"), ("Aplicado al saldo inicial", "Aplicado al saldo inicial"),
    ("Saldo liquidado con", "Movimiento liquidado"), ("Sin contrapartida", "Movimiento sin contrapartida"),
    ("Aplicación parcial", "Movimiento con aplicación parcial"), ("Movimiento negativo", "Movimiento negativo"),
]


def _tokens(*textos) -> set:
    """Referencias (folios, UUID, números de documento) para relacionar movimientos genéricos."""
    out = set()
    for t in textos:
        for tok in re.findall(r"[A-Za-z0-9\-/]{4,}", str(t or "")):
            tok = tok.strip("-/").upper()
            if len(tok) >= 4 and any(ch.isdigit() for ch in tok) and not re.fullmatch(r"20\d\d(0[1-9]|1[0-2])?", tok):
                out.add(tok)
    return out


def _categoria(motivo: str) -> str:
    for pat, cat in _CAT_PATS:
        if pat in motivo:
            return cat
    return "Sin relación identificable"


# ═══════════════════════════════════════════════════════════════════════════════
# HALLAZGOS
# ═══════════════════════════════════════════════════════════════════════════════
def _hallazgos(res: Resultado) -> None:
    m, f, rel = res.mov, res.cuentas, res.relacion.copy()
    # importe del movimiento sin importar el lado (en cuentas deudoras el pago / NC va en el abono)
    rel["Cargo"] = rel.Cargo.fillna(0) + rel.Abono.fillna(0)
    corte = res.meta.get("Fecha final")
    corte = pd.Timestamp(corte) if corte is not None else m.Fecha.max()
    H, T = [], {}

    # vista interna: Abono = aumentos (documentos), Cargo = disminuciones (pagos / cobros y notas de crédito)
    fac = m.groupby("Cta.").Abono.sum()
    ncs = m[m.isnc].groupby("Cta.").Cargo.sum()
    pag = m[~m.isnc].groupby("Cta.").Cargo.sum()
    f["facturado"] = fac.reindex(f.index).fillna(0)
    f["notas_credito"] = ncs.reindex(f.index).fillna(0)
    f["pagos"] = pag.reindex(f.index).fillna(0)

    # 1. Saldos deudores
    deu = f[f.fin < -0.005].sort_values("fin")
    T["saldos_deudores"] = deu.reset_index()[["cta", "nombre", "naturaleza", "ini", "facturado", "pagos", "notas_credito", "fin"]]
    if len(deu):
        H.append(dict(riesgo="Alto", titulo=f"{len(deu)} cuentas con saldo contrario a su naturaleza",
                      importe=-deu.fin.sum(), hoja="Saldos contrarios",
                      detalle="Cargos mayores a los abonos: pagos o notas de crédito sin factura. "
                              + "; ".join(f"{r.nombre} {_fmt(r.fin)}" for _, r in deu.head(5).iterrows()),
                      accion="Aclarar con el tercero y reclasificar (anticipos, cuentas por cobrar o por pagar)"))

    # 2. NC sin factura relacionada
    nc_sin = rel[rel.categoria == "Nota de crédito sin factura relacionada"]
    if len(nc_sin):
        g = nc_sin.groupby(["Cta.", "Cta. Desc."]).agg(notas=("row", "size"), importe=("Cargo", "sum")).reset_index()
        g = g.sort_values("importe", ascending=False)
        T["nc_sin_factura"] = g
        H.append(dict(riesgo="Alto", titulo=f"{len(nc_sin)} notas de crédito sin factura relacionada",
                      importe=nc_sin.Cargo.sum(), hoja="NC sin factura",
                      detalle="; ".join(f"{r['Cta. Desc.']}: {r.notas} NC por {_fmt(r.importe)}" for _, r in g.head(5).iterrows()),
                      accion="Validar los CFDI en el SAT (y lista 69-B) y localizar la factura de origen"))

    # 3. Pagos duplicados / en exceso
    exc = rel[rel.categoria.isin(["Pago en exceso / duplicado", "Pago mayor al saldo de la factura"])]
    T["pagos_exceso"] = exc[["row", "Cta.", "Cta. Desc.", "Num. Póliza", "Fecha", "Cargo", "motivo"]]
    if len(exc):
        H.append(dict(riesgo="Alto", titulo=f"{len(exc)} pagos duplicados o mayores a la factura",
                      importe=exc.Cargo.sum(), hoja="Pagos en exceso",
                      detalle="; ".join(f"{r['Cta. Desc.']} {_fmt(r.Cargo)} ({r.Fecha:%d-%m-%Y})" for _, r in exc.head(5).iterrows()),
                      accion="Solicitar reembolso o aplicar a facturas posteriores"))

    # 4. Egresos manuales
    man = rel[rel.categoria == "Egreso capturado a mano"]
    T["egresos_manuales"] = man[["row", "Cta.", "Cta. Desc.", "Num. Póliza", "Fecha", "Cargo"]].merge(
        m[["row", "Desc. Detalle"]], on="row")
    if len(man):
        H.append(dict(riesgo="Alto", titulo=f"{len(man)} egresos capturados a mano sin CFDI",
                      importe=man.Cargo.sum(), hoja="Egresos manuales",
                      detalle="; ".join(f"{r['Cta. Desc.']} {_fmt(r.Cargo)}" for _, r in man.head(5).iterrows()),
                      accion="Obtener los CFDI y restringir la captura manual de egresos"))

    # 5. Pagos sin factura en el auxiliar
    psf = rel[rel.categoria == "Pago sin factura en el auxiliar"]
    T["pagos_sin_factura"] = psf[["row", "Cta.", "Cta. Desc.", "Num. Póliza", "Fecha", "Cargo", "lf"]]
    if len(psf):
        H.append(dict(riesgo="Medio", titulo=f"{len(psf)} pagos a facturas que no están en el auxiliar (exceden el saldo inicial)",
                      importe=psf.Cargo.sum(), hoja="Pagos sin factura",
                      detalle="El folio pagado no aparece como factura registrada.",
                      accion="Localizar y registrar la factura, o reclasificar como anticipo"))

    # 6. NC altas respecto a lo facturado
    nr = f[(f.notas_credito > 0)].copy()
    nr["pct"] = np.where(nr.facturado > 0, nr.notas_credito / nr.facturado.replace(0, np.nan), np.nan)
    nr = nr.sort_values("notas_credito", ascending=False)
    T["nc_por_proveedor"] = nr.reset_index()[["cta", "nombre", "notas_credito", "facturado", "pct", "fin"]]
    alta = nr[(nr.pct > 0.3) | nr.pct.isna()]
    if len(alta):
        H.append(dict(riesgo="Medio", titulo=f"{len(alta)} cuentas con notas de crédito mayores al 30% de lo facturado (o sin facturas)",
                      importe=alta.notas_credito.sum(), hoja="NC por proveedor",
                      detalle="; ".join(f"{i_n} {('sin facturas' if pd.isna(p) else f'{p:.0%}')}"
                                        for i_n, p in zip(alta.nombre.head(5), alta.pct.head(5))),
                      accion="Pedir soporte comercial de cada NC (devolución, bonificación o cancelación)"))

    # 7. Catálogo vs CFDI
    pro = m.dropna(subset=["pname"]).copy()       # solo movimientos con CFDI (RFC y nombre en la descripción)
    pro["ok"] = pd.Series([(_norm(a)[:12] in _norm(b)) or (_norm(b)[:12] in _norm(a))
                           for a, b in zip(pro["Cta. Desc."], pro.pname)], index=pro.index, dtype=bool)
    cat_rows = []
    for (c, n, rfc, pn), g in pro[~pro.ok].groupby(["Cta.", "Cta. Desc.", "rfc", "pname"]):
        if str(c).endswith("9999"):
            continue
        cat_rows.append(dict(cuenta=c, nombre_catalogo=str(n).strip(), rfc=rfc, nombre_cfdi=pn, facturas=len(g),
                             importe=g.amt.sum(), observacion="Nombre de la cuenta no coincide con el CFDI"))
    multi = pro.groupby("Cta.").rfc.nunique()
    for c in multi[multi > 1].index:
        g = pro[pro["Cta."] == c]
        principal = g.rfc.value_counts().index[0]
        for rfc, gg in g[g.rfc != principal].groupby("rfc"):
            obs = ("Cuenta genérica con varios RFC" if str(c).endswith("9999")
                   else f"RFC distinto al principal de la cuenta ({principal})")
            cat_rows.append(dict(cuenta=c, nombre_catalogo=str(g["Cta. Desc."].iloc[0]).strip(), rfc=rfc,
                                 nombre_cfdi=gg.pname.iloc[0], facturas=len(gg), importe=gg.amt.sum(), observacion=obs))
    T["catalogo"] = pd.DataFrame(cat_rows)
    if cat_rows:
        H.append(dict(riesgo="Medio", titulo=f"{len(cat_rows)} diferencias de nombre o RFC entre catálogo y CFDI",
                      importe=T["catalogo"].importe.sum(), hoja="Catálogo",
                      detalle="Revisar que cada cuenta corresponda a un solo contribuyente.",
                      accion="Corregir catálogo y abrir cuentas propias para proveedores recurrentes de 'Varios'"))

    # 8. Pagos antes de registrar la factura (saldo deudor temporal)
    tmp = m[m.Saldo < -0.005].groupby("Cta.").agg(movs=("Saldo", "size"), minimo=("Saldo", "min"), primera=("Fecha", "min")).reset_index()
    tmp["nombre"] = tmp["Cta."].map(f.nombre)
    tmp["saldo_final"] = tmp["Cta."].map(f.fin)
    T["saldo_deudor_temporal"] = tmp.sort_values("minimo")
    if len(tmp):
        H.append(dict(riesgo="Medio", titulo=f"{len(tmp)} cuentas tuvieron saldo contrario a su naturaleza durante el periodo (pago o cobro antes del documento)",
                      importe=None, hoja="Saldo deudor temporal",
                      detalle="Se registró el pago o cobro antes de provisionar el documento (factura).",
                      accion="Bloquear pagos sin factura provisionada, salvo anticipos autorizados"))

    # 9. Saldos iniciales sin liquidar
    f["ini_pend"] = np.minimum(np.where(f.ini > 0, np.maximum(0, f.ini - f.cargos), 0), f.fin.clip(lower=0))
    ip = f[f.ini_pend > 0.005].sort_values("ini_pend", ascending=False)
    T["saldos_iniciales"] = ip.reset_index()[["cta", "nombre", "ini", "cargos", "fin", "ini_pend"]]
    if len(ip):
        H.append(dict(riesgo="Medio", titulo=f"{len(ip)} cuentas con saldo inicial sin liquidar",
                      importe=ip.ini_pend.sum(), hoja="Saldos iniciales",
                      detalle="; ".join(f"{r.nombre} {_fmt(r.ini_pend)}" for _, r in ip.head(5).iterrows()),
                      accion="Confirmar si son pasivos reales o depurar"))

    # 10. Antigüedad PEPS
    rows = []
    for c, a in f[f.fin > 0.005].iterrows():
        bal = a.fin
        x = m[(m["Cta."] == c) & (m.Abono > 0)].sort_values("Fecha", ascending=False)
        b = {"0-30": 0.0, "31-60": 0.0, "61-90": 0.0, "91-180": 0.0, ">180 / saldo inicial": 0.0}
        for _, r in x.iterrows():
            if bal <= 0:
                break
            t = min(bal, r.Abono)
            bal -= t
            dd = (corte - r.Fecha).days
            k = "0-30" if dd <= 30 else "31-60" if dd <= 60 else "61-90" if dd <= 90 else "91-180" if dd <= 180 else ">180 / saldo inicial"
            b[k] += t
        b[">180 / saldo inicial"] += max(bal, 0)
        rows.append(dict(cta=c, nombre=a.nombre, saldo=a.fin, **b))
    ag = pd.DataFrame(rows)
    if len(ag):
        ag = ag.sort_values("saldo", ascending=False)
        ag["mas_90"] = ag["91-180"] + ag[">180 / saldo inicial"]
        T["antiguedad"] = ag
        v90 = ag.mas_90.sum()
        if v90 > 0:
            H.append(dict(riesgo="Medio", titulo="Saldo acreedor con más de 90 días",
                          importe=v90, hoja="Antigüedad",
                          detalle=f"{v90 / ag.saldo.sum():.0%} del saldo acreedor.",
                          accion="Revisar con tesorería: programar pago o aclarar disputas"))

    # 11. Concentración
    cc = f.sort_values("facturado", ascending=False).head(20).reset_index()[["cta", "nombre", "facturado", "fin"]]
    T["concentracion"] = cc

    # 12. Observaciones menores
    wk = m[m.Fecha.dt.dayofweek >= 5]
    rd = m[(m["Tipo póliza"] == "EA") & (m.Cargo >= 50000) & (m.Cargo % 1000 == 0)]
    fz = m[m.folio.fillna("").str.fullmatch(r'.*-0')]
    z = m[(m.Cargo.fillna(0) == 0) & (m.Abono.fillna(0) == 0)]
    T["fin_de_semana"] = wk
    T["pagos_redondos"] = rd
    T["sin_folio"] = fz
    T["importe_cero"] = z
    H.append(dict(riesgo="Bajo", titulo=f"Observaciones menores: {len(wk)} movimientos en fin de semana, {len(rd)} pagos redondos ≥ $50,000, "
                                       f"{len(fz)} documentos sin folio, {len(z)} renglones en cero",
                  importe=None, hoja="Otras observaciones", detalle="", accion="Revisión selectiva del soporte"))
    if any("rotulada" in k for k in res.meta):
        H.append(dict(riesgo="Bajo", titulo="El encabezado del reporte repite 'Cuenta inicial' (la segunda es la cuenta final)",
                      importe=None, hoja="—", detalle="", accion="Corregir la plantilla del reporte en el ERP"))

    orden = {"Alto": 0, "Medio": 1, "Bajo": 2}
    res.hallazgos = sorted(H, key=lambda h: orden[h["riesgo"]])
    res.tablas = T


def resumen_relacion(res: Resultado) -> pd.DataFrame:
    rel = res.relacion
    g = rel.groupby("categoria").agg(renglones=("row", "size"), cargos=("Cargo", "sum"), abonos=("Abono", "sum"))
    out = []
    for col, cat in CATEGORIAS:
        if cat in g.index:
            r = g.loc[cat]
            out.append(dict(color="Amarillo" if col == "Y" else "Azul", resultado=cat,
                            renglones=int(r.renglones), cargos=r.cargos, abonos=r.abonos))
    return pd.DataFrame(out)


# ═══════════════════════════════════════════════════════════════════════════════
# SALIDAS EXCEL
# ═══════════════════════════════════════════════════════════════════════════════
_HF = Font(name=F, bold=True, color="FFFFFF")
_HFILL = PatternFill("solid", fgColor="1F3864")
_BOR = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


def auxiliar_marcado(data: bytes, res: Resultado) -> bytes:
    """Devuelve el auxiliar original con cada movimiento en amarillo (relacionado) o azul (sin relación)."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb["Datos"] if "Datos" in wb.sheetnames else wb.worksheets[0]
    Y, B = PatternFill("solid", fgColor=AMARILLO), PatternFill("solid", fgColor=AZUL)
    obs = ws.max_column + 1
    OB = get_column_letter(obs)
    n = ws.max_row
    ws.cell(HEADER_ROW, obs, "Relación cargo-abono (auditoría)").font = Font(name=F, bold=True)
    ws.column_dimensions[OB].width = 90
    for r in res.relacion.itertuples():
        fill = Y if r.color == "Y" else B
        for j in range(1, obs + 1):
            ws.cell(r.row, j).fill = fill
        ws.cell(r.row, obs, r.motivo)
    for c, a in res.cuentas[res.cuentas.fin < -0.005].iterrows():
        if a.frow:
            for j in range(1, obs + 1):
                ws.cell(int(a.frow), j).fill = B
            ws.cell(int(a.frow), obs, "Saldo final contrario a la naturaleza de la cuenta")
    ws.cell(1, obs, "Leyenda").font = Font(name=F, bold=True)
    ws.cell(2, obs, "Amarillo = cargo relacionado con su abono (pago o NC ligado a su factura) / factura liquidada").fill = Y
    ws.cell(3, obs, "Azul = sin relación: cargo sin factura, factura sin pago o con pago parcial, pago en exceso, importe cero").fill = B
    ws.auto_filter.ref = f"A{HEADER_ROW}:{OB}{n}"

    rs = wb.create_sheet("Resumen relación")
    rs["A1"] = "Relación de cargos con abonos"
    rs["A1"].font = Font(name=F, bold=True, size=13)
    rs["A2"] = f"Conteos e importes con fórmula sobre la hoja Datos (columnas J, K y {OB})."
    rs["A2"].font = Font(name=F, italic=True, size=9)
    for j, h in enumerate(["Color", "Resultado", "Renglones", "Cargos", "Abonos"], 1):
        c = rs.cell(4, j, h)
        c.font, c.fill = _HF, _HFILL
    pats = {cat: pat for pat, cat in _CAT_PATS}
    i = 5
    for col, cat in CATEGORIAS:
        if cat == "Sin relación identificable":
            continue
        p = f"*{pats[cat]}*"
        rs.cell(i, 1, "Amarillo" if col == "Y" else "Azul").fill = Y if col == "Y" else B
        rs.cell(i, 2, cat)
        rs.cell(i, 3, f'=COUNTIF(Datos!${OB}${FIRST_DATA_ROW}:${OB}${n},"{p}")')
        rs.cell(i, 4, f'=SUMIF(Datos!${OB}${FIRST_DATA_ROW}:${OB}${n},"{p}",Datos!$J${FIRST_DATA_ROW}:$J${n})').number_format = MON
        rs.cell(i, 5, f'=SUMIF(Datos!${OB}${FIRST_DATA_ROW}:${OB}${n},"{p}",Datos!$K${FIRST_DATA_ROW}:$K${n})').number_format = MON
        i += 1
    rs.cell(i, 2, "Total").font = Font(name=F, bold=True)
    for j, L in [(3, "C"), (4, "D"), (5, "E")]:
        c = rs.cell(i, j, f"=SUM({L}5:{L}{i - 1})")
        c.font = Font(name=F, bold=True)
        c.number_format = MON if j > 3 else "0"
    rs.cell(i + 1, 2, "Movimientos en el auxiliar")
    rs.cell(i + 1, 3, f'=SUMPRODUCT(--(Datos!A{FIRST_DATA_ROW}:A{n}<>""))')
    rs.cell(i + 2, 2, "Diferencia (debe ser 0)")
    rs.cell(i + 2, 3, f"=C{i}-C{i + 1}")
    for j, w in enumerate([11, 44, 11, 17, 17], 1):
        rs.column_dimensions[get_column_letter(j)].width = w
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _tabla(ws, start, headers, rows, fmts, widths, totales=()):
    for j, h in enumerate(headers, 1):
        c = ws.cell(start, j, h)
        c.font, c.fill, c.border = _HF, _HFILL, _BOR
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows, start + 1):
        for j, v in enumerate(r, 1):
            if isinstance(v, float) and np.isnan(v):
                v = None
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            if isinstance(v, (np.integer,)):
                v = int(v)
            if isinstance(v, (np.floating,)):
                v = float(v)
            c = ws.cell(i, j, v)
            c.font, c.border = Font(name=F), _BOR
            if fmts[j - 1]:
                c.number_format = fmts[j - 1]
    last = start + len(rows)
    if totales and rows:
        ws.cell(last + 1, 1, "Total").font = Font(name=F, bold=True)
        for j in totales:
            L = get_column_letter(j)
            c = ws.cell(last + 1, j, f"=SUBTOTAL(9,{L}{start + 1}:{L}{last})")
            c.font, c.number_format = Font(name=F, bold=True), fmts[j - 1]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(start + 1, 1)
    if rows:
        ws.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{last}"
    return last


def papel_de_trabajo(res: Resultado) -> bytes:
    """Excel de auditoría: resumen, hallazgos con seguimiento y una hoja de detalle por hallazgo."""
    T, f, meta = res.tablas, res.cuentas, res.meta
    wb = openpyxl.Workbook()
    src = (f"Fuente: auxiliar de {meta.get('Empresa', '')} (RFC {meta.get('R. F. C.', '')}), "
           f"{_d(meta.get('Fecha inicial'))} a {_d(meta.get('Fecha final'))}.")

    def hoja(nombre, titulo):
        ws = wb.create_sheet(nombre[:31])
        ws["A1"], ws["A2"] = titulo, src
        ws["A1"].font, ws["A2"].font = Font(name=F, bold=True, size=14), Font(name=F, italic=True, size=9, color="595959")
        return ws

    # Saldos por cuenta (con recálculo)
    ws = hoja("Saldos por cuenta", "Saldos por cuenta y recálculo")
    rows = []
    for k, (c, a) in enumerate(f.iterrows(), 5):
        rows.append([c, a.nombre, a.ini, a.facturado, a.pagos, a.notas_credito, a.cargos, a.abonos, a.fin,
                     f'=C{k}+IF(N{k}="Deudora",G{k}-H{k},H{k}-G{k})', f"=ROUND(I{k}-J{k},2)",
                     f'=IF(I{k}<-0.005,"Contraria",IF(I{k}>0.005,"Normal","Cero"))', a.n, a.naturaleza])
    s1 = _tabla(ws, 4, ["Cuenta", "Nombre", "Saldo inicial", "Documentos (aumentos)", "Pagos / cobros", "Notas de crédito",
                        "Total cargos", "Total abonos", "Saldo final (auxiliar)", "Saldo final recalculado", "Diferencia",
                        "Saldo vs naturaleza", "Movimientos", "Naturaleza"],
                rows, [None, None, MON, MON, MON, MON, MON, MON, MON, MON, MON, None, "0", None],
                [19, 42, 15, 15, 15, 15, 16, 16, 16, 16, 12, 13, 8, 11], totales=[3, 4, 5, 6, 7, 8, 9, 10, 11, 13])

    detalle = [
        ("Saldos contrarios", "Cuentas con saldo contrario a su naturaleza", "saldos_deudores",
         ["Cuenta", "Nombre", "Naturaleza", "Saldo inicial", "Documentos", "Pagos / cobros", "Notas de crédito", "Saldo final"],
         [None, None, None, MON, MON, MON, MON, MON], [19, 42, 11, 15, 15, 15, 15, 15], [4, 5, 6, 7, 8]),
        ("NC sin factura", "Notas de crédito sin factura relacionada, por cuenta", "nc_sin_factura",
         ["Cuenta", "Proveedor", "Notas", "Importe"], [None, None, "0", MON], [19, 44, 8, 16], [3, 4]),
        ("NC por proveedor", "Notas de crédito contra lo facturado", "nc_por_proveedor",
         ["Cuenta", "Proveedor", "Notas de crédito", "Facturado", "NC / facturado", "Saldo final"],
         [None, None, MON, MON, PCT, MON], [19, 44, 16, 16, 12, 16], [3, 4, 6]),
        ("Pagos en exceso", "Pagos duplicados o mayores a la factura", "pagos_exceso",
         ["Renglón", "Cuenta", "Proveedor", "Póliza", "Fecha", "Importe", "Motivo"],
         [None, None, None, "0", DT, MON, None], [9, 19, 36, 8, 11, 14, 80], [6]),
        ("Egresos manuales", "Egresos capturados a mano sin CFDI", "egresos_manuales",
         ["Renglón", "Cuenta", "Proveedor", "Póliza", "Fecha", "Importe", "Descripción"],
         [None, None, None, "0", DT, MON, None], [9, 19, 36, 8, 11, 14, 60], [6]),
        ("Pagos sin factura", "Pagos a facturas que no están en el auxiliar", "pagos_sin_factura",
         ["Renglón", "Cuenta", "Proveedor", "Póliza", "Fecha", "Importe", "Folio / UUID pagado"],
         [None, None, None, "0", DT, MON, None], [9, 19, 36, 8, 11, 14, 40], [6]),
        ("Catálogo", "Diferencias de nombre o RFC entre catálogo y CFDI", "catalogo",
         ["Cuenta", "Nombre en catálogo", "RFC en CFDI", "Nombre en CFDI", "Facturas", "Importe", "Observación"],
         [None, None, None, None, "0", MON, None], [19, 36, 16, 36, 9, 15, 45], [5, 6]),
        ("Saldo deudor temporal", "Cuentas con saldo contrario a su naturaleza durante el periodo", "saldo_deudor_temporal",
         ["Cuenta", "Movimientos en saldo contrario", "Saldo más negativo", "Primera fecha", "Proveedor", "Saldo final"],
         [None, "0", MON, DT, None, MON], [19, 12, 16, 12, 44, 16], []),
        ("Saldos iniciales", "Saldos iniciales sin liquidar", "saldos_iniciales",
         ["Cuenta", "Proveedor", "Saldo inicial", "Cargos del periodo", "Saldo final", "Saldo inicial pendiente"],
         [None, None, MON, MON, MON, MON], [19, 44, 15, 16, 15, 18], [3, 4, 5, 6]),
        ("Antigüedad", "Antigüedad del saldo acreedor (PEPS)", "antiguedad",
         ["Cuenta", "Proveedor", "Saldo", "0-30", "31-60", "61-90", "91-180", ">180 / saldo inicial", "Más de 90 días"],
         [None, None, MON, MON, MON, MON, MON, MON, MON], [19, 44, 15, 14, 14, 14, 14, 16, 15], [3, 4, 5, 6, 7, 8, 9]),
        ("Concentración", "Top 20 cuentas por importe de documentos (facturas / aumentos)", "concentracion",
         ["Cuenta", "Proveedor", "Facturado", "Saldo final"], [None, None, MON, MON], [19, 44, 16, 16], []),
    ]
    totales_hoja = {}
    for nombre, titulo, key, hdr, fmts, widths, tots in detalle:
        t = T.get(key)
        ws = hoja(nombre, titulo)
        rows = [] if t is None or not len(t) else t.values.tolist()
        last = _tabla(ws, 4, hdr, rows, fmts, widths, totales=tots)
        totales_hoja[nombre] = last + 1
        if nombre == "Concentración" and rows:
            ws.cell(4, 5, "% del total").font, ws.cell(4, 5).fill = _HF, _HFILL
            ws.cell(4, 6, "% acumulado").font, ws.cell(4, 6).fill = _HF, _HFILL
            for k in range(5, last + 1):
                ws.cell(k, 5, f"=IF('Saldos por cuenta'!$D${s1 + 1}=0,0,C{k}/'Saldos por cuenta'!$D${s1 + 1})").number_format = PCT
                ws.cell(k, 6, f"=SUM($E$5:E{k})").number_format = PCT

    ws = hoja("Otras observaciones", "Observaciones menores")
    r0 = 4
    for etiqueta, key in [("Movimientos en fin de semana", "fin_de_semana"), ("Pagos ≥ $50,000 con importe redondo", "pagos_redondos"),
                          ("Documentos sin serie-folio", "sin_folio"), ("Renglones con importe cero", "importe_cero")]:
        t = T[key]
        ws.cell(r0, 1, f"{etiqueta} ({len(t)})").font = Font(name=F, bold=True)
        rows = t[["row", "Cta.", "Cta. Desc.", "Tipo póliza", "Num. Póliza", "Fecha", "Cargo", "Abono"]].values.tolist()
        last = _tabla(ws, r0 + 1, ["Renglón", "Cuenta", "Proveedor", "Tipo", "Póliza", "Fecha", "Cargo", "Abono"], rows,
                      [None, None, None, None, "0", DT, MON, MON], [9, 19, 40, 6, 8, 11, 15, 15], totales=[7, 8])
        r0 = last + 4
    ws.freeze_panes = None
    ws.auto_filter.ref = None

    # Resumen (primera hoja)
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = f"Depuración de cuentas — {meta.get('Empresa', '')}"
    ws["A1"].font = Font(name=F, bold=True, size=14)
    ws["A2"] = (f"RFC {meta.get('R. F. C.', '')} · {_d(meta.get('Fecha inicial'))} a {_d(meta.get('Fecha final'))} · "
                f"{len(f)} cuentas · {len(res.mov)} movimientos · pruebas al 100%")
    ws["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    S = "'Saldos por cuenta'!"
    L = lambda col: f"{S}{col}5:{col}{s1}"
    cuadre = [("Saldo inicial", f"=SUM({L('C')})"), ("(+) Documentos (aumentos)", f"=SUM({L('D')})"), ("(−) Pagos / cobros", f"=-SUM({L('E')})"),
              ("(−) Notas de crédito", f"=-SUM({L('F')})"), ("Saldo final calculado", "=SUM(B5:B8)"),
              ("Saldo final según auxiliar", f"=SUM({L('I')})"), ("Diferencia", "=ROUND(B9-B10,2)"),
              ("Saldos de su naturaleza", f'=SUMIF({L("L")},"Normal",{L("I")})'),
              ("Saldos contrarios a su naturaleza", f'=SUMIF({L("L")},"Contraria",{L("I")})'),
              ("Cuentas con diferencia en recálculo", f'=COUNTIF({L("K")},"<>0")')]
    ws["A4"] = "Cuadre del auxiliar"
    ws["A4"].font = Font(name=F, bold=True)
    for i, (a, b) in enumerate(cuadre, 5):
        ws.cell(i, 1, a).font = Font(name=F, bold=i in (9, 10, 11))
        c = ws.cell(i, 2, b)
        c.number_format = MON if i < 14 else "0"
    r = 17
    ws.cell(r, 1, "Hallazgos").font = Font(name=F, bold=True)
    hdr = ["#", "Hallazgo", "Riesgo", "Importe (MXN)", "Hoja de detalle", "Detalle", "Acción recomendada", "Responsable", "Fecha compromiso", "Estatus"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r + 1, j, h)
        c.font, c.fill = _HF, _HFILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    colores = {"Alto": "F8CBAD", "Medio": "FFE699", "Bajo": "E2EFDA"}
    for k, h in enumerate(res.hallazgos, 1):
        i = r + 1 + k
        vals = [k, h["titulo"], h["riesgo"], h["importe"], h["hoja"], h["detalle"], h["accion"], None, None, "Pendiente"]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, float(v) if isinstance(v, (np.floating,)) else v)
            c.font, c.border = Font(name=F), _BOR
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(i, 4).number_format = MON
        ws.cell(i, 3).fill = PatternFill("solid", fgColor=colores[h["riesgo"]])
        for j in (8, 9, 10):
            ws.cell(i, j).fill = PatternFill("solid", fgColor="FFFF00")
        ws.cell(i, 9).number_format = DT
        if h["hoja"] in wb.sheetnames:
            ws.cell(i, 5).hyperlink = f"#'{h['hoja']}'!A1"
            ws.cell(i, 5).font = Font(name=F, color="0563C1", underline="single")
    n = len(res.hallazgos)
    dv = DataValidation(type="list", formula1='"Pendiente,En proceso,Resuelto,No procede"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J{r + 2}:J{r + 1 + n}")
    ws.cell(r + n + 3, 1, "Celdas en amarillo: llenar Responsable, Fecha compromiso y Estatus (lista desplegable).").font = \
        Font(name=F, italic=True, size=9, color="595959")
    for j, w in enumerate([30, 55, 9, 16, 20, 55, 45, 16, 13, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _d(v) -> str:
    return v.strftime("%d-%m-%Y") if hasattr(v, "strftime") else str(v or "")


# ═══════════════════════════════════════════════════════════════════════════════
# REPORTE POR PROVEEDOR: lo que se cancela vs lo pendiente de conciliar
# ═══════════════════════════════════════════════════════════════════════════════
def _estado_renglon(r) -> str:
    pend = abs(r.cargo_pendiente) + abs(r.abono_pendiente)
    conc = abs(r.cargo_conciliado) + abs(r.cargo_saldo_inicial) + abs(r.abono_conciliado) + abs(r.abono_saldo_inicial)
    if pend <= 0.005 and conc <= 0.005:
        return "Sin importe"
    if pend <= 0.005 or (r.color == "Y" and pend <= 0.011):   # centavo de redondeo
        return "Conciliado"
    return "Parcial" if conc > 0.005 else "Pendiente"


def detalle_conciliacion(res: Resultado) -> pd.DataFrame:
    rel = res.relacion.merge(res.mov[["row", "folio", "Desc. Detalle"]], on="row", how="left")
    for c in ("cargo_conciliado", "cargo_saldo_inicial", "abono_conciliado", "abono_saldo_inicial"):
        rel[c] = rel[c].round(2)
    # lo pendiente se obtiene por diferencia para que cada renglón cuadre al centavo con el auxiliar
    rel["cargo_pendiente"] = (rel.Cargo.fillna(0) - rel.cargo_conciliado - rel.cargo_saldo_inicial).round(2)
    rel["abono_pendiente"] = (rel.Abono.fillna(0) - rel.abono_conciliado - rel.abono_saldo_inicial).round(2)
    rel["estado"] = rel.apply(_estado_renglon, axis=1)
    # Renglones en cero: un solo grupo por cuenta
    cero = rel.estado == "Sin importe"
    rel.loc[cero, "grupo_key"] = "Z|" + rel.loc[cero, "Cta."]
    rel["pend_abs"] = rel.cargo_pendiente.abs() + rel.abono_pendiente.abs()
    g = rel.groupby("grupo_key")
    rel["grupo_fecha"] = g.Fecha.transform("min")
    rel["grupo_pend"] = g.pend_abs.transform("sum")
    rel["grupo_n"] = g.row.transform("size")
    # documento = lado que aumenta el saldo (abono en acreedoras, cargo en deudoras)
    rel["es_abono"] = np.where(rel.naturaleza == "Deudora", rel.Cargo.fillna(0) != 0, rel.Abono.fillna(0) != 0).astype(int)

    def _tipo(k, gg):
        if k.startswith("SI|"):
            return "Pagos aplicados al saldo inicial"
        if k.startswith("Z|"):
            return "Renglones con importe cero (depurar)"
        fol = [str(x) for x in gg.loc[gg.es_abono == 1, "folio"].dropna().unique()]
        if fol:
            return "Factura " + ", ".join(fol[:4]) + (" …" if len(fol) > 4 else "")
        if gg.isnc.any() and len(gg) > 1:
            return "Nota de crédito aplicada en el pago"
        docs = gg[gg.es_abono == 1]
        if len(docs) and len(gg) > 1:
            return f"Documento del renglón {int(docs.row.iloc[0])}"
        return gg.categoria.iloc[0]
    tipos = {k: _tipo(k, gg) for k, gg in rel.merge(res.mov[["row", "isnc"]], on="row").groupby("grupo_key")}
    rel["grupo_tipo"] = rel.grupo_key.map(tipos)

    def _gestado(r):
        if r.grupo_key.startswith("Z|"):
            return "Sin importe"
        if r.grupo_key.startswith("SI|"):
            return "Aplicado a saldo inicial"
        return "Se cancela" if r.grupo_pend <= 0.005 else "Pendiente"
    rel["grupo_estado"] = rel.apply(_gestado, axis=1)
    # orden: por cuenta, de lo más antiguo (arriba) a lo más reciente (abajo); los movimientos que se
    # cancelan entre sí quedan juntos: el grupo se ubica por su fecha más antigua y dentro va por fecha
    rel["orden_tipo"] = rel.grupo_key.str[0].map({"S": 0, "R": 1, "Z": 2})
    rel = rel.sort_values(["Cta.", "orden_tipo", "grupo_fecha", "grupo_key", "Fecha", "es_abono", "row"],
                          ascending=[True, True, True, True, True, False, True]).reset_index(drop=True)
    rel["grupo"] = rel.groupby("Cta.").grupo_key.transform(lambda s: pd.factorize(s)[0] + 1)
    return rel.drop(columns=["pend_abs", "orden_tipo"])


def resumen_por_proveedor(res: Resultado) -> pd.DataFrame:
    det = detalle_conciliacion(res)
    f = res.cuentas
    det = det.assign(aplicado_si=det.cargo_saldo_inicial + det.abono_saldo_inicial)
    g = det.groupby("Cta.").agg(
        cargos_saldo_inicial=("aplicado_si", "sum"), abonos_conciliados=("abono_conciliado", "sum"),
        cargos_conciliados=("cargo_conciliado", "sum"), abonos_pendientes=("abono_pendiente", "sum"),
        cargos_pendientes=("cargo_pendiente", "sum"),
        movs_conciliados=("estado", lambda s: int((s == "Conciliado").sum())),
        movs_pendientes=("estado", lambda s: int(s.isin(["Pendiente", "Parcial"]).sum())),
        movs_cero=("estado", lambda s: int((s == "Sin importe").sum())),
    ).reindex(f.index).fillna(0)
    out = pd.DataFrame({"nombre": f.nombre, "naturaleza": f.naturaleza, "saldo_inicial": f.ini}).join(g)
    signo = np.where(out.naturaleza == "Deudora", -1, 1)
    out["saldo_inicial_pendiente"] = (out.saldo_inicial - out.cargos_saldo_inicial).round(2)
    out["saldo_pendiente"] = (out.saldo_inicial_pendiente + signo * (out.abonos_pendientes - out.cargos_pendientes)).round(2)
    out["saldo_final"] = f.fin
    out["diferencia"] = (out.saldo_pendiente - out.saldo_final).round(2)
    out["estatus"] = np.where((out.saldo_inicial_pendiente.abs() < 0.01) & (out.movs_pendientes == 0), "Conciliada",
                              "Pendiente de conciliar")
    return out


def reporte_por_proveedor(res: Resultado) -> bytes:
    """Excel: resumen por cuenta (saldo inicial, cargos y abonos que se cancelan vs pendientes),
    bloques por cuenta de pendientes y de conciliados, y el detalle de todos los movimientos."""
    det = detalle_conciliacion(res)
    f, meta = res.cuentas, res.meta
    wb = openpyxl.Workbook()
    Y, B = PatternFill("solid", fgColor=AMARILLO), PatternFill("solid", fgColor=AZUL)
    GR = PatternFill("solid", fgColor="D9E1F2")
    bold = Font(name=F, bold=True)
    norm = Font(name=F)
    src = (f"{meta.get('Empresa', '')} · RFC {meta.get('R. F. C.', '')} · "
           f"{_d(meta.get('Fecha inicial'))} a {_d(meta.get('Fecha final'))}")

    def titulo(ws, t, sub):
        ws["A1"], ws["A2"] = t, sub
        ws["A1"].font = Font(name=F, bold=True, size=14)
        ws["A2"].font = Font(name=F, italic=True, size=9, color="595959")

    def encabezado(ws, row, hdr):
        for j, h in enumerate(hdr, 1):
            c = ws.cell(row, j, h)
            c.font, c.fill, c.border = _HF, _HFILL, _BOR
            c.alignment = Alignment(wrap_text=True, vertical="center")

    def val(v):
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            return None if np.isnan(v) else float(v)
        if isinstance(v, float) and np.isnan(v):
            return None
        return v

    # ── Detalle de movimientos (base de las fórmulas del resumen) ─────────────
    wd = wb.active
    wd.title = "Detalle movimientos"
    titulo(wd, "Detalle de movimientos: importe conciliado y pendiente por renglón", src +
           " · Renglón = fila del auxiliar original. Ordenado por cuenta, del más antiguo al más reciente.")
    hd = ["Renglón", "Cuenta", "Nombre de la cuenta", "Tipo", "Póliza", "Fecha", "Folio", "Descripción", "Cargo", "Abono",
          "Cargo conciliado", "Cargo aplicado a saldo inicial", "Cargo pendiente", "Abono conciliado",
          "Abono pendiente", "Estado", "Relación / motivo", "Grupo", "Tipo de grupo", "Estado del grupo",
          "Abono aplicado a saldo inicial", "Naturaleza"]
    encabezado(wd, 4, hd)
    fm = [None, None, None, None, "0", DT, None, None, MON, MON, MON, MON, MON, MON, MON, None, None, "0", None, None, MON, None]
    keys = ["row", "Cta.", "Cta. Desc.", "Tipo póliza", "Num. Póliza", "Fecha", "folio", "Desc. Detalle", "Cargo", "Abono",
            "cargo_conciliado", "cargo_saldo_inicial", "cargo_pendiente", "abono_conciliado", "abono_pendiente",
            "estado", "motivo", "grupo", "grupo_tipo", "grupo_estado", "abono_saldo_inicial", "naturaleza"]
    for i, rd in enumerate(det[keys].to_dict("records"), 5):
        vals = [rd[k] for k in keys]
        for j, v in enumerate(vals, 1):
            c = wd.cell(i, j, val(v))
            c.font = norm
            if fm[j - 1]:
                c.number_format = fm[j - 1]
        fill = Y if rd["estado"] == "Conciliado" else B
        for j in range(1, len(hd) + 1):
            wd.cell(i, j).fill = fill
    nd = 4 + len(det)
    for j, w in enumerate([9, 18, 34, 6, 8, 11, 16, 50, 14, 14, 15, 15, 14, 14, 14, 11, 70, 7, 30, 18, 15, 11], 1):
        wd.column_dimensions[get_column_letter(j)].width = w
    wd.freeze_panes = "D5"
    wd.auto_filter.ref = f"A4:{get_column_letter(len(hd))}{nd}"
    D = "'Detalle movimientos'!"
    rng = lambda col: f"{D}${col}$5:${col}${nd}"

    # ── Resumen por cuenta (con fórmulas sobre el detalle) ────────────────────
    ws = wb.create_sheet("Resumen por cuenta", 0)
    titulo(ws, "Resumen por cuenta: lo que se cancela y lo pendiente de conciliar", src)
    hr = ["Cuenta", "Nombre de la cuenta", "Saldo inicial", "Aplicado al saldo inicial", "Saldo inicial pendiente",
          "Abonos conciliados", "Cargos conciliados", "Diferencia conciliada (debe ser 0)",
          "Abonos pendientes", "Cargos pendientes", "Saldo pendiente de conciliar", "Saldo final auxiliar",
          "Diferencia vs auxiliar", "Movs. conciliados", "Movs. pendientes", "Estatus", "Naturaleza"]
    ws["A3"] = ("Saldo pendiente = saldo inicial pendiente + abonos pendientes − cargos pendientes (cuentas acreedoras) "
                "o + cargos pendientes − abonos pendientes (cuentas deudoras). "
                "Lo conciliado se cancela entre sí (abonos conciliados = cargos conciliados).")
    ws["A3"].font = Font(name=F, italic=True, size=9, color="595959")
    encabezado(ws, 5, hr)
    rp = resumen_por_proveedor(res)
    fr = [None, None, MON, MON, MON, MON, MON, MON, MON, MON, MON, MON, MON, "0", "0", None, None]
    for i, (c, a) in enumerate(rp.iterrows(), 6):
        A = f"$A{i}"
        vals = [c, a.nombre, a.saldo_inicial,
                f"=SUMIFS({rng('L')},{rng('B')},{A})+SUMIFS({rng('U')},{rng('B')},{A})", f"=C{i}-D{i}",
                f"=SUMIFS({rng('N')},{rng('B')},{A})", f"=SUMIFS({rng('K')},{rng('B')},{A})", f"=ROUND(F{i}-G{i},2)",
                f"=SUMIFS({rng('O')},{rng('B')},{A})", f"=SUMIFS({rng('M')},{rng('B')},{A})",
                f'=ROUND(E{i}+IF(Q{i}="Deudora",J{i}-I{i},I{i}-J{i}),2)', a.saldo_final, f"=ROUND(K{i}-L{i},2)",
                f'=COUNTIFS({rng("B")},{A},{rng("P")},"Conciliado")',
                f'=COUNTIFS({rng("B")},{A},{rng("P")},"Pendiente")+COUNTIFS({rng("B")},{A},{rng("P")},"Parcial")',
                f'=IF(AND(ABS(E{i})<0.01,O{i}=0),"Conciliada","Pendiente de conciliar")', a.naturaleza]
        for j, v in enumerate(vals, 1):
            cc = ws.cell(i, j, val(v))
            cc.font, cc.border = norm, _BOR
            if fr[j - 1]:
                cc.number_format = fr[j - 1]
        ws.cell(i, 16).fill = Y if a.estatus == "Conciliada" else B
    nr = 5 + len(rp)
    t = nr + 1
    ws.cell(t, 1, "Total").font = bold
    for j in range(3, 16):    # totales de C a O
        L = get_column_letter(j)
        cc = ws.cell(t, j, f"=SUBTOTAL(9,{L}6:{L}{nr})")
        cc.font, cc.number_format = bold, fr[j - 1]
    ws.cell(t + 2, 1, "Cuentas conciliadas").font = bold
    ws.cell(t + 2, 3, f'=COUNTIF(P6:P{nr},"Conciliada")')
    ws.cell(t + 3, 1, "Cuentas pendientes de conciliar").font = bold
    ws.cell(t + 3, 3, f'=COUNTIF(P6:P{nr},"Pendiente de conciliar")')
    ws.cell(t + 4, 1, "Cuentas con diferencia vs auxiliar").font = bold
    ws.cell(t + 4, 3, f'=SUMPRODUCT(--(ABS(M6:M{nr})>0.015))')
    for j, w in enumerate([18, 40, 15, 16, 15, 16, 16, 14, 16, 16, 17, 16, 12, 11, 11, 20, 11], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A5:Q{nr}"

    # ── Bloques por cuenta ────────────────────────────────────────────────────
    fecha_ini = pd.Timestamp(meta.get("Fecha inicial")) if meta.get("Fecha inicial") else det.Fecha.min()

    # Saldo inicial como renglón propio: va a Conciliados si sus pagos lo cancelan, si no a Pendientes
    det = det.copy()
    det["orden"] = range(len(det))
    det["bloque"] = 0
    pagos_si = det[det.grupo_key.str.startswith("SI|")]
    si_rows = []
    for c, a in rp.iterrows():
        if abs(a.saldo_inicial) < 0.005:
            continue
        pos = (a.naturaleza != "Deudora") == (a.saldo_inicial > 0)     # True: el saldo inicial va del lado del abono
        pend = float(a.saldo_inicial_pendiente)
        npag = int((pagos_si["Cta."] == c).sum())
        cancela = abs(pend) < 0.005
        motivo = (f"Saldo inicial liquidado con {npag} movimiento(s) aplicados" if cancela else
                  f"Saldo inicial pendiente {_fmt(pend)}" + (f" (aplicado {_fmt(a.cargos_saldo_inicial)} en {npag} movimiento(s))" if npag else " (sin movimientos aplicados en el periodo)"))
        si_rows.append({"Cta.": c, "row": None, "Tipo póliza": "SI", "Num. Póliza": None, "Fecha": fecha_ini, "folio": None,
                        "Desc. Detalle": f"Saldo inicial al {_d(fecha_ini)}",
                        "Cargo": 0.0 if pos else abs(a.saldo_inicial), "Abono": abs(a.saldo_inicial) if pos else 0.0,
                        "cargo_pendiente": 0.0 if pos else abs(pend) * np.sign(pend * a.saldo_inicial),
                        "abono_pendiente": abs(pend) * np.sign(pend * a.saldo_inicial) if pos else 0.0,
                        "estado": "Conciliado" if cancela else "Pendiente", "motivo": motivo,
                        "grupo_key": f"SI|{c}", "grupo_tipo": "Saldo inicial",
                        "grupo_estado": "Se cancela" if cancela else "Pendiente", "grupo": 0, "orden": -1, "bloque": 0})
    si_df = pd.DataFrame(si_rows)
    if len(si_df):
        est_si = si_df.set_index("grupo_key").grupo_estado
        det.loc[det.grupo_key.str.startswith("SI|"), "grupo_estado"] = det.grupo_key.map(est_si)
        det.loc[det.grupo_key.str.startswith("SI|"), "bloque"] = 0
        det.loc[~det.grupo_key.str.startswith("SI|"), "bloque"] = 1
        det.loc[det.grupo_key.str.startswith("SI|"), "grupo_tipo"] = "Saldo inicial"
        todo = pd.concat([det, si_df.dropna(axis=1, how="all")], ignore_index=True)
    else:
        todo = det
    todo = todo.sort_values(["Cta.", "bloque", "orden"]).reset_index(drop=True)

    def hoja_bloques(nombre, tit, sub, sel, hdr, cols, fmts, widths, pendientes):
        wsb = wb.create_sheet(nombre)
        titulo(wsb, tit, src + " · " + sub)
        encabezado(wsb, 4, hdr)
        nc = len(hdr)
        row = 5
        for c, gc in sel.groupby("Cta.", sort=True):
            a = rp.loc[c]
            wsb.cell(row, 1, f"{c} · {a.nombre}").font = Font(name=F, bold=True, size=11)
            for j in range(1, nc + 1):
                wsb.cell(row, j).fill = GR
            prov_row = row
            row += 1
            for rd in gc[cols].to_dict("records"):
                for j, colname in enumerate(cols, 1):
                    cc = wsb.cell(row, j, val(rd[colname]))
                    cc.font = Font(name=F, bold=True) if rd["Tipo póliza"] == "SI" else norm
                    if fmts[j - 1]:
                        cc.number_format = fmts[j - 1]
                fill = Y if (not pendientes or rd.get("estado") == "Conciliado") else B
                for j in range(1, nc + 1):
                    wsb.cell(row, j).fill = fill
                wsb.row_dimensions[row].outline_level = 1
                row += 1
            if pendientes:
                deudora = a.naturaleza == "Deudora"
                wsb.cell(row, 6, "Saldo pendiente de conciliar (" + ("cargos pendientes − abonos pendientes)" if deudora
                                  else "abonos pendientes − cargos pendientes)")).font = bold
                cg = wsb.cell(row, 9, f"=SUM(I{prov_row + 1}:I{row - 1})")
                ab = wsb.cell(row, 10, (f"=I{row}-SUM(J{prov_row + 1}:J{row - 1})" if deudora
                                        else f"=SUM(J{prov_row + 1}:J{row - 1})-I{row}"))
                for cc in (cg, ab):
                    cc.number_format, cc.font = MON, bold
                wsb.cell(row, 11, "Saldo final auxiliar:").font = bold
                cc = wsb.cell(row, 12, float(a.saldo_final))
                cc.number_format, cc.font = MON, bold
            else:
                wsb.cell(row, 6, "Total conciliado de la cuenta (abonos − cargos = 0)").font = bold
                for j, L in [(7, "G"), (8, "H")]:
                    cc = wsb.cell(row, j, f"=SUM({L}{prov_row + 1}:{L}{row - 1})")
                    cc.number_format, cc.font = MON, bold
                cc = wsb.cell(row, 9, f'=IF(ROUND(H{row}-G{row},2)=0,"Se cancela","Diferencia "&TEXT(H{row}-G{row},"#,##0.00"))')
                cc.font = bold
            for j in range(1, nc + 1):
                wsb.cell(row, j).border = Border(top=Side(style="medium", color="1F3864"))
            row += 3
        for j, w in enumerate(widths, 1):
            wsb.column_dimensions[get_column_letter(j)].width = w
        wsb.freeze_panes = "A5"
        wsb.sheet_properties.outlinePr.summaryBelow = False
        return wsb

    # Pendientes: movimientos con algo por conciliar, consecutivos (factura junto a sus pagos parciales),
    # y el saldo inicial cuando no se cancela, junto con los pagos que se le aplicaron
    sel_p = todo[todo.grupo_estado == "Pendiente"]
    hdr_p = ["Renglón", "Tipo", "Póliza", "Fecha", "Folio", "Descripción", "Cargo", "Abono",
             "Cargo pendiente", "Abono pendiente", "Estado", "Motivo"]
    cols_p = ["row", "Tipo póliza", "Num. Póliza", "Fecha", "folio", "Desc. Detalle", "Cargo", "Abono",
              "cargo_pendiente", "abono_pendiente", "estado", "motivo"]
    fm_p = [None, None, "0", DT, None, None, MON, MON, MON, MON, None, None]
    hoja_bloques("Pendientes por cuenta", "Movimientos pendientes de conciliar, por cuenta",
                 "Del más antiguo al más reciente. Azul = parte pendiente; amarillo = parte ya conciliada que acompaña al movimiento pendiente.",
                 sel_p, hdr_p, cols_p, fm_p, [9, 6, 8, 11, 16, 55, 14, 14, 15, 15, 11, 60], True)

    # Conciliados: grupos que se cancelan y el saldo inicial cuando sus pagos lo cancelan
    sel_c = todo[todo.grupo_estado == "Se cancela"]
    hdr_c = ["Renglón", "Tipo", "Póliza", "Fecha", "Folio", "Descripción", "Cargo", "Abono", "Relación"]
    cols_c = ["row", "Tipo póliza", "Num. Póliza", "Fecha", "folio", "Desc. Detalle", "Cargo", "Abono", "motivo"]
    fm_c = [None, None, "0", DT, None, None, MON, MON, None]
    hoja_bloques("Conciliados por cuenta", "Movimientos que se cancelan entre sí, por cuenta",
                 "Del más antiguo al más reciente; cada documento junto a los movimientos que lo cancelan.", sel_c, hdr_c, cols_c, fm_c,
                 [9, 6, 8, 11, 16, 55, 14, 14, 70], False)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
