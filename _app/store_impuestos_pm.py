"""
_app/store_impuestos_pm.py — Guardado de los cálculos mensuales de Impuestos PM (S.C.)

Un documento por contribuyente y ejercicio: {"rfc", "razon_social", "config", "meses": {"1": {...}}}.
  · En línea: tabla `impuestos_pm` de Supabase (si [supabase] url/key están en Secrets).
  · Local: archivo JSON en data_local/impuestos_pm/<RFC>_<ejercicio>.json.

SQL para crear la tabla en Supabase (SQL Editor):
    create table if not exists impuestos_pm (
        rfc        text    not null,
        ejercicio  int     not null,
        data       jsonb   not null,
        updated_at timestamptz default now(),
        primary key (rfc, ejercicio)
    );
"""
from __future__ import annotations

import datetime as _dt
import json
import os

SQL_TABLA = """create table if not exists impuestos_pm (
    rfc        text    not null,
    ejercicio  int     not null,
    data       jsonb   not null,
    updated_at timestamptz default now(),
    primary key (rfc, ejercicio)
);"""

_DIR_LOCAL = os.path.join(os.path.dirname(__file__), "..", "data_local", "impuestos_pm")


def _json_default(o):
    if isinstance(o, (_dt.date, _dt.datetime)):
        return o.isoformat()
    raise TypeError(f"No serializable: {type(o)}")


def limpiar(doc: dict) -> dict:
    """Convierte fechas a texto para guardarlas como JSON."""
    return json.loads(json.dumps(doc, default=_json_default))


class Almacen:
    """Elige Supabase si está configurado y la tabla existe; si no, archivo local."""

    def __init__(self, secrets=None):
        self.modo = "local"
        self.error = ""
        self._sb = None
        try:
            url = secrets["supabase"]["url"] if secrets else ""
            key = secrets["supabase"]["key"] if secrets else ""
        except Exception:
            url = key = ""
        if url and key and url.startswith("https://"):
            try:
                from supabase import create_client
                self._sb = create_client(url, key)
                self._sb.table("impuestos_pm").select("rfc").limit(1).execute()
                self.modo = "supabase"
            except Exception as e:  # tabla inexistente, red, llave…
                self._sb = None
                self.error = str(e)[:300]

    # ── local ────────────────────────────────────────────────────────────────
    @staticmethod
    def _ruta(rfc: str, ejercicio: int) -> str:
        os.makedirs(_DIR_LOCAL, exist_ok=True)
        return os.path.abspath(os.path.join(_DIR_LOCAL, f"{rfc.upper()}_{int(ejercicio)}.json"))

    # ── API ──────────────────────────────────────────────────────────────────
    def cargar(self, rfc: str, ejercicio: int) -> dict | None:
        rfc = rfc.upper().strip()
        if self.modo == "supabase":
            rows = (self._sb.table("impuestos_pm").select("data")
                    .eq("rfc", rfc).eq("ejercicio", int(ejercicio)).execute().data or [])
            return rows[0]["data"] if rows else None
        ruta = self._ruta(rfc, ejercicio)
        if os.path.exists(ruta):
            with open(ruta, encoding="utf-8") as fh:
                return json.load(fh)
        return None

    def guardar(self, doc: dict) -> None:
        doc = limpiar(doc)
        rfc, ej = doc["rfc"].upper().strip(), int(doc["config"]["ejercicio"])
        if self.modo == "supabase":
            self._sb.table("impuestos_pm").upsert({
                "rfc": rfc, "ejercicio": ej, "data": doc,
                "updated_at": _dt.datetime.utcnow().isoformat(),
            }).execute()
            return
        with open(self._ruta(rfc, ej), "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=1)

    def listar(self) -> list[tuple[str, int]]:
        if self.modo == "supabase":
            rows = self._sb.table("impuestos_pm").select("rfc,ejercicio").execute().data or []
            return sorted((r["rfc"], int(r["ejercicio"])) for r in rows)
        if not os.path.isdir(_DIR_LOCAL):
            return []
        out = []
        for f in os.listdir(_DIR_LOCAL):
            if f.endswith(".json") and "_" in f:
                rfc, ej = f[:-5].rsplit("_", 1)
                if ej.isdigit():
                    out.append((rfc, int(ej)))
        return sorted(out)
