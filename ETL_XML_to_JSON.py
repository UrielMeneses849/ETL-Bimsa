# ETL_XML_to_JSON.py
import re
import os
import json
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Union, Any, Dict, List

_AMP_NEEDS_ESCAPE = re.compile(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)")


def _escape_bare_ampersands(xml: str) -> str:
    return _AMP_NEEDS_ESCAPE.sub("&amp;", xml)


def _accent_score(s: str) -> int:
    good = "áéíóúÁÉÍÓÚñÑüÜ"
    return sum(ch in s for ch in good)


def _looks_mojibake(s: str) -> bool:
    return any(h in s for h in ("Ã", "Â", "µ", "¥", "�", "à", "¨", "¢", "¤"))


# def _repair_mojibake(s: str) -> str:
#    if not isinstance(s, str) or not s:
#        return s

    # UTF-8 leído como latin1/cp1252: MÃ©xico
    if ("Ã" in s) or ("Â" in s):
        for enc in ("latin1", "cp1252"):
            try:
                cand = s.encode(enc).decode("utf-8")
                if cand and cand != s:
                    s = cand
                    break
            except Exception:
                pass

    # CP850/CP437 vs CP1252: Ni¥os / µREA / UrbanizaciàN
    if any(ch in s for ch in ("µ", "¥", "à", "¨", "¢", "¤", "�")):
        best = s
        best_score = _accent_score(s)
        for target in ("cp850", "cp437"):
            try:
                cand = s.encode("cp1252").decode(target)
                score = _accent_score(cand)
                if score >= best_score and cand:
                    best, best_score = cand, score
            except Exception:
                pass
        s = best

    return s


def _elem_to_dict(elem: ET.Element):
    children = list(elem)
    if not children:
        return (elem.text or "").strip()

    grouped = {}
    for ch in children:
        tag = ch.tag.split("}", 1)[-1]
        grouped.setdefault(tag, []).append(_elem_to_dict(ch))

    out = {}
    for k, v in grouped.items():
        out[k] = v[0] if len(v) == 1 else v
    return out


def _force_list_payload(payload: Any) -> List[Dict[str, Any]]:
    """
    Garantiza que el JSON final SIEMPRE sea una LISTA de registros (list[dict]).
    """
    if payload is None:
        return []

    # Si ya es lista:
    if isinstance(payload, list):
        # si hay strings sueltos, envuélvelos
        out = []
        for x in payload:
            out.append(x if isinstance(x, dict) else {"value": x})
        return out

    # Si es dict:
    if isinstance(payload, dict):
        # Caso típico: { "row": [ {...}, {...} ] } o { "registro": [...] }
        list_candidates = []
        for v in payload.values():
            if isinstance(v, list) and v and all(isinstance(i, dict) for i in v):
                list_candidates.append(v)

        if list_candidates:
            # toma la lista más larga (más probable que sean los registros)
            return max(list_candidates, key=len)

        # Caso: { "datos": { "row": [...] } } ya lo habrás desempacado antes, pero por si acaso:
        for v in payload.values():
            if isinstance(v, dict):
                inner = _force_list_payload(v)
                if inner:
                    return inner

        # Si es un registro único (dict plano), lo volvemos lista
        return [payload]

    # Cualquier otro tipo:
    return [{"value": payload}]


def xml_a_json(
    xml_input: Union[str, bytes],
    tipo_reporte: str,
    carpeta_salida: str = ".",
    silent: bool = True,
) -> str:

    # 1) Normaliza entrada
    if isinstance(xml_input, bytes):
        for enc in ("utf-8-sig", "utf-8", "cp1252", "cp850", "latin1"):
            try:
                xml_string = xml_input.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            xml_string = xml_input.decode("utf-8", errors="replace")
    else:
        xml_string = xml_input or ""

    # 2) Limpieza segura
    xml_limpio = xml_string.strip().lstrip("\ufeff")

    # 3) Reparar mojibake si ya viene mal
    #if _looks_mojibake(xml_limpio):
     #   xml_limpio = _repair_mojibake(xml_limpio)

    # 4) Escapar ampersands sueltos (sin romper entidades)
    xml_limpio = _escape_bare_ampersands(xml_limpio)

    # 5) Parse
    try:
        root = ET.fromstring(xml_limpio)
    except ET.ParseError as e:
        raise ValueError(f"XML inválido tras limpieza: {e}")

    payload = _elem_to_dict(root)

    # Si viene envuelto:
    if isinstance(payload, dict) and "datos" in payload:
        payload = payload["datos"]

    payload_list = _force_list_payload(payload)

    # 6) Guardar JSON UTF-8 real
    os.makedirs(carpeta_salida, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(carpeta_salida, f"{tipo_reporte}_{ts}.json")

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload_list, f, ensure_ascii=False, indent=2)

    if not silent:
        print(f"[XML2JSON] OK -> {out_path}")

    return out_path