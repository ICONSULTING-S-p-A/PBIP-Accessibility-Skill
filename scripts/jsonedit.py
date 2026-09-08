"""
Editor JSON "chirurgico": applica modifiche puntuali a un testo JSON preservando
indentazione, ordine delle chiavi, fine riga e tutto il JSON non coinvolto.

Equivalente Python di jsonc-parser (modify/applyEdits) usato dall'MCP originale.
Serve perche' i file PBIR vanno modificati in-place senza riformattarli: una
riserializzazione completa cambierebbe dettagli di formato che Power BI Desktop
riscrive poi in diff enormi e illeggibili.
"""
from __future__ import annotations

import json
from typing import Any, List, Optional, Sequence, Union

JsonPath = Sequence[Union[str, int]]
WS = " \t\r\n"


class JsonEditError(Exception):
    pass


class Prop:
    __slots__ = ("key", "kstart", "kend", "value", "pstart", "pend")

    def __init__(self, key, kstart, kend, value, pstart, pend):
        self.key = key
        self.kstart = kstart
        self.kend = kend
        self.value = value
        self.pstart = pstart
        self.pend = pend


class Node:
    __slots__ = ("type", "offset", "end", "value", "props", "items")

    def __init__(self, type_, offset, end, value=None, props=None, items=None):
        self.type = type_
        self.offset = offset
        self.end = end
        self.value = value
        self.props: List[Prop] = props or []
        self.items: List[Node] = items or []


class _Parser:
    def __init__(self, text: str):
        self.t = text
        self.i = 0

    def _ws(self):
        while self.i < len(self.t) and self.t[self.i] in WS:
            self.i += 1

    def _peek(self) -> str:
        if self.i >= len(self.t):
            raise JsonEditError("JSON troncato")
        return self.t[self.i]

    def parse(self) -> Node:
        self._ws()
        n = self._value()
        return n

    def _value(self) -> Node:
        c = self._peek()
        if c == "{":
            return self._obj()
        if c == "[":
            return self._arr()
        if c == '"':
            s = self.i
            v = self._string()
            return Node("string", s, self.i, v)
        s = self.i
        while self.i < len(self.t) and self.t[self.i] not in ",}]" + WS:
            self.i += 1
        raw = self.t[s : self.i]
        try:
            return Node("literal", s, self.i, json.loads(raw))
        except Exception as e:
            raise JsonEditError(f"valore non valido a offset {s}: {raw!r}") from e

    def _string(self) -> str:
        j = self.i + 1
        while True:
            if j >= len(self.t):
                raise JsonEditError("stringa non terminata")
            c = self.t[j]
            if c == "\\":
                j += 2
                continue
            if c == '"':
                j += 1
                break
            j += 1
        raw = self.t[self.i : j]
        self.i = j
        return json.loads(raw)

    def _obj(self) -> Node:
        s = self.i
        self.i += 1
        props: List[Prop] = []
        self._ws()
        if self._peek() == "}":
            self.i += 1
            return Node("object", s, self.i, None, props)
        while True:
            self._ws()
            pstart = self.i
            if self._peek() != '"':
                raise JsonEditError(f"chiave attesa a offset {self.i}")
            kstart = self.i
            key = self._string()
            kend = self.i
            self._ws()
            if self._peek() != ":":
                raise JsonEditError(f"':' atteso a offset {self.i}")
            self.i += 1
            self._ws()
            vn = self._value()
            pend = self.i
            props.append(Prop(key, kstart, kend, vn, pstart, pend))
            self._ws()
            c = self._peek()
            if c == ",":
                self.i += 1
                continue
            if c == "}":
                self.i += 1
                break
            raise JsonEditError(f"',' o '}}' atteso a offset {self.i}")
        return Node("object", s, self.i, None, props)

    def _arr(self) -> Node:
        s = self.i
        self.i += 1
        items: List[Node] = []
        self._ws()
        if self._peek() == "]":
            self.i += 1
            return Node("array", s, self.i, None, None, items)
        while True:
            self._ws()
            items.append(self._value())
            self._ws()
            c = self._peek()
            if c == ",":
                self.i += 1
                continue
            if c == "]":
                self.i += 1
                break
            raise JsonEditError(f"',' o ']' atteso a offset {self.i}")
        return Node("array", s, self.i, None, None, items)


def parse_tree(text: str) -> Node:
    return _Parser(strip_bom(text)).parse()


def strip_bom(text: str) -> str:
    return text[1:] if text[:1] == "﻿" else text


def find_node(root: Node, path: JsonPath) -> Optional[Node]:
    node = root
    for seg in path:
        if isinstance(seg, int):
            if node.type != "array" or seg < 0 or seg >= len(node.items):
                return None
            node = node.items[seg]
        else:
            if node.type != "object":
                return None
            found = None
            for p in node.props:
                if p.key == seg:
                    found = p.value
                    break
            if found is None:
                return None
            node = found
    return node


def path_exists(text: str, path: JsonPath) -> bool:
    try:
        root = parse_tree(text)
    except JsonEditError:
        return False
    return find_node(root, path) is not None


# ---------------------------------------------------------------- formattazione


def detect_eol(text: str) -> str:
    return "\r\n" if "\r\n" in text else "\n"


def detect_indent(text: str) -> int:
    import re

    m = re.search(r"^\{?\s*?\n([ \t]+)\S", text, re.M) or re.search(r"\n([ ]+)\"", text)
    if m and m.group(1):
        spaces = len(m.group(1).replace("\t", "  "))
        return max(1, min(spaces, 8))
    return 2


def _line_indent(text: str, offset: int) -> str:
    """Indentazione (whitespace) della riga in cui cade offset."""
    start = text.rfind("\n", 0, offset) + 1
    i = start
    while i < len(text) and text[i] in " \t":
        i += 1
    return text[start:i]


def _serialize(value: Any, indent: str, tab: int, eol: str) -> str:
    """Serializza value indentandolo come se stesse alla colonna `indent`."""
    raw = json.dumps(value, indent=tab, ensure_ascii=False)
    lines = raw.split("\n")
    out = [lines[0]] + [indent + ln for ln in lines[1:]]
    return eol.join(out)


def _nest(remainder: JsonPath, leaf: Any) -> Any:
    """Costruisce il valore annidato per la parte di percorso mancante."""
    value = leaf
    for seg in reversed(list(remainder)):
        if isinstance(seg, int):
            value = [{} for _ in range(seg)] + [value]
        else:
            value = {seg: value}
    return value


# ---------------------------------------------------------------- edit


def set_value(text: str, path: JsonPath, value: Any) -> str:
    path = list(path)
    if not path:
        raise JsonEditError("percorso vuoto")
    root = parse_tree(text)
    tab = detect_indent(text)
    eol = detect_eol(text)

    depth = 0
    while depth < len(path) and find_node(root, path[: depth + 1]) is not None:
        depth += 1

    if depth == len(path):
        node = find_node(root, path)
        assert node is not None
        indent = _line_indent(text, node.offset)
        ser = _serialize(value, indent, tab, eol)
        return text[: node.offset] + ser + text[node.end :]

    parent = find_node(root, path[:depth]) if depth else root
    if parent is None:
        raise JsonEditError("parent non trovato")
    seg = path[depth]
    nested = _nest(path[depth + 1 :], value)

    if isinstance(seg, int):
        if parent.type != "array":
            raise JsonEditError(f"il percorso {'.'.join(map(str, path[:depth]))} non e' un array")
        if seg != len(parent.items):
            raise JsonEditError(f"indice {seg} non contiguo in un array di {len(parent.items)} elementi")
        parent_indent = _line_indent(text, parent.offset)
        if parent.items:
            last = parent.items[-1]
            item_indent = _line_indent(text, last.offset)
            ser = _serialize(nested, item_indent, tab, eol)
            return text[: last.end] + "," + eol + item_indent + ser + text[last.end :]
        item_indent = parent_indent + " " * tab
        ser = _serialize(nested, item_indent, tab, eol)
        return (
            text[: parent.offset]
            + "[" + eol + item_indent + ser + eol + parent_indent + "]"
            + text[parent.end :]
        )

    if parent.type != "object":
        raise JsonEditError(f"il percorso {'.'.join(map(str, path[:depth]))} non e' un oggetto")
    parent_indent = _line_indent(text, parent.offset)
    key = json.dumps(seg, ensure_ascii=False)
    if parent.props:
        last = parent.props[-1]
        prop_indent = _line_indent(text, last.pstart)
        ser = _serialize(nested, prop_indent, tab, eol)
        return text[: last.pend] + "," + eol + prop_indent + key + ": " + ser + text[last.pend :]
    prop_indent = parent_indent + " " * tab
    ser = _serialize(nested, prop_indent, tab, eol)
    return (
        text[: parent.offset]
        + "{" + eol + prop_indent + key + ": " + ser + eol + parent_indent + "}"
        + text[parent.end :]
    )


def remove_value(text: str, path: JsonPath) -> str:
    """Rimuove la proprieta' indicata. Se non esiste, il testo resta invariato."""
    path = list(path)
    if not path:
        raise JsonEditError("percorso vuoto")
    root = parse_tree(text)
    if find_node(root, path) is None:
        return text
    parent = find_node(root, path[:-1]) if len(path) > 1 else root
    seg = path[-1]
    if isinstance(seg, int):
        if parent is None or parent.type != "array":
            return text
        items = parent.items
        k = seg
        if len(items) == 1:
            return text[: parent.offset] + "[]" + text[parent.end :]
        if k > 0:
            return text[: items[k - 1].end] + text[items[k].end :]
        return text[: items[0].offset] + text[items[1].offset :]

    if parent is None or parent.type != "object":
        return text
    props = parent.props
    k = next((i for i, p in enumerate(props) if p.key == seg), None)
    if k is None:
        return text
    if len(props) == 1:
        return text[: parent.offset] + "{}" + text[parent.end :]
    if k > 0:
        return text[: props[k - 1].pend] + text[props[k].pend :]
    return text[: props[0].pstart] + text[props[1].pstart :]


def apply_edits(text: str, edits: Sequence[dict]) -> str:
    """
    edits: lista di {"path": [...], "value": ...}. value assente/None-sentinel
    (chiave "value" mancante) significa rimozione.
    """
    out = strip_bom(text)
    for e in edits:
        if "value" not in e or e.get("_remove"):
            out = remove_value(out, e["path"])
        else:
            out = set_value(out, e["path"], e["value"])
    return out


def write_utf8_no_bom(file_path: str, text: str) -> None:
    """
    Salva in UTF-8 senza BOM e verifica dopo il salvataggio che i primi byte
    non siano EF BB BF: Power BI Desktop non gestisce il BOM e fallisce il parsing.
    """
    clean = strip_bom(text)
    buf = clean.encode("utf-8")
    if buf[:3] == b"\xef\xbb\xbf":
        raise JsonEditError(f"Il contenuto da scrivere contiene ancora un BOM UTF-8: {file_path}")
    with open(file_path, "wb") as f:
        f.write(buf)
    with open(file_path, "rb") as f:
        head = f.read(3)
    if head == b"\xef\xbb\xbf":
        raise JsonEditError(f"Verifica fallita: il file salvato inizia con BOM UTF-8: {file_path}")
