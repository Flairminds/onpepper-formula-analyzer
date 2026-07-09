# src/ast_builder.py
"""
AST Generation – Convert each Excel formula into an Abstract Syntax Tree.

An AST shows the STRUCTURE of a formula, not just its tokens.
Example:  =SUM(B10:C10)+D5

    BinaryOp[+]
    ├── Function[SUM]
    │     └── Range[B10:C10]
    └── CellRef[D5]

How it works:
  1. openpyxl's Tokenizer turns the formula string into a flat token list.
  2. A _Stream class wraps that list so we can peek / advance easily.
  3. A recursive descent parser builds the tree, one precedence level at a time:
         comparison  (=, <>, <=, >=, <, >)   ← lowest precedence
         concat      (&)
         additive    (+, -)
         multiplicative (*, /)
         exponent    (^)
         unary       (-, %)
         primary     (function call, cell ref, literal, group)  ← highest
  4. Every node is a plain Python dict so it is JSON-serialisable.

Public API:
    build_ast(formula)            → AST dict for one formula
    build_all_asts(formulas)      → {sheet: {cell: ast}}
    ast_to_text(ast)              → pretty-printed tree string
    get_ast_stats(ast)            → {depth, node_count, leaf_count}
"""

from typing import Any, Dict, List, Optional
from openpyxl.formula import Tokenizer


# ─────────────────────────────────────────────────────────────
# 1. Token stream helper
# ─────────────────────────────────────────────────────────────

class _Stream:
    """Wraps a flat token list and provides peek / advance operations."""

    def __init__(self, tokens):
        # Drop whitespace tokens — they carry no structural meaning
        self.tokens = [t for t in tokens if t.type != "WSPACE"]
        self.pos = 0

    @property
    def current(self):
        """Token at the current position, or None when exhausted."""
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def advance(self):
        """Return the current token and move one step forward."""
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def is_infix(self, *values):
        """True if the current token is an infix operator matching one of values."""
        tok = self.current
        return tok is not None and tok.type == "OPERATOR-INFIX" and tok.value in values


# ─────────────────────────────────────────────────────────────
# 2. AST node constructors  (all return plain dicts)
# ─────────────────────────────────────────────────────────────

def _bin(op, left, right):
    return {"type": "BinaryOp", "op": op, "left": left, "right": right}

def _unary(op, operand):
    return {"type": "UnaryOp", "op": op, "operand": operand}

def _func(name, args):
    return {"type": "Function", "name": name, "args": args}

def _cell(ref, sheet=None):
    node = {"type": "CellRef", "ref": ref}
    if sheet:
        node["sheet"] = sheet
    return node

def _range(ref, sheet=None):
    node = {"type": "Range", "ref": ref}
    if sheet:
        node["sheet"] = sheet
    return node

def _number(raw: str):
    try:
        v = float(raw) if ("." in raw or "e" in raw.lower()) else int(raw)
    except ValueError:
        v = raw
    return {"type": "Number", "value": v}

def _string(raw: str):
    return {"type": "String", "value": raw.strip('"')}

def _boolean(raw: str):
    return {"type": "Boolean", "value": raw.upper() == "TRUE"}

def _error(raw: str):
    return {"type": "Error", "value": raw}

def _group(expr):
    return {"type": "Group", "expr": expr}


# ─────────────────────────────────────────────────────────────
# 3. Recursive descent parser
#    Each level handles one layer of operator precedence.
# ─────────────────────────────────────────────────────────────

def _parse(stream: _Stream) -> Dict[str, Any]:
    """Top-level entry: parse a full expression."""
    return _comparison(stream)


def _comparison(stream: _Stream) -> Dict[str, Any]:
    """Lowest precedence: comparison operators =, <>, <=, >=, <, >"""
    node = _concat(stream)
    while stream.is_infix("=", "<>", "<=", ">=", "<", ">"):
        op    = stream.advance().value
        right = _concat(stream)
        node  = _bin(op, node, right)
    return node


def _concat(stream: _Stream) -> Dict[str, Any]:
    """String concatenation: &"""
    node = _additive(stream)
    while stream.is_infix("&"):
        op    = stream.advance().value
        right = _additive(stream)
        node  = _bin(op, node, right)
    return node


def _additive(stream: _Stream) -> Dict[str, Any]:
    """Addition and subtraction: +, -"""
    node = _multiplicative(stream)
    while stream.is_infix("+", "-"):
        op    = stream.advance().value
        right = _multiplicative(stream)
        node  = _bin(op, node, right)
    return node


def _multiplicative(stream: _Stream) -> Dict[str, Any]:
    """Multiplication and division: *, /"""
    node = _exponent(stream)
    while stream.is_infix("*", "/"):
        op    = stream.advance().value
        right = _exponent(stream)
        node  = _bin(op, node, right)
    return node


def _exponent(stream: _Stream) -> Dict[str, Any]:
    """Exponentiation: ^ (right-associative — recurse instead of loop)"""
    node = _unary_node(stream)
    if stream.is_infix("^"):
        op    = stream.advance().value
        right = _exponent(stream)          # recurse for right-to-left
        return _bin(op, node, right)
    return node


def _unary_node(stream: _Stream) -> Dict[str, Any]:
    """Prefix minus and postfix percent."""
    tok = stream.current

    # Prefix: -A1  →  UnaryOp[-] → CellRef[A1]
    if tok and tok.type == "OPERATOR-PREFIX" and tok.value == "-":
        stream.advance()
        return _unary("-", _primary(stream))

    node = _primary(stream)

    # Postfix: 50%  →  UnaryOp[%] → Number[50]
    tok = stream.current
    if tok and tok.type == "OPERATOR-POSTFIX" and tok.value == "%":
        stream.advance()
        return _unary("%", node)

    return node


def _primary(stream: _Stream) -> Dict[str, Any]:
    """Highest precedence: literals, cell refs, function calls, groups."""
    tok = stream.current

    if tok is None:
        return {"type": "ParseError", "value": "unexpected end of formula"}

    # ── Function call: openpyxl gives FUNC / OPEN, e.g. "SUM(", "IF(" ────────
    if tok.type == "FUNC" and tok.subtype == "OPEN":
        name = tok.value.rstrip("(").strip()
        stream.advance()                       # consume the "SUM(" token
        return _func(name.upper(), _args(stream))

    # ── Standalone parenthesised group: (expr) ───────────────────────────────
    # openpyxl gives PAREN / OPEN for a bare "(" (not preceded by a function name)
    if tok.type == "PAREN" and tok.subtype == "OPEN":
        stream.advance()                       # consume "("
        expr = _comparison(stream)
        _eat_close(stream)
        return _group(expr)

    # ── Cell or range reference ──────────────────────────────────────────────
    if tok.type == "OPERAND" and tok.subtype == "RANGE":
        stream.advance()
        ref = tok.value.replace("$", "").upper()

        if "!" in ref:
            sheet_part, cell_part = ref.rsplit("!", 1)
            sheet_name = sheet_part.strip("'").replace("''", "'")
            if ":" in cell_part:
                return _range(cell_part, sheet=sheet_name)
            else:
                return _cell(cell_part, sheet=sheet_name)

        if ":" in ref:
            return _range(ref)

        return _cell(ref)

    # ── Numeric literal ──────────────────────────────────────────────────────
    if tok.type == "OPERAND" and tok.subtype == "NUMBER":
        stream.advance()
        return _number(tok.value)

    # ── String literal ───────────────────────────────────────────────────────
    if tok.type == "OPERAND" and tok.subtype == "TEXT":
        stream.advance()
        return _string(tok.value)

    # ── Boolean literal ──────────────────────────────────────────────────────
    if tok.type == "OPERAND" and tok.subtype == "LOGICAL":
        stream.advance()
        return _boolean(tok.value)

    # ── Excel error literal (#REF!, #VALUE!, …) ──────────────────────────────
    if tok.type == "OPERAND" and tok.subtype == "ERROR":
        stream.advance()
        return _error(tok.value)

    # ── Anything else (unknown token) ────────────────────────────────────────
    stream.advance()
    return {"type": "Unknown", "value": tok.value}


def _args(stream: _Stream) -> List[Dict[str, Any]]:
    """Parse a comma-separated argument list until FUNC CLOSE.

    Handles:
        SUM()           – zero arguments
        IF(A1>0,1,0)    – multiple arguments separated by SEP tokens
    """
    args: List[Dict[str, Any]] = []

    # Empty call: SUM()
    tok = stream.current
    if tok and tok.type == "FUNC" and tok.subtype == "CLOSE":
        stream.advance()
        return args

    # First argument
    args.append(_comparison(stream))

    # Subsequent arguments (each preceded by a SEP / comma token)
    while stream.current and stream.current.type == "SEP":
        stream.advance()                       # consume ","
        args.append(_comparison(stream))

    _eat_close(stream)
    return args


def _eat_close(stream: _Stream):
    """Consume a closing ')' token if present (resilient — skip if missing).

    A function call's ')' tokenizes as FUNC/CLOSE; a standalone group's as PAREN/CLOSE.
    """
    tok = stream.current
    if tok and tok.subtype == "CLOSE" and tok.type in ("FUNC", "PAREN"):
        stream.advance()


# ─────────────────────────────────────────────────────────────
# 4. AST statistics
# ─────────────────────────────────────────────────────────────

def get_ast_stats(node: Dict[str, Any]) -> Dict[str, int]:
    """Return depth, total node count, and leaf count for an AST.

    Useful for measuring formula complexity:
        depth      – how deeply nested the formula is
        node_count – total nodes (operators + functions + leaves)
        leaf_count – terminal nodes (cell refs, ranges, literals)
    """
    def _walk(n, depth):
        """Returns (max_depth, node_count, leaf_count)."""
        if n is None:
            return depth, 0, 0

        t = n.get("type", "")

        if t == "BinaryOp":
            ld, ln, ll = _walk(n["left"],    depth + 1)
            rd, rn, rl = _walk(n["right"],   depth + 1)
            return max(ld, rd), ln + rn + 1, ll + rl

        if t == "UnaryOp":
            d, n_cnt, l = _walk(n["operand"], depth + 1)
            return d, n_cnt + 1, l

        if t == "Function":
            if not n["args"]:
                return depth, 1, 1
            max_d, total_n, total_l = depth, 1, 0
            for arg in n["args"]:
                d, nc, lc = _walk(arg, depth + 1)
                max_d  = max(max_d, d)
                total_n += nc
                total_l += lc
            return max_d, total_n, total_l

        if t == "Group":
            d, nc, lc = _walk(n["expr"], depth + 1)
            return d, nc + 1, lc

        # Leaf node (CellRef, Range, Number, String, Boolean, Error, Unknown)
        return depth, 1, 1

    max_depth, node_count, leaf_count = _walk(node, 1)
    return {
        "depth":      max_depth,
        "node_count": node_count,
        "leaf_count": leaf_count,
    }


# ─────────────────────────────────────────────────────────────
# 5. Pretty-print
# ─────────────────────────────────────────────────────────────

def ast_to_text(node: Dict[str, Any], _indent: int = 0) -> str:
    """Return an indented string representation of an AST.

    Example output for  =SUM(B10:C10)+D5 :

        BinaryOp[+]
          Function[SUM]
            Range[B10:C10]
          CellRef[D5]
    """
    pad = "  " * _indent
    t   = node.get("type", "?")

    if t == "BinaryOp":
        return (
            f"{pad}BinaryOp[{node['op']}]\n"
            + ast_to_text(node["left"],    _indent + 1) + "\n"
            + ast_to_text(node["right"],   _indent + 1)
        )

    if t == "UnaryOp":
        return (
            f"{pad}UnaryOp[{node['op']}]\n"
            + ast_to_text(node["operand"], _indent + 1)
        )

    if t == "Function":
        header = f"{pad}Function[{node['name']}]"
        if not node["args"]:
            return header
        args_text = "\n".join(ast_to_text(a, _indent + 1) for a in node["args"])
        return f"{header}\n{args_text}"

    if t == "Group":
        return f"{pad}Group\n" + ast_to_text(node["expr"], _indent + 1)

    if t == "CellRef":
        sheet = f" sheet={node['sheet']}" if node.get("sheet") else ""
        return f"{pad}CellRef[{node['ref']}{sheet}]"

    if t == "Range":
        sheet = f" sheet={node['sheet']}" if node.get("sheet") else ""
        return f"{pad}Range[{node['ref']}{sheet}]"

    if t == "Number":
        return f"{pad}Number[{node['value']}]"

    if t == "String":
        return f'{pad}String["{node["value"]}"]'

    if t == "Boolean":
        return f"{pad}Boolean[{node['value']}]"

    if t in ("Error", "ParseError", "Unknown"):
        return f"{pad}{t}[{node.get('value', '')}]"

    return f"{pad}{t}"


# ─────────────────────────────────────────────────────────────
# 6. Public API
# ─────────────────────────────────────────────────────────────

def build_ast(formula: str) -> Dict[str, Any]:
    """Parse one Excel formula and return its AST as a JSON-serialisable dict.

    Parameters
    ----------
    formula : str
        Raw formula string, e.g. ``"=SUM(B10:C10)+D5"``

    Returns
    -------
    dict
        AST node.  On parse errors, returns
        ``{"type": "ParseError", "formula": ..., "error": ...}``
        so the caller never crashes.
    """
    try:
        tok    = Tokenizer(formula)
        stream = _Stream(tok.items)
        return _parse(stream)
    except Exception as exc:
        return {"type": "ParseError", "formula": formula, "error": str(exc)}


def build_all_asts(
    formulas: Dict[str, Dict[str, str]],
) -> Dict[str, Any]:
    """Build ASTs for every formula cell in the workbook.

    Parameters
    ----------
    formulas : dict
        ``{sheet_name: {cell_address: formula_string}}``
        — output of ``formula_extractor.extract_formulas()``

    Returns
    -------
    dict
        ``{sheet_name: {cell_address: {"ast": ..., "stats": ...}}}``

        Each cell entry contains:
            ast   – full AST tree (JSON dict)
            stats – {depth, node_count, leaf_count}

    Example
    -------
    {
      "Sheet1": {
        "B15": {
          "ast": {"type": "Function", "name": "SUM", "args": [...]},
          "stats": {"depth": 2, "node_count": 3, "leaf_count": 2}
        }
      }
    }
    """
    result: Dict[str, Any] = {}

    for sheet_name, sheet_formulas in formulas.items():
        result[sheet_name] = {}
        for cell, formula in sheet_formulas.items():
            ast   = build_ast(formula)
            stats = get_ast_stats(ast)
            result[sheet_name][cell] = {
                "ast":   ast,
                "stats": stats,
            }

    return result
