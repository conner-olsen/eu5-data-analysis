"""Parser for Paradox Engine script files (.txt format)."""

import re
from pathlib import Path


def tokenize(text: str) -> list[str]:
    """Split Paradox script text into tokens."""
    tokens = []
    i = 0
    while i < len(text):
        c = text[i]

        # Skip whitespace
        if c in " \t\r\n":
            i += 1
            continue

        # Comments - skip to end of line
        if c == "#":
            while i < len(text) and text[i] != "\n":
                i += 1
            continue

        # Braces and operators
        if c in "{}":
            tokens.append(c)
            i += 1
            continue

        # Comparison operators
        if c in "=!<>?" and i + 1 < len(text) and text[i + 1] == "=":
            tokens.append(text[i : i + 2])
            i += 2
            continue
        if c == "=":
            tokens.append("=")
            i += 1
            continue

        # Quoted string
        if c == '"':
            j = i + 1
            while j < len(text) and text[j] != '"':
                j += 1
            tokens.append(text[i + 1 : j])  # strip quotes
            i = j + 1
            continue

        # Unquoted word/number (includes colons, dots, underscores, minus for negatives)
        if c not in "{}=\n\r\t ":
            j = i
            while j < len(text) and text[j] not in " \t\r\n{}=#\"":
                # Stop at operators like != ?= >= <=
                if (
                    text[j] in "!?<>"
                    and j + 1 < len(text)
                    and text[j + 1] == "="
                ):
                    break
                j += 1
            tokens.append(text[i:j])
            i = j
            continue

        i += 1

    return tokens


def _parse_value(val: str):
    """Convert a string token to appropriate Python type."""
    if val == "yes":
        return True
    if val == "no":
        return False
    try:
        return int(val)
    except ValueError:
        pass
    try:
        return float(val)
    except ValueError:
        pass
    return val


def parse_block(tokens: list[str], pos: int) -> tuple[dict, int]:
    """Parse a block of key=value pairs. Returns (dict, new_position).

    Handles duplicate keys by converting to lists.
    """
    result = {}

    while pos < len(tokens) and tokens[pos] != "}":
        key = tokens[pos]
        pos += 1

        # Check for operator (=, !=, ?=, >=, <=)
        if pos < len(tokens) and tokens[pos] in ("=", "!=", "?=", ">=", "<="):
            op = tokens[pos]
            pos += 1
        else:
            # Bare value (e.g., inside a list like `gfx_tags = { tag1 tag2 }`)
            # Store as a list item
            if "__bare_values__" not in result:
                result["__bare_values__"] = []
            result["__bare_values__"].append(_parse_value(key))
            continue

        if pos >= len(tokens):
            break

        # Value is either a block or a scalar
        if tokens[pos] == "{":
            pos += 1  # skip opening brace
            block, pos = parse_block(tokens, pos)
            if pos < len(tokens) and tokens[pos] == "}":
                pos += 1  # skip closing brace
            value = block
        else:
            value = _parse_value(tokens[pos])
            pos += 1
            # Handle composite values like "rgb { 242 242 111 }"
            if pos < len(tokens) and tokens[pos] == "{":
                pos += 1  # skip opening brace
                block, pos = parse_block(tokens, pos)
                if pos < len(tokens) and tokens[pos] == "}":
                    pos += 1  # skip closing brace
                value = {str(value): block}

        # For non-= operators, store as a special dict
        if op != "=":
            value = {"__op__": op, "__value__": value}

        # Handle duplicate keys
        if key in result:
            existing = result[key]
            if isinstance(existing, list):
                existing.append(value)
            else:
                result[key] = [existing, value]
        else:
            result[key] = value

    return result, pos


def parse_file(filepath: str | Path) -> dict:
    """Parse a Paradox script file and return top-level definitions as a dict."""
    path = Path(filepath)
    text = path.read_text(encoding="utf-8-sig")  # handle BOM
    tokens = tokenize(text)
    result, _ = parse_block(tokens, 0)

    # Convert bare values at top level to proper format
    if "__bare_values__" in result:
        del result["__bare_values__"]

    return result


def parse_directory(dirpath: str | Path, pattern: str = "*.txt") -> dict:
    """Parse all matching files in a directory, merging results."""
    dirpath = Path(dirpath)
    merged = {}
    for filepath in sorted(dirpath.glob(pattern)):
        if filepath.name.startswith("readme") or filepath.name.endswith(".info"):
            continue
        parsed = parse_file(filepath)
        merged.update(parsed)
    return merged
