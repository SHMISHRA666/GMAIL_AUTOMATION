from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


VARIABLE_RE = re.compile(r"{{\s*([^}|]+?)(?:\s*\|[^}]*)?\s*}}")
SIMPLE_EXPR_RE = re.compile(r"^[A-Za-z_][\w.]*$")


@dataclass(frozen=True)
class RenderedTemplate:
    text: str
    variables: set[str]


def extract_liquid_variables(template_text: str) -> set[str]:
    variables: set[str] = set()
    for match in VARIABLE_RE.finditer(template_text or ""):
        expression = match.group(1).strip()
        if expression:
            variables.add(expression)
    return variables


def render_liquid_template(template_text: str, context: dict[str, Any]) -> RenderedTemplate:
    variables = extract_liquid_variables(template_text)
    missing = [name for name in sorted(variables) if _resolve_context_path(context, name) is None]
    if missing:
        raise ValueError(f"Missing template variable values: {', '.join(missing)}")
    template_text = _render_raw_column_placeholders(template_text, context)
    try:
        from liquid import Environment, StrictUndefined

        env = Environment(undefined=StrictUndefined)
        rendered = env.from_string(template_text).render(context)
    except ImportError as exc:
        raise RuntimeError("python-liquid is required for Liquid template rendering") from exc
    return RenderedTemplate(text=rendered, variables=variables)


def build_nested_context(flat_values: dict[str, str]) -> dict[str, Any]:
    context: dict[str, Any] = {}
    for key, value in flat_values.items():
        context[key] = value
        parts = key.split(".")
        cursor = context
        for part in parts[:-1]:
            child = cursor.get(part)
            if not isinstance(child, dict):
                child = {}
                cursor[part] = child
            cursor = child
        cursor[parts[-1]] = value
    return context


def _resolve_context_path(context: dict[str, Any], path: str) -> Any:
    if path in context:
        return context[path]
    cursor: Any = context
    for part in path.split("."):
        if isinstance(cursor, dict) and part in cursor:
            cursor = cursor[part]
        else:
            return None
    return cursor


def _render_raw_column_placeholders(template_text: str, context: dict[str, Any]) -> str:
    def replace(match: re.Match[str]) -> str:
        expression = match.group(1).strip()
        if SIMPLE_EXPR_RE.match(expression):
            return match.group(0)
        value = _resolve_context_path(context, expression)
        return "" if value is None else str(value)

    return VARIABLE_RE.sub(replace, template_text)
