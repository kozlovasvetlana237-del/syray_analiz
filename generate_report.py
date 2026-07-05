"""Generate a standalone CFO financial dashboard from two Excel P&L files."""

from __future__ import annotations

import html
import json
import logging
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import plotly.graph_objects as go
from plotly.io import to_html


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "financial_dashboard.html"

MONTHS = {
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}
MONTH_PATTERN = re.compile(
    r"^(Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь)\s+20\d{2}$",
    re.IGNORECASE,
)

WHAT_IF_ARTICLES = [
    "Аренда региональных офисов",
    "Текущий ремонт помещений",
    "Содержание отдела обслуживания",
    "Ремонт и обслуживание прочего оборудования и инвентаря",
    "Ремонт и обслуживание кофемашин",
    "Мебель для кухни",
    "Оргтехника",
    "Доставка материалов, инвентаря и мелкого оборудования, мебели и оргтехники",
    "Аренда оборудования, лизинг",
    "Услуги сайтов по трудоустройству и подбору персонала",
    "Транспортное обеспечение сотрудников",
    "Аренда жилья для сотрудников",
    "Командировки",
    "Управление",
    "Оплата труда АУП",
]

PLOTLY_CONFIG = {
    "displaylogo": False,
    "responsive": True,
    "toImageButtonOptions": {
        "format": "png",
        "filename": "financial_chart",
        "height": 720,
        "width": 1280,
        "scale": 2,
    },
}


@dataclass(frozen=True)
class MonthColumn:
    """Excel column that contains RUR values for a detected month."""

    period: str
    column: int


@dataclass
class LocationDataset:
    """Parsed monthly financial data for one location."""

    code: str
    name: str
    rows: list[dict[str, Any]]
    warnings: list[str]


def normalize_text(value: Any) -> str:
    """Normalize text for robust matching of Russian labels."""

    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def to_number(value: Any) -> float:
    """Convert Excel cell values to float and treat blanks as zero."""

    if value is None:
        return 0.0
    if isinstance(value, float) and math.isnan(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def period_sort_key(period: str) -> tuple[int, int]:
    """Return chronological sort key for a Russian month label."""

    month_name, year = period.split()
    return int(year), MONTHS[normalize_text(month_name)]


def format_rub(value: float) -> str:
    """Format numeric values as Russian rubles."""

    return f"{round(value):,}".replace(",", " ")


def format_pct(value: float) -> str:
    """Format ratios as percent strings."""

    if not math.isfinite(value):
        return "0,0%"
    return f"{value * 100:.1f}%".replace(".", ",")


def read_excel_matrix(path: Path) -> pd.DataFrame:
    """Read the first Excel sheet as a raw matrix without headers."""

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    return pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")


def detect_month_columns(df: pd.DataFrame) -> list[MonthColumn]:
    """Detect the row and columns that contain monthly RUR values."""

    best: list[MonthColumn] = []
    for _, row in df.iterrows():
        found: list[MonthColumn] = []
        for col_idx, value in row.items():
            label = str(value).strip() if not pd.isna(value) else ""
            if MONTH_PATTERN.match(label):
                found.append(MonthColumn(period=label, column=int(col_idx)))
        if len(found) > len(best):
            best = found

    if not best:
        raise ValueError("Could not detect month columns in workbook.")
    return sorted(best, key=lambda item: item.column)


def row_label(df: pd.DataFrame, row_idx: int) -> str:
    """Build a label string from the descriptor columns of a row."""

    values = [df.iat[row_idx, col] for col in range(min(5, df.shape[1]))]
    return " | ".join(str(value) for value in values if not pd.isna(value)).strip()


def find_row(
    df: pd.DataFrame,
    aliases: Iterable[str],
    *,
    start: int = 0,
    end: int | None = None,
    exact: bool = True,
) -> int | None:
    """Find the first row matching any alias in descriptor columns."""

    normalized_aliases = [normalize_text(alias) for alias in aliases]
    end = df.shape[0] if end is None else min(end, df.shape[0])
    for row_idx in range(start, end):
        label = normalize_text(row_label(df, row_idx))
        if not label:
            continue
        if exact and any(label == alias for alias in normalized_aliases):
            return row_idx
        if not exact and any(alias in label for alias in normalized_aliases):
            return row_idx
    return None


def extract_series(
    df: pd.DataFrame,
    row_idx: int | None,
    months: list[MonthColumn],
) -> dict[str, float]:
    """Extract month values from a specific row."""

    if row_idx is None:
        return {month.period: 0.0 for month in months}
    return {month.period: to_number(df.iat[row_idx, month.column]) for month in months}


def warn_missing(warnings: list[str], location: str, name: str) -> None:
    """Append a standardized warning for a missing source row."""

    warnings.append(f"{location}: не найдена статья '{name}', значение принято равным 0.")


def must_series(
    df: pd.DataFrame,
    months: list[MonthColumn],
    warnings: list[str],
    location: str,
    metric_name: str,
    row_idx: int | None,
) -> dict[str, float]:
    """Extract a series and register a warning if the source row is absent."""

    if row_idx is None:
        warn_missing(warnings, location, metric_name)
    return extract_series(df, row_idx, months)


def parse_location(path: Path, code: str, name: str) -> LocationDataset:
    """Parse one location workbook into normalized monthly records."""

    df = read_excel_matrix(path)
    months = detect_month_columns(df)
    warnings: list[str] = []

    revenue_row = find_row(df, ["Доходы(Выручка)"])
    cogs_row = find_row(df, ["Себестоимость продукции"])
    expenses_row = find_row(df, ["Расходы"], start=(cogs_row or 0) + 1)
    store_ebitda_row = find_row(df, ["Store-level EBITDA"])
    mgmt_row = find_row(df, ["Расходы на общее и сетевое управление"])
    ebitda_row = find_row(df, ["EBITDA"])
    net_row = find_row(df, ["Чистая прибыль(убыток)"])

    revenue = must_series(df, months, warnings, name, "Доходы(Выручка)", revenue_row)
    cogs = must_series(df, months, warnings, name, "Себестоимость продукции", cogs_row)
    operating_expenses = must_series(df, months, warnings, name, "Расходы", expenses_row)
    store_ebitda = must_series(df, months, warnings, name, "Store-level EBITDA", store_ebitda_row)
    management = must_series(df, months, warnings, name, "Расходы на общее и сетевое управление", mgmt_row)
    ebitda = extract_series(df, ebitda_row, months)
    net_profit = must_series(df, months, warnings, name, "Чистая прибыль(убыток)", net_row)

    if ebitda_row is None:
        ebitda = {
            period: store_ebitda[period] - management[period]
            for period in revenue
        }
        warnings.append(f"{name}: EBITDA рассчитана как Store-level EBITDA минус сетевое управление.")

    revenue_channel_rows = {
        "delivery_revenue": find_row(df, ["Доставка"], start=(revenue_row or 0) + 1, end=cogs_row),
        "restaurant_revenue": find_row(df, ["Ресторан"], start=(revenue_row or 0) + 1, end=cogs_row),
        "pickup_revenue": find_row(df, ["Самовывоз"], start=(revenue_row or 0) + 1, end=cogs_row),
    }

    expense_structure_rows = {
        "product_cost": find_row(df, ["Продукты"], start=(cogs_row or 0) + 1),
        "packaging_cost": find_row(df, ["Упаковка"], start=(cogs_row or 0) + 1),
        "kitchen_labor": find_row(df, ["Оплата труда кухни"], start=(cogs_row or 0) + 1),
        "delivery_expense": find_row(df, ["Расходы на доставку"], start=(expenses_row or 0) + 1),
        "marketing": find_row(df, ["Маркетинг и реклама"], start=(expenses_row or 0) + 1),
        "management_labor": find_row(df, ["Оплата труда управления"], start=(expenses_row or 0) + 1),
        "premises": find_row(df, ["Помещения"], start=(expenses_row or 0) + 1),
        "employee_costs": find_row(df, ["Расходы на сотрудников"], start=(expenses_row or 0) + 1),
        "bank_services": find_row(df, ["Банковские услуги"], start=(expenses_row or 0) + 1),
        "other_opex": find_row(df, ["Прочие операционные расходы"], start=(expenses_row or 0) + 1),
        "taxes": find_row(df, ["Налоги и сборы"], start=(expenses_row or 0) + 1),
        "royalty": find_row(df, ["Роялти"], start=(expenses_row or 0) + 1),
        "network_management": mgmt_row,
    }

    channels = {
        key: extract_series(df, row_idx, months)
        for key, row_idx in revenue_channel_rows.items()
    }
    expense_structure = {
        key: extract_series(df, row_idx, months)
        for key, row_idx in expense_structure_rows.items()
    }

    what_if_rows: dict[str, dict[str, float]] = {}
    for article in WHAT_IF_ARTICLES:
        row_idx = find_row(df, [article], exact=True)
        if row_idx is None:
            warn_missing(warnings, name, article)
        what_if_rows[article] = extract_series(df, row_idx, months)

    records: list[dict[str, Any]] = []
    for month in months:
        period = month.period
        if revenue[period] == 0 and net_profit[period] == 0:
            continue
        gross_profit = revenue[period] - cogs[period]
        total_expenses = revenue[period] - net_profit[period]
        records.append(
            {
                "location": code,
                "location_name": name,
                "period": period,
                "revenue": revenue[period],
                "total_expenses": total_expenses,
                "cogs": cogs[period],
                "operating_expenses": operating_expenses[period],
                "gross_profit": gross_profit,
                "ebitda": ebitda[period],
                "net_profit": net_profit[period],
                "net_margin": net_profit[period] / revenue[period] if revenue[period] else 0.0,
                "revenue_growth": 0.0,
                "net_profit_growth": 0.0,
                "channels": {key: series[period] for key, series in channels.items()},
                "expense_structure": {key: series[period] for key, series in expense_structure.items()},
                "what_if": {article: series[period] for article, series in what_if_rows.items()},
            }
        )

    for idx, record in enumerate(records):
        if idx == 0:
            continue
        prev = records[idx - 1]
        record["revenue_growth"] = safe_growth(record["revenue"], prev["revenue"])
        record["net_profit_growth"] = safe_growth(record["net_profit"], prev["net_profit"])

    return LocationDataset(code=code, name=name, rows=records, warnings=warnings)


def safe_growth(current: float, previous: float) -> float:
    """Calculate growth ratio while handling zero base values."""

    if previous == 0:
        return 0.0
    return (current - previous) / abs(previous)


def sum_records(records: list[dict[str, Any]], key: str) -> float:
    """Sum a numeric key across records."""

    return sum(float(record.get(key, 0.0)) for record in records)


def consolidate(locations: list[LocationDataset]) -> list[dict[str, Any]]:
    """Build consolidated monthly records across all locations."""

    periods = sorted({row["period"] for loc in locations for row in loc.rows}, key=period_sort_key)
    output: list[dict[str, Any]] = []
    for period in periods:
        items = [row for loc in locations for row in loc.rows if row["period"] == period]
        if not items:
            continue
        revenue = sum_records(items, "revenue")
        net_profit = sum_records(items, "net_profit")
        row = {
            "location": "both",
            "location_name": "M1 + M3",
            "period": period,
            "revenue": revenue,
            "total_expenses": sum_records(items, "total_expenses"),
            "cogs": sum_records(items, "cogs"),
            "operating_expenses": sum_records(items, "operating_expenses"),
            "gross_profit": sum_records(items, "gross_profit"),
            "ebitda": sum_records(items, "ebitda"),
            "net_profit": net_profit,
            "net_margin": net_profit / revenue if revenue else 0.0,
            "revenue_growth": 0.0,
            "net_profit_growth": 0.0,
            "channels": merge_nested(items, "channels"),
            "expense_structure": merge_nested(items, "expense_structure"),
            "what_if": merge_nested(items, "what_if"),
        }
        output.append(row)

    for idx, record in enumerate(output):
        if idx == 0:
            continue
        prev = output[idx - 1]
        record["revenue_growth"] = safe_growth(record["revenue"], prev["revenue"])
        record["net_profit_growth"] = safe_growth(record["net_profit"], prev["net_profit"])
    return output


def merge_nested(records: list[dict[str, Any]], key: str) -> dict[str, float]:
    """Sum nested dictionaries across records."""

    names = sorted({name for record in records for name in record.get(key, {})})
    return {
        name: sum(float(record.get(key, {}).get(name, 0.0)) for record in records)
        for name in names
    }


def latest(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Return the latest monthly record."""

    return sorted(records, key=lambda row: period_sort_key(row["period"]))[-1]


def kpi_cards(records: list[dict[str, Any]]) -> str:
    """Render KPI cards for a tab."""

    item = latest(records)
    values = [
        ("Revenue", format_rub(item["revenue"])),
        ("Total Expenses", format_rub(item["total_expenses"])),
        ("Gross Profit", format_rub(item["gross_profit"])),
        ("EBITDA", format_rub(item["ebitda"])),
        ("Net Profit", format_rub(item["net_profit"])),
        ("Net Margin %", format_pct(item["net_margin"])),
        ("Revenue Growth %", format_pct(item["revenue_growth"])),
        ("Net Profit Growth %", format_pct(item["net_profit_growth"])),
    ]
    return "\n".join(
        f'<article class="kpi"><span>{html.escape(title)}</span><strong>{html.escape(value)}</strong></article>'
        for title, value in values
    )


def make_line_chart(
    title: str,
    records: list[dict[str, Any]],
    series: dict[str, str],
    *,
    include_plotlyjs: bool,
) -> str:
    """Create an offline Plotly line chart."""

    sorted_records = sorted(records, key=lambda row: period_sort_key(row["period"]))
    fig = go.Figure()
    for label, key in series.items():
        fig.add_trace(
            go.Scatter(
                x=[row["period"] for row in sorted_records],
                y=[row[key] for row in sorted_records],
                mode="lines+markers",
                name=label,
                hovertemplate="%{x}<br>%{fullData.name}: %{y:,.0f}<extra></extra>",
            )
        )
    fig.update_layout(
        title=title,
        template="plotly_white",
        height=360,
        margin={"l": 42, "r": 20, "t": 58, "b": 42},
        legend={"orientation": "h", "y": 1.08},
        yaxis_title="RUB",
    )
    return to_html(
        fig,
        include_plotlyjs=include_plotlyjs,
        full_html=False,
        config=PLOTLY_CONFIG,
    )


def make_comparison_chart(
    title: str,
    locations: list[LocationDataset],
    metric: str,
    *,
    include_plotlyjs: bool,
) -> str:
    """Create a location comparison chart."""

    fig = go.Figure()
    for loc in locations:
        rows = sorted(loc.rows, key=lambda row: period_sort_key(row["period"]))
        fig.add_trace(
            go.Scatter(
                x=[row["period"] for row in rows],
                y=[row[metric] for row in rows],
                mode="lines+markers",
                name=loc.name,
                hovertemplate="%{x}<br>%{fullData.name}: %{y:,.0f}<extra></extra>",
            )
        )
    fig.update_layout(
        title=title,
        template="plotly_white",
        height=360,
        margin={"l": 42, "r": 20, "t": 58, "b": 42},
        legend={"orientation": "h", "y": 1.08},
    )
    return to_html(fig, include_plotlyjs=include_plotlyjs, full_html=False, config=PLOTLY_CONFIG)


def expense_table(records: list[dict[str, Any]]) -> str:
    """Render an expense structure table."""

    labels = {
        "product_cost": "Продукты",
        "packaging_cost": "Упаковка",
        "kitchen_labor": "Оплата труда кухни",
        "delivery_expense": "Расходы на доставку",
        "marketing": "Маркетинг и реклама",
        "management_labor": "Оплата труда управления",
        "premises": "Помещения",
        "employee_costs": "Расходы на сотрудников",
        "bank_services": "Банковские услуги",
        "other_opex": "Прочие операционные расходы",
        "taxes": "Налоги и сборы",
        "royalty": "Роялти",
        "network_management": "Общее и сетевое управление",
    }
    totals = {
        key: sum(float(row["expense_structure"].get(key, 0.0)) for row in records)
        for key in labels
    }
    revenue = sum_records(records, "revenue")
    body = "\n".join(
        "<tr>"
        f"<td>{html.escape(labels[key])}</td>"
        f"<td>{format_rub(value)}</td>"
        f"<td>{format_pct(value / revenue if revenue else 0)}</td>"
        "</tr>"
        for key, value in sorted(totals.items(), key=lambda item: abs(item[1]), reverse=True)
    )
    return f"""
    <div class="table-wrap">
      <table>
        <thead><tr><th>Expense item</th><th>Total</th><th>% of revenue</th></tr></thead>
        <tbody>{body}</tbody>
      </table>
    </div>
    """


def comparative_table(locations: list[LocationDataset]) -> str:
    """Render location efficiency comparison."""

    rows = []
    for loc in locations:
        revenue = sum_records(loc.rows, "revenue")
        net = sum_records(loc.rows, "net_profit")
        expenses = sum_records(loc.rows, "total_expenses")
        ebitda = sum_records(loc.rows, "ebitda")
        rows.append((loc.name, revenue, expenses, ebitda, net, net / revenue if revenue else 0))
    body = "\n".join(
        "<tr>"
        f"<td>{name}</td><td>{format_rub(revenue)}</td><td>{format_rub(expenses)}</td>"
        f"<td>{format_rub(ebitda)}</td><td>{format_rub(net)}</td><td>{format_pct(margin)}</td>"
        "</tr>"
        for name, revenue, expenses, ebitda, net, margin in rows
    )
    return f"""
    <div class="table-wrap">
      <table>
        <thead><tr><th>Location</th><th>Revenue</th><th>Expenses</th><th>EBITDA</th><th>Net Profit</th><th>Net Margin</th></tr></thead>
        <tbody>{body}</tbody>
      </table>
    </div>
    """


def memo_items(locations: list[LocationDataset], consolidated: list[dict[str, Any]]) -> tuple[list[str], list[str]]:
    """Generate analytical memo conclusions and recommendations."""

    total_revenue = sum_records(consolidated, "revenue")
    total_expenses = sum_records(consolidated, "total_expenses")
    total_net = sum_records(consolidated, "net_profit")
    margin = total_net / total_revenue if total_revenue else 0
    expense_ratio = total_expenses / total_revenue if total_revenue else 0
    latest_total = latest(consolidated)
    first_total = sorted(consolidated, key=lambda row: period_sort_key(row["period"]))[0]

    conclusions = [
        f"Совокупная выручка за анализируемый период составила {format_rub(total_revenue)}, чистая прибыль - {format_rub(total_net)}.",
        f"Совокупная чистая маржа равна {format_pct(margin)}, доля всех расходов в выручке - {format_pct(expense_ratio)}.",
        f"Последний месяц: выручка {format_rub(latest_total['revenue'])}, ЧП {format_rub(latest_total['net_profit'])}, маржа {format_pct(latest_total['net_margin'])}.",
        f"Изменение выручки от первого к последнему месяцу: {format_pct(safe_growth(latest_total['revenue'], first_total['revenue']))}.",
        f"Изменение чистой прибыли от первого к последнему месяцу: {format_pct(safe_growth(latest_total['net_profit'], first_total['net_profit']))}.",
    ]
    for loc in locations:
        revenue = sum_records(loc.rows, "revenue")
        net = sum_records(loc.rows, "net_profit")
        conclusions.append(
            f"{loc.name}: вклад в выручку {format_pct(revenue / total_revenue if total_revenue else 0)}, "
            f"вклад в ЧП {format_pct(net / total_net if total_net else 0)}."
        )

    conclusions.extend(
        [
            "Главный риск - расхождение темпов выручки и расходов: при снижении продаж фиксированные и полуфиксированные расходы ухудшают маржу.",
            "Себестоимость и операционные расходы необходимо контролировать отдельно: они реагируют на разные управленческие рычаги.",
            "Административные и сетевые расходы требуют проверки базы распределения между точками.",
            "Сезонность следует оценивать через помесячную динамику, а не через сумму периода: пиковые расходы могут маскировать прибыльные месяцы.",
            "Точки отличаются по операционной эффективности, поэтому единые лимиты расходов без учета зрелости точки могут искажать мотивацию.",
        ]
    )

    recommendations = [
        "Ввести ежемесячный CFO-контроль: темп выручки, темп расходов, маржа ЧП, EBITDA и отклонение от целевого коридора.",
        "Установить целевые лимиты расходов как долю от выручки по каждой точке отдельно.",
        "Разделить расходы на переменные, полуфиксированные и фиксированные; для каждой группы назначить владельца бюджета.",
        "Оптимизировать себестоимость через контроль food cost, paper cost, списаний и закупочных условий.",
        "Пересмотреть административные расходы и сетевое управление: подтвердить экономическую пользу каждой статьи.",
        "Сократить расходы на подбор, транспорт, жилье и командировки через лимиты, заявки и постконтроль.",
        "Связать трудовые затраты с прогнозом спроса и фактической выручкой по дням недели.",
        "Приоритизировать инвестиции в каналы, где валовая маржа и повторяемость продаж выше.",
        "Ввести stop-loss правило: при падении выручки более чем на 5% MoM запускать пересмотр затрат в том же месяце.",
        "Для убыточной точки сфокусироваться на break-even плане: требуемая выручка, сокращение расходов и срок выхода в плюс.",
        "Использовать what-if модель ежемесячно перед утверждением бюджета следующего месяца.",
    ]
    return conclusions[:12], recommendations[:12]


def render_location_tab(
    loc_name: str,
    records: list[dict[str, Any]],
    *,
    include_plotlyjs: bool,
) -> tuple[str, bool]:
    """Render a full analytics tab for one dataset."""

    charts = [
        make_line_chart("Revenue by month", records, {"Revenue": "revenue"}, include_plotlyjs=include_plotlyjs),
        make_line_chart("Expenses by month", records, {"Total Expenses": "total_expenses"}, include_plotlyjs=False),
        make_line_chart("Net Profit by month", records, {"Net Profit": "net_profit"}, include_plotlyjs=False),
        make_line_chart(
            "Revenue / Expenses / Net Profit",
            records,
            {"Revenue": "revenue", "Expenses": "total_expenses", "Net Profit": "net_profit"},
            include_plotlyjs=False,
        ),
    ]
    section = f"""
    <section class="tab-panel" id="tab-{html.escape(loc_name)}">
      <div class="sticky-kpis">{kpi_cards(records)}</div>
      <div class="chart-grid">{''.join(f'<article class="panel">{chart}</article>' for chart in charts)}</div>
      <details open class="panel"><summary>Expense Structure</summary>{expense_table(records)}</details>
    </section>
    """
    return section, False


def render_consolidated_tab(
    locations: list[LocationDataset],
    consolidated: list[dict[str, Any]],
    *,
    include_plotlyjs: bool,
) -> tuple[str, bool]:
    """Render consolidated M1 + M3 analytics."""

    charts = [
        make_line_chart(
            "M1 + M3: Revenue / Expenses / Net Profit",
            consolidated,
            {"Revenue": "revenue", "Expenses": "total_expenses", "Net Profit": "net_profit"},
            include_plotlyjs=include_plotlyjs,
        ),
        make_comparison_chart("M1 vs M3 Revenue", locations, "revenue", include_plotlyjs=False),
        make_comparison_chart("M1 vs M3 Expenses", locations, "total_expenses", include_plotlyjs=False),
        make_comparison_chart("M1 vs M3 Net Profit", locations, "net_profit", include_plotlyjs=False),
    ]
    ranking = sorted(
        (
            {
                "name": loc.name,
                "net": sum_records(loc.rows, "net_profit"),
                "revenue": sum_records(loc.rows, "revenue"),
            }
            for loc in locations
        ),
        key=lambda row: row["net"],
        reverse=True,
    )
    ranking_html = "".join(
        f"<li><strong>{item['name']}</strong>: ЧП {format_rub(item['net'])}, выручка {format_rub(item['revenue'])}</li>"
        for item in ranking
    )
    section = f"""
    <section class="tab-panel" id="tab-both">
      <div class="sticky-kpis">{kpi_cards(consolidated)}</div>
      <div class="chart-grid">{''.join(f'<article class="panel">{chart}</article>' for chart in charts)}</div>
      <details open class="panel"><summary>Comparative efficiency table</summary>{comparative_table(locations)}</details>
      <details open class="panel"><summary>Ranking of locations</summary><ol>{ranking_html}</ol></details>
      <details open class="panel"><summary>Consolidated expense structure</summary>{expense_table(consolidated)}</details>
    </section>
    """
    return section, False


def render_what_if(consolidated: list[dict[str, Any]]) -> str:
    """Render WHAT IF tab shell; calculations run in browser JavaScript."""

    current_net = sum_records(consolidated, "net_profit")
    return f"""
    <section class="tab-panel" id="tab-whatif">
      <div class="whatif-head panel">
        <span>Current Net Profit</span>
        <strong id="currentNetProfit">{format_rub(current_net)}</strong>
      </div>
      <div class="table-wrap panel">
        <table id="whatIfTable"></table>
      </div>
      <div class="whatif-grid">
        <article class="panel metric"><span>Potential Net Profit Without Selected Costs</span><strong id="potentialNetProfit"></strong></article>
        <article class="panel metric"><span>Profit Increase</span><strong id="profitIncrease"></strong></article>
        <article class="panel metric"><span>Profit Increase %</span><strong id="profitIncreasePct"></strong></article>
        <article class="panel metric"><span>Savings Summary</span><strong id="savingsSummary"></strong></article>
      </div>
      <article class="panel"><div id="whatIfChart"></div></article>
    </section>
    """


def render_memo(locations: list[LocationDataset], consolidated: list[dict[str, Any]]) -> str:
    """Render analytical memo and recommendations."""

    conclusions, recommendations = memo_items(locations, consolidated)
    return f"""
    <section class="memo panel">
      <h2>Analytical Memo</h2>
      <ol>{''.join(f'<li>{html.escape(item)}</li>' for item in conclusions)}</ol>
      <h2>Recommendations</h2>
      <ol>{''.join(f'<li>{html.escape(item)}</li>' for item in recommendations)}</ol>
    </section>
    """


def build_payload(locations: list[LocationDataset], consolidated: list[dict[str, Any]]) -> dict[str, Any]:
    """Build JSON payload for browser-side WHAT IF calculations."""

    return {
        "locations": {loc.code: loc.rows for loc in locations},
        "consolidated": consolidated,
        "whatIfArticles": WHAT_IF_ARTICLES,
    }


def html_template(content: str, payload: dict[str, Any], warnings: list[str]) -> str:
    """Wrap report content into a self-contained HTML document."""

    warning_html = ""
    if warnings:
        warning_html = (
            '<details class="warnings"><summary>Warnings</summary><ul>'
            + "".join(f"<li>{html.escape(item)}</li>" for item in warnings)
            + "</ul></details>"
        )

    template = """<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Financial Dashboard</title>
  <style>
    :root { --bg:#f5f7fb; --panel:#ffffff; --ink:#142033; --muted:#66748a; --line:#dce3ef; --blue:#2563eb; --green:#059669; --red:#dc2626; --amber:#d97706; }
    * { box-sizing:border-box; }
    body { margin:0; font-family:Inter,Segoe UI,Arial,sans-serif; background:var(--bg); color:var(--ink); }
    header { position:sticky; top:0; z-index:10; background:rgba(255,255,255,.92); backdrop-filter:blur(14px); border-bottom:1px solid var(--line); padding:18px 28px; }
    .hero { display:flex; justify-content:space-between; gap:16px; align-items:center; flex-wrap:wrap; }
    h1 { margin:0; font-size:26px; letter-spacing:0; }
    .subtitle { color:var(--muted); margin-top:4px; }
    .tabs { display:flex; gap:8px; margin-top:16px; flex-wrap:wrap; }
    button { border:1px solid var(--line); border-radius:8px; background:#fff; color:var(--ink); padding:10px 14px; font-weight:700; cursor:pointer; }
    button.active { background:var(--blue); color:#fff; border-color:var(--blue); }
    main { padding:22px 28px 48px; display:grid; gap:18px; }
    .tab-panel { display:none; gap:18px; }
    .tab-panel.active { display:grid; }
    .sticky-kpis { position:sticky; top:118px; z-index:5; display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }
    .kpi,.panel { background:var(--panel); border:1px solid var(--line); border-radius:12px; box-shadow:0 12px 30px rgba(15,23,42,.06); }
    .kpi { padding:14px; }
    .kpi span,.metric span,.whatif-head span { color:var(--muted); font-size:12px; text-transform:uppercase; font-weight:800; }
    .kpi strong,.metric strong,.whatif-head strong { display:block; margin-top:7px; font-size:22px; }
    .chart-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:16px; }
    .panel { padding:16px; min-width:0; }
    summary { font-weight:800; cursor:pointer; }
    .table-wrap { overflow:auto; max-width:100%; }
    table { width:100%; border-collapse:separate; border-spacing:0; font-size:14px; }
    th,td { padding:10px 12px; border-bottom:1px solid var(--line); text-align:right; white-space:nowrap; }
    th:first-child,td:first-child { text-align:left; position:sticky; left:0; background:#fff; z-index:1; }
    th { color:var(--muted); font-size:12px; text-transform:uppercase; }
    .whatif-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }
    .warnings { background:#fff7ed; border:1px solid #fed7aa; border-radius:12px; padding:12px 16px; color:#9a3412; }
    .memo li { margin:8px 0; line-height:1.45; }
    .export { display:flex; gap:8px; }
    @media (max-width: 980px) { .sticky-kpis,.chart-grid,.whatif-grid { grid-template-columns:1fr 1fr; } }
    @media (max-width: 640px) { header,main { padding-left:14px; padding-right:14px; } .sticky-kpis,.chart-grid,.whatif-grid { grid-template-columns:1fr; position:static; } }
    @media print { header { position:static; } button,.tabs,.export { display:none; } body { background:#fff; } .panel,.kpi { box-shadow:none; break-inside:avoid; } .tab-panel { display:grid !important; } }
  </style>
</head>
<body>
  <header>
    <div class="hero">
      <div>
        <h1>Financial Dashboard</h1>
        <div class="subtitle">ОПУ analytics: M1, M3, M1 + M3 and WHAT IF</div>
      </div>
      <div class="export"><button type="button" onclick="window.print()">Export to PDF</button></div>
    </div>
    <nav class="tabs">
      <button class="active" data-tab="m1">M1</button>
      <button data-tab="m3">M3</button>
      <button data-tab="both">M1 + M3</button>
      <button data-tab="whatif">WHAT IF</button>
    </nav>
  </header>
  <main>
    __WARNINGS__
    __CONTENT__
  </main>
  <script>
    const REPORT_DATA = __PAYLOAD__;
    const rub = value => `${Math.round(value).toLocaleString("ru-RU")}`;
    const pct = value => `${(value * 100).toLocaleString("ru-RU", { maximumFractionDigits: 1 })}%`;

    document.querySelectorAll("[data-tab]").forEach(button => {
      button.addEventListener("click", () => {
        document.querySelectorAll("[data-tab]").forEach(item => item.classList.toggle("active", item === button));
        document.querySelectorAll(".tab-panel").forEach(panel => panel.classList.remove("active"));
        document.querySelector(`#tab-${button.dataset.tab}`).classList.add("active");
      });
    });

    function buildWhatIfTable() {
      const periods = REPORT_DATA.consolidated.map(row => row.period);
      const table = document.querySelector("#whatIfTable");
      table.innerHTML = `
        <thead><tr><th>Expense article</th>${periods.map(period => `<th>${period}</th>`).join("")}<th>Total</th></tr></thead>
        <tbody>${REPORT_DATA.whatIfArticles.map(article => {
          const values = REPORT_DATA.consolidated.map(row => row.what_if[article] || 0);
          const total = values.reduce((sum, value) => sum + value, 0);
          return `<tr><td><label><input type="checkbox" data-article="${article}"> ${article}</label></td>${values.map(value => `<td>${rub(value)}</td>`).join("")}<td>${rub(total)}</td></tr>`;
        }).join("")}</tbody>
      `;
      table.addEventListener("change", updateWhatIf);
      updateWhatIf();
    }

    function updateWhatIf() {
      const selected = Array.from(document.querySelectorAll("#whatIfTable input:checked")).map(input => input.dataset.article);
      const currentNet = REPORT_DATA.consolidated.reduce((sum, row) => sum + row.net_profit, 0);
      const currentExpenses = REPORT_DATA.consolidated.reduce((sum, row) => sum + row.total_expenses, 0);
      const currentEbitda = REPORT_DATA.consolidated.reduce((sum, row) => sum + row.ebitda, 0);
      const savings = REPORT_DATA.consolidated.reduce((sum, row) => sum + selected.reduce((inner, article) => inner + (row.what_if[article] || 0), 0), 0);
      const potentialNet = currentNet + savings;
      const potentialExpenses = currentExpenses - savings;
      const potentialEbitda = currentEbitda + savings;
      document.querySelector("#potentialNetProfit").textContent = rub(potentialNet);
      document.querySelector("#profitIncrease").textContent = rub(savings);
      document.querySelector("#profitIncreasePct").textContent = pct(currentNet ? savings / Math.abs(currentNet) : 0);
      document.querySelector("#savingsSummary").textContent = `${selected.length} articles / EBITDA after: ${rub(potentialEbitda)} / expenses after: ${rub(potentialExpenses)}`;
      if (window.Plotly) {
        Plotly.react("whatIfChart", [
          { x: ["Before", "After"], y: [currentNet, potentialNet], type: "bar", name: "Net Profit", marker: { color: ["#2563eb", "#059669"] } }
        ], { title: "Before / After Net Profit", template: "plotly_white", height: 360, margin: { l: 42, r: 20, t: 58, b: 42 } }, { displaylogo: false, responsive: true });
      }
    }

    buildWhatIfTable();
  </script>
</body>
</html>"""
    return (
        template.replace("__WARNINGS__", warning_html)
        .replace("__CONTENT__", content)
        .replace("__PAYLOAD__", json.dumps(payload, ensure_ascii=False))
    )


def build_report() -> str:
    """Build the complete report HTML."""

    m1 = parse_location(DATA_DIR / "m1.xlsx", "m1", "M1")
    m3 = parse_location(DATA_DIR / "m3.xlsx", "m3", "M3")
    locations = [m1, m3]
    consolidated = consolidate(locations)

    content_parts: list[str] = []
    first_plotly = True
    for loc in locations:
        tab_html, first_plotly = render_location_tab(
            loc.code,
            loc.rows,
            include_plotlyjs=first_plotly,
        )
        active_class = " active" if loc.code == "m1" else ""
        tab_html = tab_html.replace(
            '<section class="tab-panel"',
            f'<section class="tab-panel{active_class}"',
            1,
        )
        content_parts.append(tab_html)

    consolidated_html, first_plotly = render_consolidated_tab(
        locations,
        consolidated,
        include_plotlyjs=first_plotly,
    )
    content_parts.append(consolidated_html)
    content_parts.append(render_what_if(consolidated))
    content_parts.append(render_memo(locations, consolidated))

    warnings = [warning for loc in locations for warning in loc.warnings]
    return html_template("\n".join(content_parts), build_payload(locations, consolidated), warnings)


def main() -> None:
    """CLI entrypoint."""

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    html_report = build_report()
    OUTPUT_FILE.write_text(html_report, encoding="utf-8")
    logging.info("Report generated: %s", OUTPUT_FILE)


if __name__ == "__main__":
    main()
