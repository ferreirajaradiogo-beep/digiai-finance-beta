from __future__ import annotations

from collections import Counter
from datetime import date


def build_future_ai_bridge_payload(snapshot, provider="manus"):
    return {
        "provider": provider,
        "ready": True,
        "connected": False,
        "message": "Camada pronta para integrar IA externa sem alterar a interface atual.",
        "snapshot": snapshot,
    }


def _safe_amount(value):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _build_status(income_total, expense_total, balance_total):
    if income_total <= 0 and expense_total <= 0:
        return {"tone": "neutral", "label": "Sem dados"}
    if balance_total < 0:
        return {"tone": "danger", "label": "Atenção máxima"}
    if income_total > 0 and expense_total >= income_total * 0.9:
        return {"tone": "warning", "label": "Quase no limite"}
    if income_total > 0 and balance_total >= income_total * 0.2:
        return {"tone": "positive", "label": "Mês saudável"}
    return {"tone": "info", "label": "Sob controle"}


def _build_alert_and_tip(snapshot):
    income_total = snapshot["totals"]["income"]
    expense_total = snapshot["totals"]["expense"]
    balance_total = snapshot["totals"]["balance"]
    top_category = snapshot["top_expense_category"]
    share_pct = top_category["share_pct"]

    if snapshot["entries_count"] == 0:
        return (
            "Ainda não há movimentações suficientes para gerar uma leitura do seu mês.",
            "Comece registrando receitas e despesas para o assistente montar recomendações reais.",
            "Seu painel está pronto para acompanhar o mês assim que os primeiros lançamentos entrarem.",
        )

    if balance_total < 0:
        category_name = top_category["name"].lower()
        return (
            f"Seu mês está no vermelho. As despesas passaram as receitas em {abs(balance_total):.2f}.",
            f"Revise {category_name} primeiro e tente cortar pelo menos 10% desse grupo ainda esta semana.",
            "O foco imediato é reduzir a principal fonte de gasto e evitar novos lançamentos não essenciais.",
        )

    if income_total > 0 and expense_total >= income_total * 0.9:
        return (
            "Seu orçamento está muito apertado neste período e quase todo o valor recebido já foi consumido.",
            "Evite compras por impulso pelos próximos dias e priorize apenas despesas fixas e essenciais.",
            "Você ainda está positivo, mas a margem do mês ficou curta e merece atenção.",
        )

    if share_pct >= 45:
        return (
            f"A categoria {top_category['name']} concentrou a maior parte das suas despesas no período.",
            f"Defina um teto para {top_category['name']} no próximo ciclo para distribuir melhor os gastos.",
            "Existe concentração de despesa em um único grupo, o que é uma boa oportunidade de ajuste.",
        )

    if income_total > 0 and balance_total >= income_total * 0.2:
        return (
            "Seu mês está financeiramente saudável, com sobra relevante depois dos gastos.",
            "Separe parte desse saldo para reserva ou meta específica antes que ele vire gasto invisível.",
            "Você fechou o período com boa folga e espaço para poupar.",
        )

    return (
        "Seu mês está equilibrado, mas ainda vale acompanhar o ritmo de gastos para manter o saldo positivo.",
        "Uma boa estratégia é revisar semanalmente os lançamentos e limitar as despesas variáveis.",
        "A leitura geral é estável, com espaço para pequenas otimizações de rotina.",
    )


def build_financial_assistant_report(
    transactions,
    settings,
    period_label,
    start_date,
    end_date,
    provider="local",
):
    income_total = 0.0
    expense_total = 0.0
    category_totals = Counter()
    recent_expenses = []

    for item in transactions:
        amount = _safe_amount(item["valor"])
        if item["tipo"] == "receita":
            income_total += amount
        else:
            expense_total += amount
            category_name = item["categoria_nome"] or "Sem categoria"
            category_totals[category_name] += amount
            recent_expenses.append(
                {
                    "description": item["descricao"],
                    "category": category_name,
                    "value": amount,
                    "date": item["data"] or "-",
                }
            )

    balance_total = income_total - expense_total
    top_name = "Sem despesas"
    top_value = 0.0
    top_share = 0.0
    if category_totals:
        top_name, top_value = category_totals.most_common(1)[0]
        top_share = round((top_value / expense_total) * 100, 2) if expense_total else 0.0

    snapshot = {
        "period": {
            "label": period_label,
            "start_date": start_date,
            "end_date": end_date,
            "generated_at": date.today().isoformat(),
        },
        "entries_count": len(transactions),
        "totals": {
            "income": round(income_total, 2),
            "expense": round(expense_total, 2),
            "balance": round(balance_total, 2),
        },
        "top_expense_category": {
            "name": top_name,
            "value": round(top_value, 2),
            "share_pct": top_share,
        },
        "top_categories": [
            {"name": name, "value": round(value, 2)}
            for name, value in category_totals.most_common(4)
        ],
        "recent_expenses": recent_expenses[:5],
        "currency": settings.get("moeda", "BRL"),
        "status": _build_status(income_total, expense_total, balance_total),
    }

    financial_alert, saving_tip, summary = _build_alert_and_tip(snapshot)
    future_bridge = build_future_ai_bridge_payload(snapshot, provider="manus")

    return {
        "provider": provider,
        "snapshot": snapshot,
        "financial_alert": financial_alert,
        "saving_tip": saving_tip,
        "summary": summary,
        "future_bridge": future_bridge,
    }
