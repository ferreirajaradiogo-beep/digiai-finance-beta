import json
import os
from csv import writer as csv_writer
from collections import Counter
from datetime import date, timedelta
from functools import wraps
from io import BytesIO, StringIO
from pathlib import Path
from urllib import error as urllib_error
from urllib import parse as urllib_parse
from urllib import request as urllib_request

from flask import (
    Flask,
    flash,
    g,
    make_response,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from werkzeug.security import check_password_hash, generate_password_hash

from db_compat import DB_BACKEND, DB_INTEGRITY_ERRORS, connect, get_table_columns
from services import build_financial_assistant_report

app = Flask(
    __name__,
    template_folder="templates_v3",
    static_folder="static_v3",
    static_url_path="/static_v3",
)
app.config.update(
    SECRET_KEY=os.environ.get("FINANCEIRO_SECRET_KEY", "segredo"),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("NOTAFACIL_COOKIE_SECURE", "false").lower() == "true",
    PREFERRED_URL_SCHEME=os.environ.get("NOTAFACIL_URL_SCHEME", "https"),
)

BASE_DIR = Path(__file__).resolve().parent
APP_DISPLAY_NAME = os.environ.get("NOTAFACIL_APP_NAME", "DiGiaI Caixa")
APP_TAGLINE = os.environ.get("NOTAFACIL_APP_TAGLINE", "Controle financeiro pessoal inteligente")
COMPANY_NAME = os.environ.get("NOTAFACIL_COMPANY_NAME", "DigiAI")
SUPPORT_EMAIL = os.environ.get("NOTAFACIL_SUPPORT_EMAIL", "digiai.oficial@gmail.com")
PUBLIC_BETA_MODE = True
DB_PATH = Path(os.environ.get("NOTAFACIL_DB_PATH", str(BASE_DIR / "financeiro_beta.db"))).expanduser()
BACKUP_DIR = Path(os.environ.get("NOTAFACIL_BACKUP_DIR", str(BASE_DIR / "backups"))).expanduser()
BACKEND_API_URL = (os.environ.get("NOTAFACIL_BACKEND_URL", "") or "").strip().rstrip("/")

DEFAULT_SETTINGS = {"idioma": "pt-BR", "moeda": "BRL", "tema": "black-ocean", "plano": "free"}
FREE_ALLOWED_COLORS = ["#58d5ff", "#7cf0bb", "#ffab7a", "#ff7c93"]
PAID_COLOR_PRESETS = [
    "#58d5ff",
    "#7cf0bb",
    "#ffab7a",
    "#ff7c93",
    "#b794f4",
    "#f6ad55",
    "#4fd1c5",
    "#f687b3",
    "#90cdf4",
    "#68d391",
    "#f56565",
    "#fbd38d",
    "#63b3ed",
    "#a3bffa",
    "#fbb6ce",
    "#c6f6d5",
]
PLANS = {
    "free": {
        "name": "Gratuito",
        "transaction_limit_per_month": 20,
        "account_limit": 1,
        "category_limit": 10,
        "allow_multi_currency": False,
        "allow_currency_conversion": False,
        "allow_advanced_reports": False,
        "allow_alerts": False,
        "allow_unlimited_colors": False,
    },
    "pro": {
        "name": "Pago",
        "transaction_limit_per_month": None,
        "account_limit": None,
        "category_limit": None,
        "allow_multi_currency": True,
        "allow_currency_conversion": True,
        "allow_advanced_reports": True,
        "allow_alerts": True,
        "allow_unlimited_colors": True,
    },
}
CURRENCY_SYMBOLS = {"BRL": "R$", "USD": "$", "EUR": "EUR", "GBP": "GBP"}
LANGUAGE_OPTIONS = [("pt-BR", "Português"), ("en-US", "English"), ("es-ES", "Español")]
CURRENCY_OPTIONS = [
    ("BRL", "Real brasileiro (BRL)"),
    ("USD", "US Dollar (USD)"),
    ("EUR", "Euro (EUR)"),
    ("GBP", "Pound Sterling (GBP)"),
]
BACKGROUND_THEME_OPTIONS = [
    ("black", "Preto"),
    ("white", "Branco"),
]
FREE_ACCENT_THEME_OPTIONS = [
    ("ocean", "Ocean"),
]
PAID_ACCENT_THEME_OPTIONS = FREE_ACCENT_THEME_OPTIONS + [
    ("graphite", "Graphite"),
    ("ember", "Ember"),
    ("royal", "Royal"),
    ("sand", "Sand"),
    ("mint", "Mint"),
]

ACCENT_THEME_META = {
    "ocean": {"label": "Ocean", "premium": False},
    "graphite": {"label": "Graphite", "premium": True},
    "ember": {"label": "Ember", "premium": True},
    "royal": {"label": "Royal", "premium": True},
    "sand": {"label": "Sand", "premium": True},
    "mint": {"label": "Mint", "premium": True},
}


def compose_theme_key(background, accent):
    background_key = str(background or "black").strip().lower()
    accent_key = str(accent or "ocean").strip().lower()
    allowed_backgrounds = {key for key, _ in BACKGROUND_THEME_OPTIONS}
    allowed_accents = set(ACCENT_THEME_META)
    if background_key not in allowed_backgrounds:
        background_key = "black"
    if accent_key not in allowed_accents:
        accent_key = "ocean"
    return f"{background_key}-{accent_key}"


def split_theme_key(value):
    normalized = str(value or "").strip().lower()
    legacy_map = {
        "dark": ("black", "ocean"),
        "light": ("white", "ocean"),
        "black": ("black", "ocean"),
        "white": ("white", "ocean"),
        "ocean": ("black", "ocean"),
        "graphite": ("black", "graphite"),
        "forest": ("black", "mint"),
        "sunset": ("black", "ember"),
        "aurora": ("black", "royal"),
        "ember": ("black", "ember"),
        "royal": ("black", "royal"),
        "sand": ("black", "sand"),
        "mint": ("black", "mint"),
    }
    if normalized in legacy_map:
        return legacy_map[normalized]
    if "-" in normalized:
        background_key, accent_key = normalized.split("-", 1)
        return (
            background_key if background_key in {key for key, _ in BACKGROUND_THEME_OPTIONS} else "black",
            accent_key if accent_key in ACCENT_THEME_META else "ocean",
        )
    return ("black", "ocean")


def normalize_theme_key(value, plan_key="free"):
    background_key, accent_key = split_theme_key(value)
    return compose_theme_key(background_key, accent_key)


def get_background_theme_class(theme_key):
    background_key, _ = split_theme_key(theme_key)
    return "light" if background_key == "white" else "dark"


def get_accent_theme_class(theme_key):
    _, accent_key = split_theme_key(theme_key)
    return accent_key

TRANSLATIONS = {
    "pt-BR": {
        "app_title": "DiGiaI Caixa",
        "dashboard": "Dashboard",
        "settings": "Configurações",
        "management": "Categorias e contas",
        "plans": "Planos",
        "logout": "Sair",
        "welcome": "Olá",
        "summary": "Controle seus lançamentos com filtro, edição e instalação no celular.",
        "month": "Mês",
        "year": "Ano",
        "all_months": "Todos",
        "all_years": "Todos",
        "apply_filters": "Filtrar",
        "clear_filters": "Limpar",
        "income": "Receita",
        "expense": "Despesa",
        "balance": "Saldo",
        "transactions": "Transações",
        "none_transactions": "Nenhuma transação encontrada para esse filtro.",
        "description": "Descrição",
        "category": "Categoria",
        "account": "Conta",
        "categories": "Categorias",
        "accounts": "Contas",
        "new_category": "Nova categoria",
        "new_account": "Nova conta",
        "manage_help": "Organize seus lançamentos com categorias e contas próprias.",
        "name": "Nome",
        "color": "Cor",
        "account_type": "Tipo de conta",
        "create": "Criar",
        "save_changes": "Salvar alterações",
        "delete_item": "Excluir",
        "wallet": "Carteira",
        "bank": "Banco",
        "credit_card": "Cartão",
        "savings": "Poupança",
        "category_created": "Categoria criada com sucesso.",
        "account_created_item": "Conta criada com sucesso.",
        "category_updated": "Categoria atualizada com sucesso.",
        "account_updated": "Conta atualizada com sucesso.",
        "category_deleted": "Categoria excluída com sucesso.",
        "account_deleted": "Conta excluída com sucesso.",
        "category_exists": "Essa categoria já existe para este usuário.",
        "account_exists": "Essa conta já existe para este usuário.",
        "field_required": "Preencha o nome antes de salvar.",
        "free_plan": "Gratuito",
        "paid_plan": "Pago",
        "current_plan": "Plano atual",
        "upgrade_plan": "Ativar plano pago",
        "free_plan_badge": "Até 20 lançamentos por mês",
        "paid_plan_badge": "Recursos avançados e ilimitados",
        "limit_transactions_reached": "Você atingiu o limite de 20 lançamentos neste mês no plano gratuito.",
        "limit_categories_reached": "Você atingiu o limite de 10 categorias no plano gratuito.",
        "limit_accounts_reached": "O plano gratuito permite apenas 1 conta.",
        "free_account_hint": "No plano gratuito, você pode usar 1 conta.",
        "limited_color_message": "No plano gratuito, escolha uma cor da paleta liberada.",
        "premium_feature_message": "Esse recurso faz parte do plano pago.",
        "premium_themes_help": "Temas extras do sistema ficam disponíveis no plano pago.",
        "export_excel": "Exportar Excel",
        "export_pdf": "Exportar PDF",
        "download_backup": "Baixar backup",
        "backup_ready": "Backup gerado com sucesso.",
        "reports_title": "Relatórios",
        "reports_subtitle": "Os relatórios avançados ficam disponíveis no plano pago.",
        "advanced_reports_locked": "Gráficos por mês, gastos por categoria e comparação mensal são recursos premium.",
        "alerts_locked": "Alertas e notificações de gasto excessivo são recursos premium.",
        "alerts_title": "Alertas",
        "alert_threshold": "Limite mensal de alerta",
        "alert_threshold_help": "Receba um aviso quando suas despesas do período ultrapassarem esse valor.",
        "overspending_alert": "Alerta: suas despesas passaram do limite configurado.",
        "paid_colors_title": "Cores premium",
        "paid_colors_help": "O plano pago libera uma paleta maior para personalizar categorias.",
        "monthly_chart": "Gráfico mensal",
        "category_chart": "Gastos por categoria",
        "monthly_comparison": "Comparação mensal",
        "no_report_data": "Ainda não há dados suficientes para gerar este relatório.",
        "date": "Data",
        "amount": "Valor",
        "action": "Ação",
        "edit": "Editar",
        "delete": "Excluir",
        "new_transaction": "Novo lançamento",
        "edit_transaction": "Editar lançamento",
        "type": "Tipo",
        "save": "Salvar",
        "update": "Atualizar",
        "preferences": "Preferências",
        "preferences_help": "Escolha idioma, moeda e tema visual para sua conta.",
        "language": "Idioma",
        "currency": "Moeda",
        "theme": "Cor de fundo",
        "login_title": "Entrar",
        "login_help": "Entre com seu usuário e senha para acessar seu painel financeiro.",
        "register_title": "Criar conta",
        "register_help": "Cadastre um novo usuário para acessar o sistema com seus próprios dados.",
        "username": "Usuário",
        "password": "Senha",
        "sign_in": "Entrar",
        "register": "Cadastrar",
        "no_account": "Ainda não tem conta?",
        "has_account": "Já tem conta?",
        "create_account": "Criar conta",
        "settings_saved": "Configurações salvas com sucesso.",
        "account_created": "Conta criada com sucesso.",
        "transaction_created": "Lançamento adicionado com sucesso.",
        "transaction_updated": "Lançamento atualizado com sucesso.",
        "transaction_deleted": "Lançamento excluído com sucesso.",
        "invalid_amount": "Informe um valor válido.",
        "missing_credentials": "Preencha usuário e senha.",
        "user_exists": "Esse usuário já existe. Faça login ou escolha outro nome.",
        "user_not_found": "Usuário não encontrado. Crie uma conta para começar.",
        "password_too_short": "Use uma senha com pelo menos 4 caracteres.",
        "username_too_short": "Use um usuário com pelo menos 3 caracteres.",
        "wrong_password": "Senha incorreta.",
        "not_found": "Lançamento não encontrado.",
        "delete_confirm": "Tem certeza que deseja excluir este lançamento?",
        "install_app": "Instalar app",
        "period_summary": "Resumo do período",
        "entries_count": "Lançamentos",
        "top_category": "Categoria em destaque",
        "latest_date": "Data mais recente",
        "active_currency": "Moeda ativa",
        "date_from": "Data inicial",
        "date_to": "Data final",
        "offline_title": "Sem conexão",
        "offline_text": "Você está offline agora. Quando a conexão voltar, atualize a página para sincronizar.",
        "back_dashboard": "Voltar ao painel",
        "total_registered": "Total de lançamentos",
        "assistant_nav": "Assistente",
        "assistant_title": "Assistente Financeiro",
        "assistant_subtitle": "Leitura automatica do periodo com suporte operacional para sync, plano, senha e gasto mensal.",
        "assistant_cta": "Abrir assistente",
        "assistant_refresh": "Atualizar análise",
        "assistant_generated": "Análise gerada com base nos lançamentos do período selecionado.",
        "assistant_engine_local": "Motor local ativo",
        "assistant_engine_ready": "Arquitetura pronta para Manus/API externa",
        "assistant_alert": "Alerta financeiro",
        "assistant_tip": "Dica de economia",
        "assistant_summary_title": "Resumo inteligente",
        "assistant_no_data_title": "Sem dados suficientes",
        "assistant_no_data_text": "Cadastre receitas e despesas para o assistente montar uma leitura financeira automática.",
        "assistant_open_api": "Ver endpoint JSON",
        "assistant_overview": "Visão automática do período",
        "assistant_current_month": "Mês atual",
        "assistant_chat_title": "Pergunte ao assistente",
        "assistant_chat_placeholder": "Ex.: como sincronizar meu app? se aparecer token invalido, o que faco? qual categoria mais pesa?",
        "assistant_chat_submit": "Perguntar",
        "assistant_chat_hint": "Posso ajudar com sincronizacao, token invalido, plano, senha e leitura do seu mes.",
        "assistant_answer_title": "Resposta",
        "assistant_provider_local": "Resposta local",
        "assistant_provider_openai": "Resposta com IA",
        "assistant_provider_reason_missing_api_key": "A chave ASSISTANT_API_KEY ainda nao foi lida pelo backend publicado.",
        "assistant_provider_reason_local_fallback": "A resposta caiu no modo local de seguranca.",
        "assistant_provider_reason_provider_http_error": "O provedor de IA recusou a chamada ou ficou sem cota. Revise a chave e o limite do provedor.",
        "assistant_provider_reason_provider_request_failed": "Nao consegui falar com o provedor de IA agora. Vou responder em modo local por seguranca.",
        "assistant_provider_reason_provider_empty_output": "O provedor de IA respondeu sem texto util. Vou manter a resposta local por seguranca.",
        "assistant_provider_reason_unsupported_provider": "O provedor configurado para o assistente nao e suportado por esta versao do backend.",
        "assistant_provider_reason_local_provider_selected": "O assistente esta configurado para operar apenas em modo local.",
    },
    "en-US": {
        "app_title": "DiGiaI Caixa",
        "dashboard": "Dashboard",
        "settings": "Settings",
        "management": "Categories and accounts",
        "plans": "Plans",
        "logout": "Sign out",
        "welcome": "Hello",
        "summary": "Track your entries with filters, editing and installable mobile access.",
        "month": "Month",
        "year": "Year",
        "all_months": "All",
        "all_years": "All",
        "apply_filters": "Apply",
        "clear_filters": "Clear",
        "income": "Income",
        "expense": "Expense",
        "balance": "Balance",
        "transactions": "Transactions",
        "none_transactions": "No transactions found for this filter.",
        "description": "Description",
        "category": "Category",
        "account": "Account",
        "categories": "Categories",
        "accounts": "Accounts",
        "new_category": "New category",
        "new_account": "New account",
        "manage_help": "Organize your entries with your own categories and accounts.",
        "name": "Name",
        "color": "Color",
        "account_type": "Account type",
        "create": "Create",
        "save_changes": "Save changes",
        "delete_item": "Delete",
        "wallet": "Wallet",
        "bank": "Bank",
        "credit_card": "Card",
        "savings": "Savings",
        "category_created": "Category created successfully.",
        "account_created_item": "Account created successfully.",
        "category_updated": "Category updated successfully.",
        "account_updated": "Account updated successfully.",
        "category_deleted": "Category deleted successfully.",
        "account_deleted": "Account deleted successfully.",
        "category_exists": "This category already exists for this user.",
        "account_exists": "This account already exists for this user.",
        "field_required": "Fill in the name before saving.",
        "free_plan": "Free",
        "paid_plan": "Paid",
        "current_plan": "Current plan",
        "upgrade_plan": "Activate paid plan",
        "free_plan_badge": "Up to 20 entries per month",
        "paid_plan_badge": "Advanced and unlimited features",
        "limit_transactions_reached": "You reached the 20 entries monthly limit on the free plan.",
        "limit_categories_reached": "You reached the 10 categories limit on the free plan.",
        "limit_accounts_reached": "The free plan allows only 1 account.",
        "free_account_hint": "On the free plan, you can use 1 account.",
        "limited_color_message": "On the free plan, choose a color from the allowed palette.",
        "premium_feature_message": "This feature is part of the paid plan.",
        "premium_themes_help": "Extra app themes are available on the paid plan.",
        "export_excel": "Export Excel",
        "export_pdf": "Export PDF",
        "download_backup": "Download backup",
        "backup_ready": "Backup generated successfully.",
        "reports_title": "Reports",
        "reports_subtitle": "Advanced reports are available on the paid plan.",
        "advanced_reports_locked": "Monthly charts, category spending and monthly comparison are premium features.",
        "alerts_locked": "Alerts and overspending notifications are premium features.",
        "alerts_title": "Alerts",
        "alert_threshold": "Monthly alert threshold",
        "alert_threshold_help": "Get a warning when your period expenses exceed this amount.",
        "overspending_alert": "Alert: your expenses exceeded the configured threshold.",
        "paid_colors_title": "Premium colors",
        "paid_colors_help": "The paid plan unlocks a wider palette to customize categories.",
        "monthly_chart": "Monthly chart",
        "category_chart": "Spending by category",
        "monthly_comparison": "Monthly comparison",
        "no_report_data": "There is not enough data to generate this report yet.",
        "date": "Date",
        "amount": "Amount",
        "action": "Action",
        "edit": "Edit",
        "delete": "Delete",
        "new_transaction": "New entry",
        "edit_transaction": "Edit entry",
        "type": "Type",
        "save": "Save",
        "update": "Update",
        "preferences": "Preferences",
        "preferences_help": "Choose language, currency and visual theme for your account.",
        "language": "Language",
        "currency": "Currency",
        "theme": "Background color",
        "login_title": "Sign in",
        "login_help": "Sign in with your username and password to access your financial dashboard.",
        "register_title": "Create account",
        "register_help": "Create a new user account to access the system with separate data.",
        "username": "Username",
        "password": "Password",
        "sign_in": "Sign in",
        "register": "Register",
        "no_account": "Don't have an account yet?",
        "has_account": "Already have an account?",
        "create_account": "Create account",
        "settings_saved": "Settings saved successfully.",
        "account_created": "Account created successfully.",
        "transaction_created": "Entry added successfully.",
        "transaction_updated": "Entry updated successfully.",
        "transaction_deleted": "Entry deleted successfully.",
        "invalid_amount": "Enter a valid amount.",
        "missing_credentials": "Enter username and password.",
        "user_exists": "This username already exists. Sign in or choose another one.",
        "user_not_found": "User not found. Create an account to get started.",
        "password_too_short": "Use a password with at least 4 characters.",
        "username_too_short": "Use a username with at least 3 characters.",
        "wrong_password": "Incorrect password.",
        "not_found": "Entry not found.",
        "delete_confirm": "Are you sure you want to delete this entry?",
        "install_app": "Install app",
        "period_summary": "Period summary",
        "entries_count": "Entries",
        "top_category": "Top category",
        "latest_date": "Latest date",
        "active_currency": "Active currency",
        "date_from": "Start date",
        "date_to": "End date",
        "offline_title": "You are offline",
        "offline_text": "You are offline right now. Refresh when the connection returns.",
        "back_dashboard": "Back to dashboard",
        "total_registered": "Total entries",
        "assistant_nav": "Assistant",
        "assistant_title": "Financial Assistant",
        "assistant_subtitle": "An automatic reading of your month with alerts, savings tips and architecture ready for future AI.",
        "assistant_cta": "Open assistant",
        "assistant_refresh": "Refresh analysis",
        "assistant_generated": "Analysis generated from the entries in the selected period.",
        "assistant_engine_local": "Local engine active",
        "assistant_engine_ready": "Architecture ready for Manus/external API",
        "assistant_alert": "Financial alert",
        "assistant_tip": "Savings tip",
        "assistant_summary_title": "Smart summary",
        "assistant_no_data_title": "Not enough data",
        "assistant_no_data_text": "Add income and expenses so the assistant can build an automatic financial reading.",
        "assistant_open_api": "Open JSON endpoint",
        "assistant_overview": "Automatic period overview",
        "assistant_current_month": "Current month",
        "assistant_chat_title": "Ask the assistant",
        "assistant_chat_placeholder": "Example: how do I sync my app? which category weighs the most? how do I change plans?",
        "assistant_chat_submit": "Ask",
        "assistant_chat_hint": "I can help with sync, plan, password and your monthly financial reading.",
        "assistant_answer_title": "Answer",
        "assistant_provider_local": "Local answer",
        "assistant_provider_openai": "AI answer",
        "assistant_provider_reason_missing_api_key": "The published backend has not picked up ASSISTANT_API_KEY yet.",
        "assistant_provider_reason_local_fallback": "The answer fell back to the local safe mode.",
        "assistant_provider_reason_provider_http_error": "The AI provider rejected the request or ran out of quota. Check the provider key and limits.",
        "assistant_provider_reason_provider_request_failed": "I could not reach the AI provider right now, so I answered in the local safe mode.",
        "assistant_provider_reason_provider_empty_output": "The AI provider returned no useful text, so I kept the local safe answer.",
        "assistant_provider_reason_unsupported_provider": "The configured assistant provider is not supported by this backend version.",
        "assistant_provider_reason_local_provider_selected": "The assistant is configured to run only in local mode.",
    },
    "es-ES": {
        "app_title": "DiGiaI Caixa",
        "dashboard": "Panel",
        "settings": "Configuración",
        "management": "Categorías y cuentas",
        "plans": "Planes",
        "logout": "Salir",
        "welcome": "Hola",
        "summary": "Controla tus registros con filtros, edición e instalación en el móvil.",
        "month": "Mes",
        "year": "Año",
        "all_months": "Todos",
        "all_years": "Todos",
        "apply_filters": "Filtrar",
        "clear_filters": "Limpiar",
        "income": "Ingreso",
        "expense": "Gasto",
        "balance": "Saldo",
        "transactions": "Movimientos",
        "none_transactions": "No hay movimientos para este filtro.",
        "description": "Descripción",
        "category": "Categoría",
        "account": "Cuenta",
        "categories": "Categorías",
        "accounts": "Cuentas",
        "new_category": "Nueva categoría",
        "new_account": "Nueva cuenta",
        "manage_help": "Organiza tus registros con categorías y cuentas propias.",
        "name": "Nombre",
        "color": "Color",
        "account_type": "Tipo de cuenta",
        "create": "Crear",
        "save_changes": "Guardar cambios",
        "delete_item": "Eliminar",
        "wallet": "Billetera",
        "bank": "Banco",
        "credit_card": "Tarjeta",
        "savings": "Ahorros",
        "category_created": "Categoría creada con éxito.",
        "account_created_item": "Cuenta creada con éxito.",
        "category_updated": "Categoría actualizada con éxito.",
        "account_updated": "Cuenta actualizada con éxito.",
        "category_deleted": "Categoría eliminada con éxito.",
        "account_deleted": "Cuenta eliminada con éxito.",
        "category_exists": "Esa categoría ya existe para este usuario.",
        "account_exists": "Esa cuenta ya existe para este usuario.",
        "field_required": "Completa el nombre antes de guardar.",
        "free_plan": "Gratis",
        "paid_plan": "Pago",
        "current_plan": "Plan actual",
        "upgrade_plan": "Activar plan pago",
        "free_plan_badge": "Hasta 20 registros por mes",
        "paid_plan_badge": "Funciones avanzadas e ilimitadas",
        "limit_transactions_reached": "Alcanzaste el límite de 20 registros este mes en el plan gratis.",
        "limit_categories_reached": "Alcanzaste el límite de 10 categorías en el plan gratis.",
        "limit_accounts_reached": "El plan gratis permite solo 1 cuenta.",
        "free_account_hint": "En el plan gratis, puedes usar 1 cuenta.",
        "limited_color_message": "En el plan gratis, elige un color de la paleta permitida.",
        "premium_feature_message": "Esta función forma parte del plan pago.",
        "premium_themes_help": "Los temas extra del sistema están disponibles en el plan pago.",
        "export_excel": "Exportar Excel",
        "export_pdf": "Exportar PDF",
        "download_backup": "Descargar backup",
        "backup_ready": "Backup generado con éxito.",
        "reports_title": "Reportes",
        "reports_subtitle": "Los reportes avanzados están disponibles en el plan pago.",
        "advanced_reports_locked": "Gráficos por mes, gastos por categoría y comparación mensual son funciones premium.",
        "alerts_locked": "Alertas y notificaciones de gasto excesivo son funciones premium.",
        "alerts_title": "Alertas",
        "alert_threshold": "Límite mensual de alerta",
        "alert_threshold_help": "Recibe un aviso cuando tus gastos del período superen ese valor.",
        "overspending_alert": "Alerta: tus gastos superaron el límite configurado.",
        "paid_colors_title": "Colores premium",
        "paid_colors_help": "El plan pago libera una paleta más amplia para personalizar categorías.",
        "monthly_chart": "Gráfico mensual",
        "category_chart": "Gastos por categoría",
        "monthly_comparison": "Comparación mensual",
        "no_report_data": "Todavía no hay datos suficientes para generar este reporte.",
        "date": "Fecha",
        "amount": "Valor",
        "action": "Acción",
        "edit": "Editar",
        "delete": "Eliminar",
        "new_transaction": "Nuevo registro",
        "edit_transaction": "Editar registro",
        "type": "Tipo",
        "save": "Guardar",
        "update": "Actualizar",
        "preferences": "Preferencias",
        "preferences_help": "Elige idioma, moneda y tema visual para tu cuenta.",
        "language": "Idioma",
        "currency": "Moneda",
        "theme": "Color de fondo",
        "login_title": "Ingresar",
        "login_help": "Ingresa con tu usuario y contraseña para acceder a tu panel financiero.",
        "register_title": "Crear cuenta",
        "register_help": "Crea un nuevo usuario para acceder al sistema con datos separados.",
        "username": "Usuario",
        "password": "Contraseña",
        "sign_in": "Ingresar",
        "register": "Registrar",
        "no_account": "¿Todavía no tienes cuenta?",
        "has_account": "¿Ya tienes cuenta?",
        "create_account": "Crear cuenta",
        "settings_saved": "Configuración guardada con éxito.",
        "account_created": "Cuenta creada con éxito.",
        "transaction_created": "Registro agregado con éxito.",
        "transaction_updated": "Registro actualizado con éxito.",
        "transaction_deleted": "Registro eliminado con éxito.",
        "invalid_amount": "Ingresa un valor válido.",
        "missing_credentials": "Completa usuario y contraseña.",
        "user_exists": "Ese usuario ya existe. Inicia sesión o elige otro nombre.",
        "user_not_found": "Usuario no encontrado. Crea una cuenta para comenzar.",
        "password_too_short": "Usa una contraseña de al menos 4 caracteres.",
        "username_too_short": "Usa un usuario de al menos 3 caracteres.",
        "wrong_password": "Contraseña incorrecta.",
        "not_found": "Registro no encontrado.",
        "delete_confirm": "¿Seguro que deseas eliminar este registro?",
        "install_app": "Instalar app",
        "period_summary": "Resumen del período",
        "entries_count": "Registros",
        "top_category": "Categoría destacada",
        "latest_date": "Fecha más reciente",
        "active_currency": "Moneda activa",
        "date_from": "Fecha inicial",
        "date_to": "Fecha final",
        "offline_title": "Sin conexión",
        "offline_text": "Ahora estás sin conexión. Cuando vuelva internet, actualiza la página.",
        "back_dashboard": "Volver al panel",
        "total_registered": "Total de registros",
        "assistant_nav": "Asistente",
        "assistant_title": "Asistente Financiero",
        "assistant_subtitle": "Una lectura automática de tu mes, con alertas, ahorro y arquitectura lista para IA futura.",
        "assistant_cta": "Abrir asistente",
        "assistant_refresh": "Actualizar análisis",
        "assistant_generated": "Análisis generado con base en los registros del período seleccionado.",
        "assistant_engine_local": "Motor local activo",
        "assistant_engine_ready": "Arquitectura lista para Manus/API externa",
        "assistant_alert": "Alerta financiera",
        "assistant_tip": "Consejo de ahorro",
        "assistant_summary_title": "Resumen inteligente",
        "assistant_no_data_title": "Sin datos suficientes",
        "assistant_no_data_text": "Registra ingresos y gastos para que el asistente genere una lectura automática.",
        "assistant_open_api": "Ver endpoint JSON",
        "assistant_overview": "Visión automática del período",
        "assistant_current_month": "Mes actual",
        "assistant_chat_title": "Pregunta al asistente",
        "assistant_chat_placeholder": "Ej.: como sincronizar mi app? que categoria pesa mas? como cambiar de plan?",
        "assistant_chat_submit": "Preguntar",
        "assistant_chat_hint": "Puedo ayudar con sincronización, plan, contraseña y lectura de tu mes.",
        "assistant_answer_title": "Respuesta",
        "assistant_provider_local": "Respuesta local",
        "assistant_provider_openai": "Respuesta con IA",
        "assistant_provider_reason_missing_api_key": "El backend publicado aun no tomo la ASSISTANT_API_KEY.",
        "assistant_provider_reason_local_fallback": "La respuesta cayo al modo local de seguridad.",
        "assistant_provider_reason_provider_http_error": "El proveedor de IA rechazo la llamada o se quedo sin cuota. Revisa la clave y el limite del proveedor.",
        "assistant_provider_reason_provider_request_failed": "No pude hablar con el proveedor de IA ahora, asi que respondi en modo local de seguridad.",
        "assistant_provider_reason_provider_empty_output": "El proveedor de IA respondio sin texto util, asi que mantuve la respuesta local de seguridad.",
        "assistant_provider_reason_unsupported_provider": "El proveedor configurado para el asistente no es compatible con esta version del backend.",
        "assistant_provider_reason_local_provider_selected": "El asistente esta configurado para funcionar solo en modo local.",
    },
}

MONTH_LABELS = {
    "pt-BR": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
    "en-US": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    "es-ES": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"],
}


def backend_mode_enabled():
    return bool(BACKEND_API_URL)


def get_backend_token():
    return session.get("api_token", "")


def backend_user_from_session():
    user = session.get("backend_user")
    return user if isinstance(user, dict) else {}


def backend_snapshot_from_session():
    snapshot = session.get("backend_snapshot_cache")
    return snapshot if isinstance(snapshot, dict) else {}


class BackendAuthError(RuntimeError):
    pass


class BackendRateLimitError(RuntimeError):
    pass


def store_backend_snapshot(snapshot):
    if isinstance(snapshot, dict):
        session["backend_snapshot_cache"] = snapshot
        session.modified = True


def clear_backend_session(message="Sua sessao online expirou. Entre novamente para atualizar seus dados."):
    session.clear()
    flash(message, "warning")


def mark_backend_unavailable(message="O backend esta acordando. Tente novamente em instantes."):
    g.backend_unavailable = True
    if not getattr(g, "backend_status_message", ""):
        g.backend_status_message = message


def update_backend_session_user(payload):
    if not payload:
        return
    session["backend_user"] = dict(payload)
    session["user"] = payload.get("username", session.get("user", ""))
    session.permanent = True


def backend_api_request(path, method="GET", token="", body=None, query=None):
    if not backend_mode_enabled():
        raise RuntimeError("Backend API nao configurada")

    url = f"{BACKEND_API_URL}{path}"
    if query:
        clean_query = {key: value for key, value in query.items() if value not in (None, "", [])}
        if clean_query:
            url = f"{url}?{urllib_parse.urlencode(clean_query)}"

    headers = {"Accept": "application/json"}
    payload = None
    if body is not None:
        payload = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    if token:
        headers["Authorization"] = f"Bearer {token}"

    req = urllib_request.Request(url, data=payload, headers=headers, method=method.upper())
    try:
        with urllib_request.urlopen(req, timeout=15) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib_error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="ignore")
        detail = ""
        if raw:
            try:
                parsed = json.loads(raw)
                detail = parsed.get("detail") if isinstance(parsed, dict) else raw
            except json.JSONDecodeError:
                detail = raw
        if exc.code == 401 and token:
            raise BackendAuthError(detail or "Token invalido") from exc
        if exc.code == 429:
            raise BackendRateLimitError(
                "Muitas tentativas em pouco tempo. Aguarde cerca de 1 minuto e tente entrar novamente."
            ) from exc
        raise ValueError(detail or f"Erro HTTP {exc.code}") from exc
    except urllib_error.URLError as exc:
        raise ConnectionError("Nao foi possivel conectar ao backend.") from exc


def refresh_backend_me():
    if not (backend_mode_enabled() and get_backend_token()):
        return backend_user_from_session()
    try:
        payload = backend_api_request("/me", token=get_backend_token())
        update_backend_session_user(payload)
        return payload
    except ConnectionError:
        mark_backend_unavailable()
        return backend_user_from_session()
    except BackendAuthError:
        raise
    except ValueError as exc:
        mark_backend_unavailable(str(exc) or "Nao foi possivel atualizar a sessao online agora.")
        return backend_user_from_session()


def get_backend_snapshot(force=False):
    if not backend_mode_enabled():
        return {}
    if not get_backend_token():
        return backend_snapshot_from_session()
    if force or not hasattr(g, "backend_snapshot"):
        try:
            g.backend_snapshot = backend_api_request("/sync/pull", token=get_backend_token())
            store_backend_snapshot(g.backend_snapshot)
        except ConnectionError:
            mark_backend_unavailable()
            g.backend_snapshot = backend_snapshot_from_session()
        except BackendAuthError:
            raise
        except ValueError as exc:
            mark_backend_unavailable(str(exc) or "Nao foi possivel atualizar os dados online agora.")
            g.backend_snapshot = backend_snapshot_from_session()
    return g.backend_snapshot


def backend_categories_map():
    snapshot = get_backend_snapshot()
    categories = snapshot.get("categories") or []
    return {item["id"]: item for item in categories if isinstance(item, dict) and item.get("id") is not None}


def backend_accounts_map():
    snapshot = get_backend_snapshot()
    accounts = snapshot.get("accounts") or []
    return {item["id"]: item for item in accounts if isinstance(item, dict) and item.get("id") is not None}


def conectar():
    return connect(DB_PATH)


def table_columns(cursor, table_name):
    return get_table_columns(cursor, table_name)


def criar_tabelas():
    conn = conectar()
    c = conn.cursor()
    if DB_BACKEND == "postgres":
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                usuario TEXT UNIQUE,
                senha TEXT
            )
            """
        )
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS movimentacoes (
                id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                usuario TEXT,
                tipo TEXT,
                descricao TEXT,
                valor DOUBLE PRECISION,
                categoria TEXT,
                conta TEXT,
                data TEXT
            )
            """
        )
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS configuracoes (
                usuario TEXT PRIMARY KEY,
                idioma TEXT NOT NULL DEFAULT 'pt-BR',
                moeda TEXT NOT NULL DEFAULT 'BRL',
                tema TEXT NOT NULL DEFAULT 'ocean',
                plano TEXT NOT NULL DEFAULT 'free',
                alerta_limite DOUBLE PRECISION NOT NULL DEFAULT 0
            )
            """
        )

        c.execute(
            """
            CREATE TABLE IF NOT EXISTS categorias (
                id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                usuario TEXT NOT NULL,
                nome TEXT NOT NULL,
                cor TEXT NOT NULL DEFAULT '#58d5ff',
                UNIQUE(usuario, nome)
            )
            """
        )

        c.execute(
            """
            CREATE TABLE IF NOT EXISTS contas (
                id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                usuario TEXT NOT NULL,
                nome TEXT NOT NULL,
                tipo TEXT NOT NULL DEFAULT 'bank',
                UNIQUE(usuario, nome)
            )
            """
        )
    else:
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT UNIQUE,
                senha TEXT
            )
            """
        )
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS movimentacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT,
                tipo TEXT,
                descricao TEXT,
                valor REAL,
                categoria TEXT,
                conta TEXT,
                data TEXT
            )
            """
        )
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS configuracoes (
                usuario TEXT PRIMARY KEY,
                idioma TEXT NOT NULL DEFAULT 'pt-BR',
                moeda TEXT NOT NULL DEFAULT 'BRL',
                tema TEXT NOT NULL DEFAULT 'ocean',
                plano TEXT NOT NULL DEFAULT 'free',
                alerta_limite REAL NOT NULL DEFAULT 0
            )
            """
        )

        c.execute(
            """
            CREATE TABLE IF NOT EXISTS categorias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT NOT NULL,
                nome TEXT NOT NULL,
                cor TEXT NOT NULL DEFAULT '#58d5ff',
                UNIQUE(usuario, nome)
            )
            """
        )

        c.execute(
            """
            CREATE TABLE IF NOT EXISTS contas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT NOT NULL,
                nome TEXT NOT NULL,
                tipo TEXT NOT NULL DEFAULT 'bank',
                UNIQUE(usuario, nome)
            )
            """
        )

    movimentacoes_columns = table_columns(c, "movimentacoes")
    if "categoria_id" not in movimentacoes_columns:
        if DB_BACKEND == "postgres":
            c.execute("ALTER TABLE movimentacoes ADD COLUMN IF NOT EXISTS categoria_id INTEGER")
        else:
            c.execute("ALTER TABLE movimentacoes ADD COLUMN categoria_id INTEGER")
    if "conta_id" not in movimentacoes_columns:
        if DB_BACKEND == "postgres":
            c.execute("ALTER TABLE movimentacoes ADD COLUMN IF NOT EXISTS conta_id INTEGER")
        else:
            c.execute("ALTER TABLE movimentacoes ADD COLUMN conta_id INTEGER")
    if "created_at" not in movimentacoes_columns:
        if DB_BACKEND == "postgres":
            c.execute("ALTER TABLE movimentacoes ADD COLUMN IF NOT EXISTS created_at TEXT")
        else:
            c.execute("ALTER TABLE movimentacoes ADD COLUMN created_at TEXT")
        c.execute(
            """
            UPDATE movimentacoes
            SET created_at = COALESCE(NULLIF(data, ''), date('now'))
            WHERE created_at IS NULL OR created_at = ''
            """
        )

    configuracoes_columns = table_columns(c, "configuracoes")
    if "plano" not in configuracoes_columns:
        if DB_BACKEND == "postgres":
            c.execute("ALTER TABLE configuracoes ADD COLUMN IF NOT EXISTS plano TEXT NOT NULL DEFAULT 'free'")
        else:
            c.execute("ALTER TABLE configuracoes ADD COLUMN plano TEXT NOT NULL DEFAULT 'free'")
    if "alerta_limite" not in configuracoes_columns:
        if DB_BACKEND == "postgres":
            c.execute("ALTER TABLE configuracoes ADD COLUMN IF NOT EXISTS alerta_limite DOUBLE PRECISION NOT NULL DEFAULT 0")
        else:
            c.execute("ALTER TABLE configuracoes ADD COLUMN alerta_limite REAL NOT NULL DEFAULT 0")

    conn.commit()
    conn.close()


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        if backend_mode_enabled() and not session.get("api_token"):
            session.clear()
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped_view


@app.errorhandler(BackendAuthError)
def handle_backend_auth_error(_error):
    clear_backend_session()
    return redirect(url_for("login"))


def get_texts(language):
    return TRANSLATIONS.get(language, TRANSLATIONS["pt-BR"])


def get_month_options(language):
    labels = MONTH_LABELS.get(language, MONTH_LABELS["pt-BR"])
    return [(f"{index:02d}", label) for index, label in enumerate(labels, start=1)]


def parse_amount(value):
    cleaned = (value or "").strip().replace(" ", "")
    if not cleaned:
        raise ValueError("empty")
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    else:
        cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def parse_date_filter(value):
    raw = (value or "").strip()
    if not raw:
        return ""
    try:
        return date.fromisoformat(raw).isoformat()
    except ValueError:
        return ""


def parse_optional_int(value):
    raw = (value or "").strip()
    if not raw or not raw.isdigit():
        return None
    return int(raw)


def validate_login_credentials(usuario, senha, language):
    texts = get_texts(language)
    if not usuario or not senha:
        return texts["missing_credentials"]
    return None


def validate_registration_credentials(usuario, senha, language):
    texts = get_texts(language)
    if not usuario or not senha:
        return texts["missing_credentials"]
    if len(usuario) < 3:
        return texts["username_too_short"]
    if len(senha) < 4:
        return texts["password_too_short"]
    return None


def format_currency(value, currency):
    amount = float(value or 0)
    if currency in {"BRL", "EUR"}:
        number = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        number = f"{amount:,.2f}"
    return f"{CURRENCY_SYMBOLS.get(currency, currency)} {number}"


def ensure_user_settings(usuario):
    if backend_mode_enabled():
        return
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT usuario FROM configuracoes WHERE usuario=?", (usuario,))
    if not c.fetchone():
        c.execute(
            "INSERT INTO configuracoes (usuario, idioma, moeda, tema, plano, alerta_limite) VALUES (?, ?, ?, ?, ?, ?)",
            (
                usuario,
                DEFAULT_SETTINGS["idioma"],
                DEFAULT_SETTINGS["moeda"],
                DEFAULT_SETTINGS["tema"],
                DEFAULT_SETTINGS["plano"],
                0,
            ),
        )
        conn.commit()
    conn.close()


def get_settings(usuario):
    if not usuario:
        return DEFAULT_SETTINGS.copy()
    if backend_mode_enabled():
        backend_user = refresh_backend_me() or backend_user_from_session()
        return {
            "idioma": backend_user.get("language", DEFAULT_SETTINGS["idioma"]),
            "moeda": backend_user.get("currency", DEFAULT_SETTINGS["moeda"]),
            "tema": normalize_theme_key(backend_user.get("theme_key", DEFAULT_SETTINGS["tema"]), backend_user.get("plan", "free")),
            "plano": backend_user.get("plan", "free"),
            "alerta_limite": float(backend_user.get("monthly_limit") or 0),
        }
    ensure_user_settings(usuario)
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT idioma, moeda, tema, plano, alerta_limite FROM configuracoes WHERE usuario=?", (usuario,))
    row = c.fetchone()
    conn.close()
    if not row:
        return DEFAULT_SETTINGS.copy()
    return {
        "idioma": row["idioma"],
        "moeda": row["moeda"],
        "tema": normalize_theme_key(row["tema"], row["plano"] if row["plano"] in PLANS else "free"),
        "plano": row["plano"] if row["plano"] in PLANS else "free",
        "alerta_limite": float(row["alerta_limite"] or 0),
    }


def save_settings(usuario, idioma, moeda, tema, plano=None, alerta_limite=None):
    normalized_plan = plano if plano in PLANS else get_plan_key(usuario)
    normalized_theme = normalize_theme_key(tema, normalized_plan)
    if backend_mode_enabled():
        payload = {
            "language": idioma,
            "currency": moeda,
            "theme_key": normalized_theme,
            "monthly_limit": float(alerta_limite or 0),
        }
        response = backend_api_request("/me/settings", method="PATCH", token=get_backend_token(), body=payload)
        update_backend_session_user(response)
        if hasattr(g, "backend_snapshot"):
            delattr(g, "backend_snapshot")
        return
    if plano is None:
        plano = get_plan_key(usuario)
    plano = normalized_plan
    if alerta_limite is None:
        alerta_limite = get_settings(usuario).get("alerta_limite", 0)
    conn = conectar()
    c = conn.cursor()
    c.execute(
        """
        INSERT INTO configuracoes (usuario, idioma, moeda, tema, plano, alerta_limite)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(usuario) DO UPDATE SET
            idioma=excluded.idioma,
            moeda=excluded.moeda,
            tema=excluded.tema,
            plano=excluded.plano,
            alerta_limite=excluded.alerta_limite
        """,
        (usuario, idioma, moeda, normalized_theme, plano, alerta_limite),
    )
    conn.commit()
    conn.close()


def get_plan_key(usuario):
    if backend_mode_enabled():
        return get_settings(usuario).get("plano", "free")
    if PUBLIC_BETA_MODE:
        return "free"
    settings = get_settings(usuario)
    return settings.get("plano", "free")


def get_plan_config(plan_key):
    if backend_mode_enabled():
        return PLANS.get(plan_key, PLANS["free"])
    if PUBLIC_BETA_MODE:
        return PLANS["free"]
    return PLANS.get(plan_key, PLANS["free"])


def get_theme_options_for_plan(plan_key):
    accent_options = get_accent_theme_options_for_plan(plan_key)
    return [(compose_theme_key(background_key, accent_key), f"{background_label} + {accent_label}") for background_key, background_label in BACKGROUND_THEME_OPTIONS for accent_key, accent_label in accent_options]


def get_background_theme_options():
    return BACKGROUND_THEME_OPTIONS


def get_accent_theme_options_for_plan(plan_key):
    return PAID_ACCENT_THEME_OPTIONS


def is_paid_user(usuario):
    if backend_mode_enabled():
        return get_plan_key(usuario) == "pro"
    if PUBLIC_BETA_MODE:
        return False
    return get_plan_key(usuario) == "pro"


def count_transactions_this_month(usuario):
    if backend_mode_enabled():
        current_month = date.today().strftime("%Y-%m")
        rows = fetch_transactions(usuario)
        return sum(1 for row in rows if (row["created_at"] or row["data"] or "").startswith(current_month))
    conn = conectar()
    c = conn.cursor()
    current_month = date.today().strftime("%Y-%m")
    if DB_BACKEND == "postgres":
        c.execute(
            """
            SELECT COUNT(*) AS total
            FROM movimentacoes
            WHERE usuario=%s
              AND LEFT(COALESCE(NULLIF(created_at, ''), CAST(CURRENT_DATE AS TEXT)), 7) = %s
            """,
            (usuario, current_month),
        )
    else:
        c.execute(
            """
            SELECT COUNT(*) AS total
            FROM movimentacoes
            WHERE usuario=?
              AND strftime('%Y-%m', COALESCE(NULLIF(created_at, ''), date('now'))) = strftime('%Y-%m', date('now'))
            """,
            (usuario,),
        )
    row = c.fetchone()
    conn.close()
    return row["total"] if row else 0


def count_user_categories(usuario):
    if backend_mode_enabled():
        return len(get_user_categories(usuario))
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) AS total FROM categorias WHERE usuario=?", (usuario,))
    row = c.fetchone()
    conn.close()
    return row["total"] if row else 0


def count_user_accounts(usuario):
    if backend_mode_enabled():
        return len(get_user_accounts(usuario))
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) AS total FROM contas WHERE usuario=?", (usuario,))
    row = c.fetchone()
    conn.close()
    return row["total"] if row else 0


def get_available_years(usuario):
    if backend_mode_enabled():
        years = sorted({(row["data"] or "")[:4] for row in fetch_transactions(usuario) if row.get("data")}, reverse=True)
        return [year for year in years if year]
    conn = conectar()
    c = conn.cursor()
    if DB_BACKEND == "postgres":
        c.execute(
            """
            SELECT DISTINCT LEFT(data, 4) AS ano
            FROM movimentacoes
            WHERE usuario=%s AND data IS NOT NULL AND data != ''
            ORDER BY ano DESC
            """,
            (usuario,),
        )
    else:
        c.execute(
            """
            SELECT DISTINCT strftime('%Y', data) AS ano
            FROM movimentacoes
            WHERE usuario=? AND data IS NOT NULL AND data != ''
            ORDER BY ano DESC
            """,
            (usuario,),
        )
    rows = c.fetchall()
    conn.close()
    return [row["ano"] for row in rows if row["ano"]]


def enforce_free_plan_limit(usuario, resource, language, extra=None):
    texts = get_texts(language)
    plan = get_plan_config(get_plan_key(usuario))

    if resource == "transaction":
        limit = plan["transaction_limit_per_month"]
        if limit is not None and count_transactions_this_month(usuario) >= limit:
            return texts["limit_transactions_reached"]
    elif resource == "category":
        limit = plan["category_limit"]
        if limit is not None and count_user_categories(usuario) >= limit:
            return texts["limit_categories_reached"]
    elif resource == "account":
        limit = plan["account_limit"]
        if limit is not None and count_user_accounts(usuario) >= limit:
            return texts["limit_accounts_reached"]
    elif resource == "color":
        if not plan["allow_unlimited_colors"] and extra and extra not in FREE_ALLOWED_COLORS:
            return texts["limited_color_message"]
    elif resource == "currency":
        if not plan["allow_multi_currency"]:
            return texts["premium_feature_message"]

    return None


def ensure_default_categories_and_accounts(usuario):
    if backend_mode_enabled():
        return
    conn = conectar()
    c = conn.cursor()

    default_categories = [
        ("Alimentação", "#58d5ff"),
        ("Moradia", "#7cf0bb"),
        ("Transporte", "#ffab7a"),
        ("Salário", "#72f0a9"),
    ]
    default_accounts = [
        ("Conta principal", "bank"),
    ]

    for nome, cor in default_categories:
        c.execute(
            "INSERT OR IGNORE INTO categorias (usuario, nome, cor) VALUES (?, ?, ?)",
            (usuario, nome, cor),
        )

    for nome, tipo in default_accounts:
        c.execute(
            "INSERT OR IGNORE INTO contas (usuario, nome, tipo) VALUES (?, ?, ?)",
            (usuario, nome, tipo),
        )

    c.execute(
        """
        SELECT id, categoria, categoria_id
        FROM movimentacoes
        WHERE usuario=? AND categoria IS NOT NULL AND categoria != ''
        """,
        (usuario,),
    )
    for row in c.fetchall():
        c.execute(
            "INSERT OR IGNORE INTO categorias (usuario, nome, cor) VALUES (?, ?, ?)",
            (usuario, row["categoria"], "#58d5ff"),
        )
        if not row["categoria_id"]:
            c.execute(
                "SELECT id FROM categorias WHERE usuario=? AND nome=?",
                (usuario, row["categoria"]),
            )
            categoria = c.fetchone()
            if categoria:
                c.execute(
                    "UPDATE movimentacoes SET categoria_id=? WHERE id=?",
                    (categoria["id"], row["id"]),
                )

    c.execute(
        """
        SELECT id, conta, conta_id
        FROM movimentacoes
        WHERE usuario=? AND conta IS NOT NULL AND conta != ''
        """,
        (usuario,),
    )
    for row in c.fetchall():
        c.execute(
            "INSERT OR IGNORE INTO contas (usuario, nome, tipo) VALUES (?, ?, ?)",
            (usuario, row["conta"], "bank"),
        )
        if not row["conta_id"]:
            c.execute(
                "SELECT id FROM contas WHERE usuario=? AND nome=?",
                (usuario, row["conta"]),
            )
            conta = c.fetchone()
            if conta:
                c.execute(
                    "UPDATE movimentacoes SET conta_id=? WHERE id=?",
                    (conta["id"], row["id"]),
                )

    conn.commit()
    conn.close()


def get_user_categories(usuario):
    if backend_mode_enabled():
        categories = list((get_backend_snapshot().get("categories") or []))
        return [
            {"id": item["id"], "nome": item["name"], "cor": item.get("color", "#58d5ff")}
            for item in categories
        ]
    ensure_default_categories_and_accounts(usuario)
    conn = conectar()
    c = conn.cursor()
    c.execute(
        "SELECT id, nome, cor FROM categorias WHERE usuario=? ORDER BY nome COLLATE NOCASE",
        (usuario,),
    )
    rows = c.fetchall()
    conn.close()
    return rows


def get_user_accounts(usuario):
    if backend_mode_enabled():
        accounts = list((get_backend_snapshot().get("accounts") or []))
        return [
            {"id": item["id"], "nome": item["name"], "tipo": item.get("kind", "bank")}
            for item in accounts
        ]
    ensure_default_categories_and_accounts(usuario)
    conn = conectar()
    c = conn.cursor()
    c.execute(
        "SELECT id, nome, tipo FROM contas WHERE usuario=? ORDER BY nome COLLATE NOCASE",
        (usuario,),
    )
    rows = c.fetchall()
    conn.close()
    return rows


def get_category_name(usuario, categoria_id):
    if not categoria_id:
        return ""
    if backend_mode_enabled():
        item = backend_categories_map().get(categoria_id)
        return item.get("name", "") if item else ""
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT nome FROM categorias WHERE id=? AND usuario=?", (categoria_id, usuario))
    row = c.fetchone()
    conn.close()
    return row["nome"] if row else ""


def get_account_name(usuario, conta_id):
    if not conta_id:
        return ""
    if backend_mode_enabled():
        item = backend_accounts_map().get(conta_id)
        return item.get("name", "") if item else ""
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT nome FROM contas WHERE id=? AND usuario=?", (conta_id, usuario))
    row = c.fetchone()
    conn.close()
    return row["nome"] if row else ""


def build_insights(dados, settings):
    top_category = "-"
    categories = Counter(dado["categoria_nome"] for dado in dados if dado["categoria_nome"])
    if categories:
        top_category = categories.most_common(1)[0][0]

    latest_date = "-"
    for dado in dados:
        if dado["data"]:
            latest_date = dado["data"]
            break

    return {
        "entries_count": len(dados),
        "top_category": top_category,
        "latest_date": latest_date,
        "active_currency": settings["moeda"],
    }


def build_advanced_reports(dados, language):
    month_names = MONTH_LABELS.get(language, MONTH_LABELS["pt-BR"])
    income_by_month = {index: 0 for index in range(1, 13)}
    expense_by_month = {index: 0 for index in range(1, 13)}
    category_totals = Counter()

    for dado in dados:
        data = dado["data"] or ""
        month_number = None
        if len(data) >= 7:
            try:
                month_number = int(data[5:7])
            except ValueError:
                month_number = None

        if dado["tipo"] == "receita":
            if month_number:
                income_by_month[month_number] += float(dado["valor"] or 0)
        else:
            if month_number:
                expense_by_month[month_number] += float(dado["valor"] or 0)
            category_name = dado["categoria_nome"] or "-"
            category_totals[category_name] += float(dado["valor"] or 0)

    max_month_value = max([1, *income_by_month.values(), *expense_by_month.values()])
    monthly_series = []
    for index in range(1, 13):
        income_value = income_by_month[index]
        expense_value = expense_by_month[index]
        monthly_series.append(
            {
                "label": month_names[index - 1][:3],
                "income": income_value,
                "expense": expense_value,
                "income_pct": round((income_value / max_month_value) * 100, 2) if max_month_value else 0,
                "expense_pct": round((expense_value / max_month_value) * 100, 2) if max_month_value else 0,
            }
        )

    top_categories = category_totals.most_common(6)
    max_category_value = max([1, *(value for _, value in top_categories)]) if top_categories else 1
    category_series = [
        {
            "label": name,
            "value": value,
            "pct": round((value / max_category_value) * 100, 2) if max_category_value else 0,
        }
        for name, value in top_categories
    ]

    comparison = {
        "income_total": sum(income_by_month.values()),
        "expense_total": sum(expense_by_month.values()),
    }
    comparison["balance_total"] = comparison["income_total"] - comparison["expense_total"]

    return {
        "monthly_series": monthly_series,
        "category_series": category_series,
        "comparison": comparison,
        "has_data": bool(dados),
    }


def build_alerts(dados, settings, language):
    plan = get_plan_config(settings.get("plano", "free"))
    if not plan["allow_alerts"]:
        return {"enabled": False, "triggered": False, "expense_total": 0, "threshold": 0}

    threshold = float(settings.get("alerta_limite") or 0)
    expense_total = sum(float(dado["valor"] or 0) for dado in dados if dado["tipo"] == "despesa")
    return {
        "enabled": threshold > 0,
        "triggered": threshold > 0 and expense_total > threshold,
        "expense_total": expense_total,
        "threshold": threshold,
        "message": get_texts(language)["overspending_alert"],
    }


def fetch_transactions(usuario, data_inicial="", data_final=""):
    if backend_mode_enabled():
        snapshot = get_backend_snapshot()
        accounts = backend_accounts_map()
        categories = backend_categories_map()
        rows = []
        for item in snapshot.get("transactions") or []:
            raw_date = item.get("date", "")
            if data_inicial and raw_date < data_inicial:
                continue
            if data_final and raw_date > data_final:
                continue
            tx_type = item.get("type", "despesa")
            raw_value = float(item.get("value") or 0)
            rows.append(
                {
                    "id": item["id"],
                    "tipo": tx_type,
                    "descricao": item.get("description", ""),
                    "valor": abs(raw_value),
                    "categoria_nome": categories.get(item.get("category_id"), {}).get("name", ""),
                    "conta_nome": accounts.get(item.get("account_id"), {}).get("name", ""),
                    "data": raw_date,
                    "created_at": raw_date,
                    "categoria_id": item.get("category_id"),
                    "conta_id": item.get("account_id"),
                }
            )
        rows.sort(key=lambda row: ((row["data"] or ""), row["id"]), reverse=True)
        return rows
    conn = conectar()
    c = conn.cursor()
    query = """
        SELECT
            m.id,
            m.tipo,
            m.descricao,
            m.valor,
            COALESCE(cat.nome, m.categoria) AS categoria_nome,
            COALESCE(acc.nome, m.conta) AS conta_nome,
            m.data,
            m.created_at
        FROM movimentacoes m
        LEFT JOIN categorias cat ON cat.id = m.categoria_id
        LEFT JOIN contas acc ON acc.id = m.conta_id
        WHERE m.usuario=?
    """
    params = [usuario]
    if data_inicial:
        query += " AND m.data >= ?"
        params.append(data_inicial)
    if data_final:
        query += " AND m.data <= ?"
        params.append(data_final)
    query += " ORDER BY m.data DESC, m.id DESC"
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return rows


def resolve_assistant_period(data_inicial_raw="", data_final_raw=""):
    today = date.today()
    month_start = today.replace(day=1)
    start_date = parse_date_filter(data_inicial_raw) or month_start.isoformat()
    end_date = parse_date_filter(data_final_raw) or today.isoformat()
    if start_date > end_date:
        start_date, end_date = end_date, start_date
    return start_date, end_date


def build_period_label(language, start_date, end_date):
    month_names = MONTH_LABELS.get(language, MONTH_LABELS["pt-BR"])
    if start_date and end_date and start_date[:7] == end_date[:7]:
        month_index = int(start_date[5:7]) - 1
        return f"{month_names[month_index]} {start_date[:4]}"
    if start_date and end_date:
        return f"{start_date} - {end_date}"
    return get_texts(language)["assistant_current_month"]


def build_assistant_payload(usuario, settings, data_inicial="", data_final=""):
    start_date, end_date = resolve_assistant_period(data_inicial, data_final)
    transactions = fetch_transactions(usuario, start_date, end_date)
    period_label = build_period_label(settings["idioma"], start_date, end_date)
    return build_financial_assistant_report(
        transactions=transactions,
        settings=settings,
        period_label=period_label,
        start_date=start_date,
        end_date=end_date,
        provider="local",
    )


def build_local_assistant_chat_reply(question, assistant_payload):
    normalized = str(question or "").strip().lower()
    suggestions = [
        "Como sincronizar meu app com o site?",
        "Se aparecer token invalido, o que faco?",
        "Resumo do meu mes",
        "O que muda do Gratis para o Pro?",
    ]

    if any(term in normalized for term in ["sincron", "token", "api", "site", "app", "login"]):
        answer = (
            "Use a mesma conta no app e no site. No app, mantenha a URL da API em "
            "https://notafacil-api.onrender.com. Se aparecer token invalido, "
            "saia da conta e entre novamente para renovar a sessao."
        )
        mode = "account"
    elif any(term in normalized for term in ["plano", "pro", "free", "gratis", "senha"]):
        answer = (
            "O plano da conta vem do backend. Se voce quiser mudar entre Free e Pro, "
            "faça isso com a conta online conectada. Para senha, use Configuracoes > Acesso da conta."
        )
        mode = "help"
    else:
        snapshot = assistant_payload["snapshot"]
        answer = (
            f"No periodo {snapshot['period']['label']}, voce tem {snapshot['entries_count']} lancamentos, "
            f"receitas de {format_currency(snapshot['totals']['income'], snapshot['currency'])}, "
            f"gastos de {format_currency(snapshot['totals']['expense'], snapshot['currency'])} e "
            f"saldo de {format_currency(snapshot['totals']['balance'], snapshot['currency'])}. "
            f"A categoria com maior gasto foi {snapshot['top_expense_category']['name']}."
        )
        mode = "finance"

    return {
        "answer": answer,
        "provider": "local",
        "mode": mode,
        "provider_reason": "local_fallback",
        "suggestions": suggestions,
    }


def get_assistant_provider_reason_text(texts, reason):
    mapping = {
        "missing_api_key": texts.get("assistant_provider_reason_missing_api_key", ""),
        "local_fallback": texts.get("assistant_provider_reason_local_fallback", ""),
        "provider_http_error": texts.get("assistant_provider_reason_provider_http_error", ""),
        "provider_request_failed": texts.get("assistant_provider_reason_provider_request_failed", ""),
        "provider_empty_output": texts.get("assistant_provider_reason_provider_empty_output", ""),
        "unsupported_provider": texts.get("assistant_provider_reason_unsupported_provider", ""),
        "local_provider_selected": texts.get("assistant_provider_reason_local_provider_selected", ""),
    }
    return mapping.get(reason, "")


def build_backup_payload(usuario):
    settings = get_settings(usuario)
    categorias = get_user_categories(usuario)
    contas = get_user_accounts(usuario)
    dados = fetch_transactions(usuario)
    return {
        "usuario": usuario,
        "plano": settings.get("plano", "free"),
        "configuracoes": settings,
        "categorias": [dict(item) for item in categorias],
        "contas": [dict(item) for item in contas],
        "movimentacoes": [dict(item) for item in dados],
    }


def write_local_backup(usuario):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    payload = build_backup_payload(usuario)
    backup_file = BACKUP_DIR / f"{usuario}_latest.json"
    backup_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


@app.context_processor
def inject_layout_context():
    user = session.get("user")
    settings = get_settings(user) if user else DEFAULT_SETTINGS.copy()
    texts = get_texts(settings["idioma"])
    active_plan_key = get_plan_key(user) if user else "free"
    effective_beta_mode = PUBLIC_BETA_MODE and not backend_mode_enabled()
    backend_user = backend_user_from_session() if user else {}
    account_email = backend_user.get("email", "") if isinstance(backend_user, dict) else ""
    account_is_online = bool(user and backend_mode_enabled() and get_backend_token())
    settings["plano"] = active_plan_key
    settings["tema"] = normalize_theme_key(settings.get("tema"), active_plan_key)
    allowed_theme_values = {key for key, _ in get_theme_options_for_plan(active_plan_key)}
    if settings.get("tema") not in allowed_theme_values:
        settings["tema"] = normalize_theme_key(DEFAULT_SETTINGS["tema"], active_plan_key)
    selected_background_theme, selected_accent_theme = split_theme_key(settings["tema"])
    account_type_options = [
        ("wallet", texts["wallet"]),
        ("bank", texts["bank"]),
        ("credit_card", texts["credit_card"]),
        ("savings", texts["savings"]),
    ]
    return {
        "app_display_name": APP_DISPLAY_NAME,
        "app_tagline": APP_TAGLINE,
        "company_name": COMPANY_NAME,
        "support_email": SUPPORT_EMAIL,
        "beta_public_mode": effective_beta_mode,
        "current_year": date.today().year,
        "settings": settings,
        "texts": texts,
        "plan_key": active_plan_key,
        "plan_config": get_plan_config(active_plan_key),
        "plans_catalog": PLANS,
        "free_allowed_colors": FREE_ALLOWED_COLORS,
        "paid_color_presets": PAID_COLOR_PRESETS,
        "month_options": get_month_options(settings["idioma"]),
        "language_options": LANGUAGE_OPTIONS,
        "currency_options": CURRENCY_OPTIONS,
        "theme_options": get_theme_options_for_plan(active_plan_key),
        "background_theme_options": get_background_theme_options(),
        "accent_theme_options": get_accent_theme_options_for_plan(active_plan_key),
        "selected_background_theme": selected_background_theme,
        "selected_accent_theme": selected_accent_theme,
        "theme_background_class": get_background_theme_class(settings["tema"]),
        "theme_accent_class": get_accent_theme_class(settings["tema"]),
        "account_type_options": account_type_options,
        "account_username": user or "",
        "account_email": account_email,
        "account_is_online": account_is_online,
        "backend_unavailable": getattr(g, "backend_unavailable", False),
        "backend_status_message": getattr(g, "backend_status_message", ""),
        "format_money": lambda value: format_currency(value, settings["moeda"]),
    }


criar_tabelas()


@app.route("/login", methods=["GET", "POST"])
def login():
    if "user" in session:
        return redirect(url_for("index"))

    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "")
        error = validate_login_credentials(usuario, senha, "pt-BR")
        if error:
            return render_template("login_v3.html", error=error)

        if backend_mode_enabled():
            try:
                payload = backend_api_request(
                    "/auth/login",
                    method="POST",
                    body={"username": usuario, "password": senha},
                )
                session["api_token"] = payload["access_token"]
                update_backend_session_user(payload["user"])
                get_backend_snapshot(force=True)
                return redirect(url_for("index"))
            except ValueError as exc:
                error = str(exc) or TRANSLATIONS["pt-BR"]["wrong_password"]
            except BackendRateLimitError as exc:
                error = str(exc)
            except ConnectionError:
                error = "Nao foi possivel conectar ao backend agora."
            return render_template("login_v3.html", error=error)

        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE usuario=?", (usuario,))
        user = c.fetchone()

        if user:
            if check_password_hash(user["senha"], senha):
                conn.close()
                session["user"] = usuario
                session.permanent = True
                ensure_user_settings(usuario)
                ensure_default_categories_and_accounts(usuario)
                return redirect(url_for("index"))
            error = TRANSLATIONS["pt-BR"]["wrong_password"]
            conn.close()
        else:
            conn.close()
            error = TRANSLATIONS["pt-BR"]["user_not_found"]

    return render_template("login_v3.html", error=error)


@app.route("/cadastro", methods=["GET", "POST"])
def cadastro():
    if "user" in session:
        return redirect(url_for("index"))

    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "")
        error = validate_registration_credentials(usuario, senha, "pt-BR")
        if error:
            return render_template("cadastro_v3.html", error=error)

        if backend_mode_enabled():
            if "@" not in email or len(email) < 5:
                error = "Informe um e-mail valido."
                return render_template("cadastro_v3.html", error=error)
            if len(senha) < 6:
                error = "Use uma senha com pelo menos 6 caracteres."
                return render_template("cadastro_v3.html", error=error)
            try:
                payload = backend_api_request(
                    "/auth/register",
                    method="POST",
                    body={"username": usuario, "email": email, "password": senha},
                )
                session["api_token"] = payload["access_token"]
                update_backend_session_user(payload["user"])
                get_backend_snapshot(force=True)
                flash(get_texts("pt-BR")["account_created"], "success")
                return redirect(url_for("index"))
            except ValueError as exc:
                error = str(exc) or TRANSLATIONS["pt-BR"]["user_exists"]
            except BackendRateLimitError as exc:
                error = str(exc)
            except ConnectionError:
                error = "Nao foi possivel conectar ao backend agora."
            return render_template("cadastro_v3.html", error=error)

        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT id FROM usuarios WHERE usuario=?", (usuario,))
        existing = c.fetchone()
        if existing:
            conn.close()
            error = TRANSLATIONS["pt-BR"]["user_exists"]
            return render_template("cadastro_v3.html", error=error)

        c.execute(
            "INSERT INTO usuarios (usuario, senha) VALUES (?, ?)",
            (usuario, generate_password_hash(senha)),
        )
        conn.commit()
        conn.close()
        ensure_user_settings(usuario)
        ensure_default_categories_and_accounts(usuario)
        session["user"] = usuario
        session.permanent = True
        write_local_backup(usuario)
        flash(get_texts("pt-BR")["account_created"], "success")
        return redirect(url_for("index"))

    return render_template("cadastro_v3.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/atualizar-dados")
@login_required
def atualizar_dados():
    if backend_mode_enabled():
        refresh_backend_me()
        get_backend_snapshot(force=True)
        flash("Dados atualizados com o backend oficial.", "success")
    else:
        flash("Conta local atualizada neste dispositivo.", "info")
    next_url = request.args.get("next") or url_for("index")
    if not str(next_url).startswith("/"):
        next_url = url_for("index")
    return redirect(next_url)


@app.route("/")
@login_required
def index():
    user = session["user"]
    settings = get_settings(user)
    texts = get_texts(settings["idioma"])
    ensure_default_categories_and_accounts(user)
    data_inicial = parse_date_filter(request.args.get("data_inicial"))
    data_final = parse_date_filter(request.args.get("data_final"))

    if data_inicial and data_final and data_inicial > data_final:
        data_inicial, data_final = data_final, data_inicial

    dados = fetch_transactions(user, data_inicial, data_final)
    anos_disponiveis = get_available_years(user)

    receita = sum(dado["valor"] for dado in dados if dado["tipo"] == "receita")
    despesa = sum(dado["valor"] for dado in dados if dado["tipo"] == "despesa")
    saldo = receita - despesa
    insights = build_insights(dados, settings)
    advanced_reports = build_advanced_reports(dados, settings["idioma"])
    alerts = build_alerts(dados, settings, settings["idioma"])
    assistant_preview = build_financial_assistant_report(
        transactions=dados,
        settings=settings,
        period_label=build_period_label(settings["idioma"], data_inicial or date.today().replace(day=1).isoformat(), data_final or date.today().isoformat()),
        start_date=data_inicial or date.today().replace(day=1).isoformat(),
        end_date=data_final or date.today().isoformat(),
        provider="local",
    )

    return render_template(
        "dashboard_v3.html",
        dados=dados,
        receita=round(receita, 2),
        despesa=round(despesa, 2),
        saldo=round(saldo, 2),
        filtros={"data_inicial": data_inicial, "data_final": data_final},
        anos_disponiveis=anos_disponiveis,
        insights=insights,
        categorias=get_user_categories(user),
        contas=get_user_accounts(user),
        advanced_reports=advanced_reports,
        alerts=alerts,
        assistant_preview=assistant_preview,
    )


@app.route("/assistente", methods=["GET", "POST"])
@login_required
def assistente_financeiro():
    user = session["user"]
    settings = get_settings(user)
    data_inicial, data_final = resolve_assistant_period(
        request.args.get("data_inicial"),
        request.args.get("data_final"),
    )
    assistant = build_assistant_payload(user, settings, data_inicial, data_final)
    assistant_chat = None
    assistant_question = ""

    if request.method == "POST":
        assistant_question = request.form.get("question", "").strip()
        data_inicial, data_final = resolve_assistant_period(
            request.form.get("data_inicial"),
            request.form.get("data_final"),
        )
        assistant = build_assistant_payload(user, settings, data_inicial, data_final)
        if assistant_question:
            if backend_mode_enabled():
                try:
                    assistant_chat = backend_api_request(
                        "/assistant/chat",
                        method="POST",
                        token=get_backend_token(),
                        body={"question": assistant_question},
                    )
                except (ValueError, ConnectionError):
                    flash("O backend do assistente nao respondeu agora. Vou usar a leitura local do periodo.", "warning")
                    assistant_chat = build_local_assistant_chat_reply(assistant_question, assistant)
            else:
                assistant_chat = build_local_assistant_chat_reply(assistant_question, assistant)

    return render_template(
        "assistente_v3.html",
        assistant=assistant,
        filtros={"data_inicial": data_inicial, "data_final": data_final},
        assistant_chat=assistant_chat,
        assistant_question=assistant_question,
        assistant_provider_reason_text=get_assistant_provider_reason_text(get_texts(settings["idioma"]), assistant_chat.get("provider_reason")) if assistant_chat else "",
        assistant_api_url=url_for(
            "assistente_financeiro_api",
            data_inicial=data_inicial,
            data_final=data_final,
        ),
    )


@app.route("/assistente/analisar")
@login_required
def assistente_financeiro_api():
    user = session["user"]
    settings = get_settings(user)
    data_inicial, data_final = resolve_assistant_period(
        request.args.get("data_inicial"),
        request.args.get("data_final"),
    )
    assistant = build_assistant_payload(user, settings, data_inicial, data_final)
    return {
        "app": "DiGiaI Caixa",
        "feature": "Assistente Financeiro",
        "provider": assistant["provider"],
        "period": assistant["snapshot"]["period"],
        "totals": assistant["snapshot"]["totals"],
        "top_expense_category": assistant["snapshot"]["top_expense_category"],
        "financial_alert": assistant["financial_alert"],
        "saving_tip": assistant["saving_tip"],
        "summary": assistant["summary"],
        "future_bridge": assistant["future_bridge"],
    }


@app.route("/add", methods=["POST"])
@login_required
def add():
    language = get_settings(session["user"])["idioma"]
    plan_error = enforce_free_plan_limit(session["user"], "transaction", language)
    if plan_error:
        flash(plan_error, "error")
        return redirect(url_for("index"))

    try:
        valor = parse_amount(request.form.get("valor"))
    except ValueError:
        flash(get_texts(language)["invalid_amount"], "error")
        return redirect(url_for("index"))

    categoria_id = parse_optional_int(request.form.get("categoria_id"))
    conta_id = parse_optional_int(request.form.get("conta_id"))
    categoria_nome = get_category_name(session["user"], categoria_id)
    conta_nome = get_account_name(session["user"], conta_id)

    if backend_mode_enabled():
        try:
            backend_api_request(
                "/transactions",
                method="POST",
                token=get_backend_token(),
                body={
                    "description": request.form.get("descricao", "").strip(),
                    "value": valor,
                    "type": request.form.get("tipo", "despesa"),
                    "date": request.form.get("data", ""),
                    "account_id": conta_id,
                    "category_id": categoria_id,
                },
            )
            refresh_backend_me()
            get_backend_snapshot(force=True)
            flash(get_texts(get_settings(session["user"])["idioma"])["transaction_created"], "success")
        except ValueError as exc:
            flash(str(exc) or get_texts(language)["premium_feature_message"], "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("index"))

    conn = conectar()
    c = conn.cursor()
    c.execute(
        """
        INSERT INTO movimentacoes (usuario, tipo, descricao, valor, categoria, conta, data, categoria_id, conta_id, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            session["user"],
            request.form.get("tipo", "despesa"),
            request.form.get("descricao", "").strip(),
            valor,
            categoria_nome,
            conta_nome,
            request.form.get("data", ""),
            categoria_id,
            conta_id,
            date.today().isoformat(),
        ),
    )
    conn.commit()
    conn.close()
    write_local_backup(session["user"])
    flash(get_texts(get_settings(session["user"])["idioma"])["transaction_created"], "success")
    return redirect(url_for("index"))


@app.route("/editar/<int:item_id>", methods=["GET", "POST"])
@login_required
def editar(item_id):
    if backend_mode_enabled():
        dado = next((row for row in fetch_transactions(session["user"]) if row["id"] == item_id), None)
        if not dado:
            flash(get_texts(get_settings(session["user"])["idioma"])["not_found"], "error")
            return redirect(url_for("index"))

        if request.method == "POST":
            try:
                valor = parse_amount(request.form.get("valor"))
            except ValueError:
                flash(get_texts(get_settings(session["user"])["idioma"])["invalid_amount"], "error")
                return redirect(url_for("editar", item_id=item_id))

            try:
                backend_api_request(
                    f"/transactions/{item_id}",
                    method="PUT",
                    token=get_backend_token(),
                    body={
                        "description": request.form.get("descricao", "").strip(),
                        "value": valor,
                        "type": request.form.get("tipo", "despesa"),
                        "date": request.form.get("data", ""),
                        "account_id": parse_optional_int(request.form.get("conta_id")),
                        "category_id": parse_optional_int(request.form.get("categoria_id")),
                    },
                )
                get_backend_snapshot(force=True)
                flash(get_texts(get_settings(session["user"])["idioma"])["transaction_updated"], "success")
                return redirect(url_for("index"))
            except ValueError as exc:
                flash(str(exc) or "Erro no backend.", "error")
                return redirect(url_for("editar", item_id=item_id))
            except ConnectionError:
                flash("Nao foi possivel conectar ao backend agora.", "error")
                return redirect(url_for("editar", item_id=item_id))

        return render_template(
            "editar_v3.html",
            dado=dado,
            categorias=get_user_categories(session["user"]),
            contas=get_user_accounts(session["user"]),
        )

    conn = conectar()
    c = conn.cursor()
    c.execute(
        """
        SELECT id, tipo, descricao, valor, categoria, conta, data, categoria_id, conta_id
        FROM movimentacoes
        WHERE id=? AND usuario=?
        """,
        (item_id, session["user"]),
    )
    dado = c.fetchone()
    if not dado:
        conn.close()
        flash(get_texts(get_settings(session["user"])["idioma"])["not_found"], "error")
        return redirect(url_for("index"))

    if request.method == "POST":
        try:
            valor = parse_amount(request.form.get("valor"))
        except ValueError:
            conn.close()
            flash(get_texts(get_settings(session["user"])["idioma"])["invalid_amount"], "error")
            return redirect(url_for("editar", item_id=item_id))

        c.execute(
            """
            UPDATE movimentacoes
            SET tipo=?, descricao=?, valor=?, categoria=?, conta=?, data=?, categoria_id=?, conta_id=?
            WHERE id=? AND usuario=?
            """,
            (
                request.form.get("tipo", "despesa"),
                request.form.get("descricao", "").strip(),
                valor,
                get_category_name(session["user"], parse_optional_int(request.form.get("categoria_id"))),
                get_account_name(session["user"], parse_optional_int(request.form.get("conta_id"))),
                request.form.get("data", ""),
                parse_optional_int(request.form.get("categoria_id")),
                parse_optional_int(request.form.get("conta_id")),
                item_id,
                session["user"],
            ),
        )
        conn.commit()
        conn.close()
        write_local_backup(session["user"])
        flash(get_texts(get_settings(session["user"])["idioma"])["transaction_updated"], "success")
        return redirect(url_for("index"))

    conn.close()
    return render_template(
        "editar_v3.html",
        dado=dado,
        categorias=get_user_categories(session["user"]),
        contas=get_user_accounts(session["user"]),
    )


@app.route("/delete/<int:item_id>", methods=["POST"])
@login_required
def delete(item_id):
    if backend_mode_enabled():
        try:
            backend_api_request(f"/transactions/{item_id}", method="DELETE", token=get_backend_token())
            get_backend_snapshot(force=True)
            flash(get_texts(get_settings(session["user"])["idioma"])["transaction_deleted"], "success")
        except ValueError as exc:
            flash(str(exc) or "Erro no backend.", "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("index"))

    conn = conectar()
    c = conn.cursor()
    c.execute("DELETE FROM movimentacoes WHERE id=? AND usuario=?", (item_id, session["user"]))
    conn.commit()
    conn.close()
    write_local_backup(session["user"])
    flash(get_texts(get_settings(session["user"])["idioma"])["transaction_deleted"], "success")
    return redirect(url_for("index"))


@app.route("/configuracoes", methods=["GET", "POST"])
@login_required
def configuracoes():
    user = session["user"]
    current_settings = get_settings(user)
    language = current_settings["idioma"]

    if request.method == "POST":
        form_name = request.form.get("form_name", "settings").strip().lower()
        if form_name == "password":
            if not backend_mode_enabled():
                flash("A conta local nao permite troca de senha online nesta tela.", "info")
                return redirect(url_for("configuracoes"))

            current_password = request.form.get("current_password", "")
            new_password = request.form.get("new_password", "")
            confirm_password = request.form.get("confirm_password", "")

            if not current_password:
                flash("Informe a senha atual.", "error")
                return redirect(url_for("configuracoes"))
            if len(new_password) < 6:
                flash("Use uma nova senha com pelo menos 6 caracteres.", "error")
                return redirect(url_for("configuracoes"))
            if new_password != confirm_password:
                flash("A confirmacao da nova senha nao confere.", "error")
                return redirect(url_for("configuracoes"))

            try:
                backend_api_request(
                    "/me/password",
                    method="PATCH",
                    token=get_backend_token(),
                    body={"current_password": current_password, "new_password": new_password},
                )
                flash("Senha atualizada com sucesso.", "success")
            except ValueError as exc:
                flash(str(exc) or "Nao foi possivel atualizar a senha agora.", "error")
            except ConnectionError:
                flash("Nao foi possivel conectar ao backend agora.", "error")
            return redirect(url_for("configuracoes"))

        idioma = request.form.get("idioma", DEFAULT_SETTINGS["idioma"])
        moeda = request.form.get("moeda", DEFAULT_SETTINGS["moeda"])
        background_theme = request.form.get("background_theme", split_theme_key(DEFAULT_SETTINGS["tema"])[0])
        accent_theme = request.form.get("accent_theme", split_theme_key(DEFAULT_SETTINGS["tema"])[1])
        tema = compose_theme_key(background_theme, accent_theme)
        allowed_themes = {item[0] for item in get_theme_options_for_plan(get_plan_key(user))}

        if idioma not in {item[0] for item in LANGUAGE_OPTIONS}:
            idioma = DEFAULT_SETTINGS["idioma"]
        if moeda not in {item[0] for item in CURRENCY_OPTIONS}:
            moeda = DEFAULT_SETTINGS["moeda"]
        if tema not in allowed_themes:
            tema = normalize_theme_key(DEFAULT_SETTINGS["tema"], get_plan_key(user))

        if not get_plan_config(get_plan_key(user))["allow_multi_currency"]:
            moeda = "BRL"

        alerta_limite = 0
        if get_plan_config(get_plan_key(user))["allow_alerts"]:
            try:
                alerta_limite = parse_amount(request.form.get("alerta_limite", "0"))
            except ValueError:
                alerta_limite = 0

        if backend_mode_enabled():
            try:
                payload = backend_api_request(
                    "/me/settings",
                    method="PATCH",
                    token=get_backend_token(),
                    body={
                        "language": idioma,
                        "currency": moeda,
                        "theme_key": tema,
                        "monthly_limit": alerta_limite,
                    },
                )
                update_backend_session_user(payload)
                get_backend_snapshot(force=True)
                current_settings = get_settings(user)
                flash(get_texts(current_settings["idioma"])["settings_saved"], "success")
            except ValueError as exc:
                flash(str(exc) or "Nao foi possivel salvar as preferencias agora.", "error")
            except ConnectionError:
                flash("Nao foi possivel conectar ao backend agora.", "error")
            return redirect(url_for("configuracoes"))

        save_settings(user, idioma, moeda, tema, alerta_limite=alerta_limite)
        write_local_backup(user)
        flash(get_texts(idioma)["settings_saved"], "success")
        return redirect(url_for("configuracoes"))

    return render_template("configuracoes_v3.html", current_settings=current_settings)


@app.route("/planos")
@login_required
def planos():
    return render_template("planos_v3.html")


@app.route("/planos/<plan_key>", methods=["POST"])
@login_required
def alterar_plano(plan_key):
    if plan_key not in PLANS:
        return redirect(url_for("planos"))

    user = session["user"]
    if backend_mode_enabled():
        try:
            payload = backend_api_request(
                "/me/plan",
                method="PATCH",
                token=get_backend_token(),
                body={"plan": plan_key},
            )
            update_backend_session_user(payload)
            get_backend_snapshot(force=True)
            flash(f"{get_texts(get_settings(user)['idioma'])['current_plan']}: {PLANS[plan_key]['name']}", "success")
        except ValueError as exc:
            flash(str(exc) or "Erro ao alterar o plano.", "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("planos"))

    if PUBLIC_BETA_MODE:
        flash(f"A beta publica do {APP_DISPLAY_NAME} funciona apenas no plano gratuito nesta fase.", "info")
        return redirect(url_for("planos"))

    settings = get_settings(user)
    moeda = settings["moeda"] if PLANS[plan_key]["allow_multi_currency"] else "BRL"
    alerta_limite = settings.get("alerta_limite", 0) if PLANS[plan_key]["allow_alerts"] else 0
    allowed_themes = {item[0] for item in get_theme_options_for_plan(plan_key)}
    tema = settings["tema"] if settings["tema"] in allowed_themes else normalize_theme_key(DEFAULT_SETTINGS["tema"], plan_key)
    save_settings(user, settings["idioma"], moeda, tema, plan_key, alerta_limite)
    write_local_backup(user)
    flash(f"{get_texts(settings['idioma'])['current_plan']}: {PLANS[plan_key]['name']}", "success")
    return redirect(url_for("planos"))


@app.route("/organizacao")
@login_required
def organizacao():
    user = session["user"]
    return render_template(
        "organizacao_v3.html",
        categorias=get_user_categories(user),
        contas=get_user_accounts(user),
    )


@app.route("/categorias", methods=["POST"])
@login_required
def criar_categoria():
    user = session["user"]
    settings = get_settings(user)
    texts = get_texts(settings["idioma"])
    nome = request.form.get("nome", "").strip()
    cor = request.form.get("cor", "#58d5ff").strip() or "#58d5ff"

    if not nome:
        flash(texts["field_required"], "error")
        return redirect(url_for("organizacao"))

    plan_error = enforce_free_plan_limit(user, "category", settings["idioma"])
    if plan_error:
        flash(plan_error, "error")
        return redirect(url_for("organizacao"))

    plan_error = enforce_free_plan_limit(user, "color", settings["idioma"], cor)
    if plan_error:
        flash(plan_error, "error")
        return redirect(url_for("organizacao"))

    if backend_mode_enabled():
        try:
            backend_api_request(
                "/categories",
                method="POST",
                token=get_backend_token(),
                body={"name": nome, "color": cor},
            )
            get_backend_snapshot(force=True)
            flash(texts["category_created"], "success")
        except ValueError as exc:
            flash(str(exc) or texts["category_exists"], "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO categorias (usuario, nome, cor) VALUES (?, ?, ?)",
            (user, nome, cor),
        )
        conn.commit()
        write_local_backup(user)
        flash(texts["category_created"], "success")
    except DB_INTEGRITY_ERRORS:
        flash(texts["category_exists"], "error")
    finally:
        conn.close()

    return redirect(url_for("organizacao"))


@app.route("/exportar/excel")
@login_required
def exportar_excel():
    user = session["user"]
    data_inicial = parse_date_filter(request.args.get("data_inicial"))
    data_final = parse_date_filter(request.args.get("data_final"))
    dados = fetch_transactions(user, data_inicial, data_final)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lancamentos"
    sheet.append(["ID", "Tipo", "Descricao", "Categoria", "Conta", "Data", "Valor"])
    for row in dados:
        sheet.append([
            row["id"],
            row["tipo"],
            row["descricao"],
            row["categoria_nome"],
            row["conta_nome"],
            row["data"],
            row["valor"],
        ])

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="lancamentos.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/exportar/pdf")
@login_required
def exportar_pdf():
    user = session["user"]
    data_inicial = parse_date_filter(request.args.get("data_inicial"))
    data_final = parse_date_filter(request.args.get("data_final"))
    dados = fetch_transactions(user, data_inicial, data_final)
    settings = get_settings(user)

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 40
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(40, y, "Financeiro - Lancamentos")
    y -= 24
    pdf.setFont("Helvetica", 10)

    for row in dados:
        line = f"{row['data'] or '-'} | {row['tipo']} | {row['descricao']} | {row['categoria_nome'] or '-'} | {format_currency(row['valor'], settings['moeda'])}"
        pdf.drawString(40, y, line[:110])
        y -= 16
        if y <= 40:
            pdf.showPage()
            y = height - 40
            pdf.setFont("Helvetica", 10)

    pdf.save()
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="lancamentos.pdf", mimetype="application/pdf")


@app.route("/backup/download")
@login_required
def backup_download():
    user = session["user"]
    payload = build_backup_payload(user)

    buffer = BytesIO(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"backup_{user}.json", mimetype="application/json")


@app.route("/categorias/<int:categoria_id>/editar", methods=["POST"])
@login_required
def editar_categoria(categoria_id):
    user = session["user"]
    settings = get_settings(user)
    texts = get_texts(settings["idioma"])
    nome = request.form.get("nome", "").strip()
    cor = request.form.get("cor", "#58d5ff").strip() or "#58d5ff"

    if not nome:
        flash(texts["field_required"], "error")
        return redirect(url_for("organizacao"))

    plan_error = enforce_free_plan_limit(user, "color", settings["idioma"], cor)
    if plan_error:
        flash(plan_error, "error")
        return redirect(url_for("organizacao"))

    if backend_mode_enabled():
        try:
            backend_api_request(
                f"/categories/{categoria_id}",
                method="PUT",
                token=get_backend_token(),
                body={"name": nome, "color": cor},
            )
            get_backend_snapshot(force=True)
            flash(texts["category_updated"], "success")
        except ValueError as exc:
            flash(str(exc) or texts["category_exists"], "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    try:
        c.execute(
            """
            UPDATE categorias
            SET nome=?, cor=?
            WHERE id=? AND usuario=?
            """,
            (nome, cor, categoria_id, user),
        )
        c.execute(
            """
            UPDATE movimentacoes
            SET categoria=?
            WHERE usuario=? AND categoria_id=?
            """,
            (nome, user, categoria_id),
        )
        conn.commit()
        write_local_backup(user)
        flash(texts["category_updated"], "success")
    except DB_INTEGRITY_ERRORS:
        flash(texts["category_exists"], "error")
    finally:
        conn.close()

    return redirect(url_for("organizacao"))


@app.route("/categorias/<int:categoria_id>/excluir", methods=["POST"])
@login_required
def excluir_categoria(categoria_id):
    user = session["user"]
    texts = get_texts(get_settings(user)["idioma"])
    if backend_mode_enabled():
        try:
            backend_api_request(f"/categories/{categoria_id}", method="DELETE", token=get_backend_token())
            get_backend_snapshot(force=True)
            flash(texts["category_deleted"], "success")
        except ValueError as exc:
            flash(str(exc) or "Erro ao excluir categoria.", "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT nome FROM categorias WHERE id=? AND usuario=?", (categoria_id, user))
    categoria = c.fetchone()
    categoria_nome = categoria["nome"] if categoria else ""
    c.execute(
        "UPDATE movimentacoes SET categoria_id=NULL, categoria=? WHERE usuario=? AND categoria_id=?",
        (categoria_nome, user, categoria_id),
    )
    c.execute("DELETE FROM categorias WHERE id=? AND usuario=?", (categoria_id, user))
    conn.commit()
    conn.close()
    write_local_backup(user)
    flash(texts["category_deleted"], "success")
    return redirect(url_for("organizacao"))


@app.route("/contas", methods=["POST"])
@login_required
def criar_conta():
    user = session["user"]
    settings = get_settings(user)
    texts = get_texts(settings["idioma"])
    nome = request.form.get("nome", "").strip()
    tipo = request.form.get("tipo", "bank").strip() or "bank"

    if not nome:
        flash(texts["field_required"], "error")
        return redirect(url_for("organizacao"))

    if tipo not in {"wallet", "bank", "credit_card", "savings"}:
        tipo = "bank"

    plan_error = enforce_free_plan_limit(user, "account", settings["idioma"])
    if plan_error:
        flash(plan_error, "error")
        return redirect(url_for("organizacao"))

    if backend_mode_enabled():
        try:
            backend_api_request(
                "/accounts",
                method="POST",
                token=get_backend_token(),
                body={"name": nome, "kind": tipo},
            )
            get_backend_snapshot(force=True)
            flash(texts["account_created_item"], "success")
        except ValueError as exc:
            flash(str(exc) or texts["account_exists"], "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO contas (usuario, nome, tipo) VALUES (?, ?, ?)",
            (user, nome, tipo),
        )
        conn.commit()
        write_local_backup(user)
        flash(texts["account_created_item"], "success")
    except DB_INTEGRITY_ERRORS:
        flash(texts["account_exists"], "error")
    finally:
        conn.close()

    return redirect(url_for("organizacao"))


@app.route("/contas/<int:conta_id>/editar", methods=["POST"])
@login_required
def editar_conta(conta_id):
    user = session["user"]
    settings = get_settings(user)
    texts = get_texts(settings["idioma"])
    nome = request.form.get("nome", "").strip()
    tipo = request.form.get("tipo", "bank").strip() or "bank"

    if not nome:
        flash(texts["field_required"], "error")
        return redirect(url_for("organizacao"))

    if tipo not in {"wallet", "bank", "credit_card", "savings"}:
        tipo = "bank"

    if backend_mode_enabled():
        try:
            backend_api_request(
                f"/accounts/{conta_id}",
                method="PUT",
                token=get_backend_token(),
                body={"name": nome, "kind": tipo},
            )
            get_backend_snapshot(force=True)
            flash(texts["account_updated"], "success")
        except ValueError as exc:
            flash(str(exc) or texts["account_exists"], "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    try:
        c.execute(
            """
            UPDATE contas
            SET nome=?, tipo=?
            WHERE id=? AND usuario=?
            """,
            (nome, tipo, conta_id, user),
        )
        c.execute(
            """
            UPDATE movimentacoes
            SET conta=?
            WHERE usuario=? AND conta_id=?
            """,
            (nome, user, conta_id),
        )
        conn.commit()
        write_local_backup(user)
        flash(texts["account_updated"], "success")
    except DB_INTEGRITY_ERRORS:
        flash(texts["account_exists"], "error")
    finally:
        conn.close()

    return redirect(url_for("organizacao"))


@app.route("/contas/<int:conta_id>/excluir", methods=["POST"])
@login_required
def excluir_conta(conta_id):
    user = session["user"]
    texts = get_texts(get_settings(user)["idioma"])
    if backend_mode_enabled():
        try:
            backend_api_request(f"/accounts/{conta_id}", method="DELETE", token=get_backend_token())
            get_backend_snapshot(force=True)
            flash(texts["account_deleted"], "success")
        except ValueError as exc:
            flash(str(exc) or "Erro ao excluir conta.", "error")
        except ConnectionError:
            flash("Nao foi possivel conectar ao backend agora.", "error")
        return redirect(url_for("organizacao"))

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT nome FROM contas WHERE id=? AND usuario=?", (conta_id, user))
    conta = c.fetchone()
    conta_nome = conta["nome"] if conta else ""
    c.execute(
        "UPDATE movimentacoes SET conta_id=NULL, conta=? WHERE usuario=? AND conta_id=?",
        (conta_nome, user, conta_id),
    )
    c.execute("DELETE FROM contas WHERE id=? AND usuario=?", (conta_id, user))
    conn.commit()
    conn.close()
    write_local_backup(user)
    flash(texts["account_deleted"], "success")
    return redirect(url_for("organizacao"))


@app.route("/offline")
def offline():
    return render_template("offline_v3.html")


@app.route("/politica-de-privacidade")
def politica_privacidade():
    return render_template("politica_privacidade_v3.html")


@app.route("/termos-de-uso")
def termos_de_uso():
    return render_template("termos_v3.html")


@app.route("/health")
def health():
    return {
        "status": "ok",
        "mode": "beta-publica",
        "database_backend": DB_BACKEND,
        "database_path": str(DB_PATH),
        "backend_mode": backend_mode_enabled(),
        "backend_url": BACKEND_API_URL or None,
    }


@app.route("/manifest.webmanifest")
def manifest():
    data = {
        "name": APP_DISPLAY_NAME,
        "short_name": APP_DISPLAY_NAME,
        "start_url": url_for("index"),
        "display": "standalone",
        "background_color": "#08131f",
        "theme_color": "#08131f",
        "description": APP_TAGLINE,
        "icons": [
            {
                "src": url_for("static", filename="icons/icon-app.png"),
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any",
            },
            {
                "src": url_for("static", filename="icons/icon-maskable.png"),
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "maskable",
            },
        ],
    }
    response = make_response(json.dumps(data))
    response.mimetype = "application/manifest+json"
    return response


@app.route("/service-worker.js")
def service_worker():
    response = make_response(send_from_directory(app.static_folder, "service-worker.js"))
    response.mimetype = "application/javascript"
    response.headers["Cache-Control"] = "no-cache"
    return response


if __name__ == "__main__":
    app.run(
        debug=os.environ.get("FLASK_DEBUG", "true").lower() == "true",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "5000")),
    )
