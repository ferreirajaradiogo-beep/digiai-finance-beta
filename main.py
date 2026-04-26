from datetime import date
from pathlib import Path
import sqlite3

from kivy.app import App
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle
from kivy.metrics import dp
from kivy.properties import ListProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from werkzeug.security import check_password_hash, generate_password_hash


COLORS = {
    "bg": [0.03, 0.07, 0.12, 1],
    "surface": [0.07, 0.13, 0.22, 1],
    "surface_soft": [0.10, 0.18, 0.30, 1],
    "accent": [0.35, 0.84, 1.0, 1],
    "accent_dark": [0.00, 0.44, 0.70, 1],
    "text": [0.93, 0.97, 1.0, 1],
    "muted": [0.62, 0.72, 0.84, 1],
    "danger": [1.0, 0.48, 0.58, 1],
    "success": [0.45, 0.94, 0.66, 1],
}

Window.clearcolor = COLORS["bg"]

FREE_ALLOWED_COLORS = ["#58d5ff", "#7cf0bb", "#ffab7a", "#ff7c93"]
PAID_ALLOWED_COLORS = FREE_ALLOWED_COLORS + [
    "#b794f4",
    "#f6ad55",
    "#4fd1c5",
    "#f687b3",
    "#90cdf4",
    "#68d391",
    "#f56565",
    "#fbd38d",
]
FREE_THEMES = ["ocean", "graphite", "forest", "sunset"]
PAID_THEMES = FREE_THEMES + ["aurora", "ember", "royal", "sand"]
CURRENCIES = ["BRL", "USD", "EUR", "GBP"]
ACCOUNT_TYPES = ["wallet", "bank", "credit_card", "savings"]


def get_db_path():
    app = App.get_running_app()
    if app:
        return Path(app.user_data_dir) / "financeiro.db"
    return Path(__file__).resolve().parent / "financeiro.db"


def conectar():
    db_path = get_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def table_columns(cursor, table_name):
    cursor.execute(f"PRAGMA table_info({table_name})")
    return {row["name"] for row in cursor.fetchall()}


def criar_tabelas():
    conn = conectar()
    c = conn.cursor()
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
    columns = table_columns(c, "movimentacoes")
    if "categoria_id" not in columns:
        c.execute("ALTER TABLE movimentacoes ADD COLUMN categoria_id INTEGER")
    if "conta_id" not in columns:
        c.execute("ALTER TABLE movimentacoes ADD COLUMN conta_id INTEGER")
    if "created_at" not in columns:
        c.execute("ALTER TABLE movimentacoes ADD COLUMN created_at TEXT")
    conn.commit()
    conn.close()


def show_message(title, message):
    box = RoundedBox(orientation="vertical", padding=dp(14), spacing=dp(10))
    box.add_widget(MutedLabel(text=message))
    close = PrimaryButton(text="OK")
    box.add_widget(close)
    popup = Popup(title=title, content=box, size_hint=(0.84, 0.38))
    close.bind(on_release=popup.dismiss)
    popup.open()


def parse_amount(value):
    raw = (value or "").strip().replace(" ", "")
    if not raw:
        raise ValueError("empty")
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".")
    else:
        raw = raw.replace(",", ".")
    return float(raw)


def money(value, currency="BRL"):
    symbols = {"BRL": "R$", "USD": "$", "EUR": "EUR", "GBP": "GBP"}
    return f"{symbols.get(currency, currency)} {float(value or 0):.2f}"


class RoundedBox(BoxLayout):
    bg_color = ListProperty(COLORS["surface"])

    def __init__(self, radius=22, **kwargs):
        super().__init__(**kwargs)
        self.radius = dp(radius)
        with self.canvas.before:
            Color(*self.bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[self.radius])
        self.bind(pos=self.update_rect, size=self.update_rect, bg_color=self.update_color)

    def update_rect(self, *_):
        self.rect.pos = self.pos
        self.rect.size = self.size

    def update_color(self, *_):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[self.radius])


class PrimaryButton(Button):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_down", "")
        kwargs.setdefault("background_color", COLORS["accent"])
        kwargs.setdefault("color", [0.02, 0.08, 0.12, 1])
        kwargs.setdefault("bold", True)
        kwargs.setdefault("size_hint_y", None)
        kwargs.setdefault("height", dp(48))
        super().__init__(**kwargs)


class SoftButton(Button):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_down", "")
        kwargs.setdefault("background_color", COLORS["surface_soft"])
        kwargs.setdefault("color", COLORS["text"])
        kwargs.setdefault("size_hint_y", None)
        kwargs.setdefault("height", dp(44))
        super().__init__(**kwargs)


class DangerButton(Button):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_down", "")
        kwargs.setdefault("background_color", COLORS["danger"])
        kwargs.setdefault("color", [0.12, 0.02, 0.04, 1])
        kwargs.setdefault("bold", True)
        super().__init__(**kwargs)


class AppTextInput(TextInput):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_active", "")
        kwargs.setdefault("background_color", COLORS["surface_soft"])
        kwargs.setdefault("foreground_color", COLORS["text"])
        kwargs.setdefault("cursor_color", COLORS["accent"])
        kwargs.setdefault("hint_text_color", COLORS["muted"])
        kwargs.setdefault("padding", [dp(12), dp(12), dp(12), dp(12)])
        super().__init__(**kwargs)


class AppSpinner(Spinner):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_color", COLORS["surface_soft"])
        kwargs.setdefault("color", COLORS["text"])
        super().__init__(**kwargs)


class TitleLabel(Label):
    def __init__(self, **kwargs):
        kwargs.setdefault("color", COLORS["text"])
        kwargs.setdefault("bold", True)
        super().__init__(**kwargs)


class MutedLabel(Label):
    def __init__(self, **kwargs):
        kwargs.setdefault("color", COLORS["muted"])
        super().__init__(**kwargs)


class StatCard(RoundedBox):
    def __init__(self, title, accent_color, **kwargs):
        kwargs.setdefault("orientation", "vertical")
        kwargs.setdefault("padding", dp(12))
        kwargs.setdefault("spacing", dp(4))
        kwargs.setdefault("bg_color", COLORS["surface"])
        super().__init__(**kwargs)
        self.title_label = MutedLabel(text=title, font_size=dp(12), size_hint_y=None, height=dp(20))
        self.value_label = TitleLabel(text="0", font_size=dp(17))
        self.accent = accent_color
        self.add_widget(self.title_label)
        self.add_widget(self.value_label)
        with self.canvas.after:
            Color(*accent_color)
            self.accent_rect = RoundedRectangle(pos=self.pos, size=(dp(4), self.height), radius=[dp(4)])
        self.bind(pos=self.update_accent, size=self.update_accent)

    def update_accent(self, *_):
        self.accent_rect.pos = self.pos
        self.accent_rect.size = (dp(4), self.height)

    def set_value(self, value):
        self.value_label.text = value


class Badge(Label):
    def __init__(self, text, bg_color, **kwargs):
        kwargs.setdefault("text", text)
        kwargs.setdefault("color", COLORS["text"])
        kwargs.setdefault("bold", True)
        kwargs.setdefault("font_size", dp(12))
        kwargs.setdefault("size_hint_x", None)
        kwargs.setdefault("width", dp(82))
        super().__init__(**kwargs)
        self.bg_color = bg_color
        with self.canvas.before:
            Color(*bg_color)
            self.badge_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(12)])
        self.bind(pos=self.update_badge, size=self.update_badge)

    def update_badge(self, *_):
        self.badge_rect.pos = self.pos
        self.badge_rect.size = self.size


def get_settings(usuario):
    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT * FROM configuracoes WHERE usuario=?", (usuario,))
    row = c.fetchone()
    conn.close()
    if not row:
        return {"moeda": "BRL", "tema": "ocean", "plano": "free", "alerta_limite": 0}
    return {
        "moeda": row["moeda"],
        "tema": row["tema"],
        "plano": row["plano"] if row["plano"] in {"free", "pro"} else "free",
        "alerta_limite": float(row["alerta_limite"] or 0),
    }


def save_settings(usuario, moeda, tema, plano, alerta_limite):
    if plano == "free":
        moeda = "BRL"
        if tema not in FREE_THEMES:
            tema = "ocean"
        alerta_limite = 0
    conn = conectar()
    c = conn.cursor()
    c.execute(
        """
        INSERT INTO configuracoes (usuario, idioma, moeda, tema, plano, alerta_limite)
        VALUES (?, 'pt-BR', ?, ?, ?, ?)
        ON CONFLICT(usuario) DO UPDATE SET
            moeda=excluded.moeda,
            tema=excluded.tema,
            plano=excluded.plano,
            alerta_limite=excluded.alerta_limite
        """,
        (usuario, moeda, tema, plano, alerta_limite),
    )
    conn.commit()
    conn.close()


def ensure_defaults(usuario):
    save_settings(usuario, "BRL", "ocean", get_settings(usuario)["plano"], get_settings(usuario)["alerta_limite"])
    conn = conectar()
    c = conn.cursor()
    c.execute(
        "INSERT OR IGNORE INTO categorias (usuario, nome, cor) VALUES (?, 'Geral', '#58d5ff')",
        (usuario,),
    )
    c.execute(
        "INSERT OR IGNORE INTO contas (usuario, nome, tipo) VALUES (?, 'Conta principal', 'bank')",
        (usuario,),
    )
    conn.commit()
    conn.close()


def fetch_names(table, usuario):
    conn = conectar()
    c = conn.cursor()
    c.execute(f"SELECT nome FROM {table} WHERE usuario=? ORDER BY nome", (usuario,))
    rows = [row["nome"] for row in c.fetchall()]
    conn.close()
    return rows


def fetch_items(table, usuario):
    conn = conectar()
    c = conn.cursor()
    c.execute(f"SELECT * FROM {table} WHERE usuario=? ORDER BY nome", (usuario,))
    rows = c.fetchall()
    conn.close()
    return rows


def count_table(table, usuario):
    conn = conectar()
    c = conn.cursor()
    c.execute(f"SELECT COUNT(*) AS total FROM {table} WHERE usuario=?", (usuario,))
    row = c.fetchone()
    conn.close()
    return row["total"] if row else 0


def count_month_transactions(usuario):
    conn = conectar()
    c = conn.cursor()
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


class BaseScreen(Screen):
    def app(self):
        return App.get_running_app()

    def user(self):
        return self.app().usuario

    def nav_bar(self):
        bar = GridLayout(cols=5, spacing=dp(6), size_hint_y=None, height=dp(48))
        buttons = [
            ("Painel", "dashboard"),
            ("Organizar", "organization"),
            ("Planos", "plans"),
            ("Config", "settings"),
            ("Sair", "login"),
        ]
        for label, target in buttons:
            button = SoftButton(text=label)
            if target == "login":
                button.bind(on_release=lambda *_: self.logout())
            else:
                button.bind(on_release=lambda _, name=target: self.go(name))
            bar.add_widget(button)
        return bar

    def go(self, screen_name):
        self.manager.current = screen_name
        screen = self.manager.get_screen(screen_name)
        if hasattr(screen, "refresh"):
            screen.refresh()

    def logout(self):
        self.app().usuario = None
        self.manager.current = "login"


class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        layout = BoxLayout(orientation="vertical", padding=dp(18), spacing=dp(12))
        with layout.canvas.before:
            Color(*COLORS["bg"])
            layout.bg = RoundedRectangle(pos=layout.pos, size=layout.size, radius=[0])
        layout.bind(pos=lambda *_: setattr(layout.bg, "pos", layout.pos), size=lambda *_: setattr(layout.bg, "size", layout.size))

        hero = RoundedBox(orientation="vertical", padding=dp(22), spacing=dp(10), size_hint_y=None, height=dp(180))
        hero.add_widget(MutedLabel(text="FINANCEIRO OFFLINE", font_size=dp(12), size_hint_y=None, height=dp(22)))
        hero.add_widget(TitleLabel(text="Controle seu dinheiro", font_size=dp(28)))
        hero.add_widget(MutedLabel(text="App offline com login, planos e lancamentos.", font_size=dp(15)))
        layout.add_widget(hero)

        card = RoundedBox(orientation="vertical", padding=dp(18), spacing=dp(12))
        self.usuario = AppTextInput(hint_text="Usuario", multiline=False, size_hint_y=None, height=dp(50))
        self.senha = AppTextInput(hint_text="Senha", password=True, multiline=False, size_hint_y=None, height=dp(50))
        card.add_widget(self.usuario)
        card.add_widget(self.senha)
        login = PrimaryButton(text="Entrar")
        register = SoftButton(text="Criar conta")
        login.bind(on_release=self.login)
        register.bind(on_release=self.register)
        card.add_widget(login)
        card.add_widget(register)
        layout.add_widget(card)
        self.add_widget(layout)

    def login(self, *_):
        usuario = self.usuario.text.strip()
        senha = self.senha.text
        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT senha FROM usuarios WHERE usuario=?", (usuario,))
        row = c.fetchone()
        conn.close()
        if not row or not check_password_hash(row["senha"], senha):
            show_message("Atencao", "Usuario ou senha invalidos.")
            return
        ensure_defaults(usuario)
        App.get_running_app().usuario = usuario
        self.manager.current = "dashboard"
        self.manager.get_screen("dashboard").refresh()

    def register(self, *_):
        usuario = self.usuario.text.strip()
        senha = self.senha.text
        if len(usuario) < 3 or len(senha) < 4:
            show_message("Atencao", "Use usuario com 3+ caracteres e senha com 4+.")
            return
        conn = conectar()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO usuarios (usuario, senha) VALUES (?, ?)", (usuario, generate_password_hash(senha)))
            conn.commit()
        except sqlite3.IntegrityError:
            show_message("Atencao", "Usuario ja existe.")
            conn.close()
            return
        conn.close()
        ensure_defaults(usuario)
        show_message("Pronto", "Conta criada. Toque em Entrar.")


class DashboardScreen(BaseScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.root_box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        with self.root_box.canvas.before:
            Color(*COLORS["bg"])
            self.root_bg = RoundedRectangle(pos=self.root_box.pos, size=self.root_box.size, radius=[0])
        self.root_box.bind(pos=lambda *_: setattr(self.root_bg, "pos", self.root_box.pos), size=lambda *_: setattr(self.root_bg, "size", self.root_box.size))

        self.title = TitleLabel(text="", font_size=dp(20), size_hint_y=None, height=dp(34))
        self.summary = RoundedBox(orientation="vertical", padding=dp(14), size_hint_y=None, height=dp(104))
        self.summary_income = TitleLabel(text="", font_size=dp(16))
        self.summary_expense = TitleLabel(text="", font_size=dp(16))
        self.summary_balance = TitleLabel(text="", font_size=dp(16))
        self.summary.add_widget(self.summary_income)
        self.summary.add_widget(self.summary_expense)
        self.summary.add_widget(self.summary_balance)
        self.alert = MutedLabel(text="", size_hint_y=None, height=dp(30), color=COLORS["danger"])
        self.root_box.add_widget(self.nav_bar())
        hero = RoundedBox(orientation="vertical", padding=dp(16), spacing=dp(6), size_hint_y=None, height=dp(92))
        hero.add_widget(MutedLabel(text="DASHBOARD", font_size=dp(12), size_hint_y=None, height=dp(20)))
        hero.add_widget(self.title)
        hero.add_widget(MutedLabel(text="Controle offline com lancamentos e planos.", font_size=dp(14)))
        self.root_box.add_widget(hero)

        self.stats_grid = GridLayout(cols=3, spacing=dp(8), size_hint_y=None, height=dp(92))
        self.income_card = StatCard("Receita", COLORS["success"])
        self.expense_card = StatCard("Despesa", COLORS["danger"])
        self.balance_card = StatCard("Saldo", COLORS["accent"])
        self.stats_grid.add_widget(self.income_card)
        self.stats_grid.add_widget(self.expense_card)
        self.stats_grid.add_widget(self.balance_card)
        self.root_box.add_widget(self.stats_grid)
        self.root_box.add_widget(self.alert)
        self.root_box.add_widget(self.form())
        self.list_box = BoxLayout(orientation="vertical", spacing=dp(6), size_hint_y=None)
        self.list_box.bind(minimum_height=self.list_box.setter("height"))
        scroll = ScrollView()
        scroll.add_widget(self.list_box)
        self.root_box.add_widget(scroll)
        self.add_widget(self.root_box)

    def form(self):
        box = RoundedBox(orientation="vertical", padding=dp(12), spacing=dp(8), size_hint_y=None, height=dp(300))
        grid = GridLayout(cols=2, spacing=dp(6))
        self.tipo = AppSpinner(text="despesa", values=["receita", "despesa"])
        self.descricao = AppTextInput(hint_text="Descricao", multiline=False)
        self.valor = AppTextInput(hint_text="Valor", multiline=False)
        self.categoria = AppSpinner(text="Geral", values=["Geral"])
        self.conta = AppSpinner(text="Conta principal", values=["Conta principal"])
        self.data = AppTextInput(hint_text="Data AAAA-MM-DD", multiline=False)
        add = PrimaryButton(text="Adicionar")
        add.bind(on_release=self.add_transaction)
        widgets = [
            ("Tipo", self.tipo),
            ("Descricao", self.descricao),
            ("Valor", self.valor),
            ("Categoria", self.categoria),
            ("Conta", self.conta),
            ("Data", self.data),
        ]
        for label, widget in widgets:
            grid.add_widget(MutedLabel(text=label))
            grid.add_widget(widget)
        box.add_widget(grid)
        box.add_widget(add)
        return box

    def refresh(self):
        usuario = self.user()
        if not usuario:
            return
        settings = get_settings(usuario)
        categorias = fetch_names("categorias", usuario)
        contas = fetch_names("contas", usuario)
        self.categoria.values = categorias or ["Geral"]
        self.conta.values = contas or ["Conta principal"]
        self.categoria.text = self.categoria.values[0]
        self.conta.text = self.conta.values[0]
        conn = conectar()
        c = conn.cursor()
        c.execute(
            """
            SELECT * FROM movimentacoes
            WHERE usuario=?
            ORDER BY COALESCE(NULLIF(data, ''), '0000-00-00') DESC, id DESC
            """,
            (usuario,),
        )
        rows = c.fetchall()
        conn.close()
        receita = sum(row["valor"] for row in rows if row["tipo"] == "receita")
        despesa = sum(row["valor"] for row in rows if row["tipo"] == "despesa")
        saldo = receita - despesa
        self.title.text = f"Ola, {usuario} | Plano: {settings['plano']}"
        self.income_card.set_value(money(receita, settings["moeda"]))
        self.expense_card.set_value(money(despesa, settings["moeda"]))
        self.balance_card.set_value(money(saldo, settings["moeda"]))
        self.alert.text = ""
        if settings["plano"] == "pro" and settings["alerta_limite"] > 0 and despesa > settings["alerta_limite"]:
            self.alert.text = f"Alerta: despesas acima de {money(settings['alerta_limite'], settings['moeda'])}"
        self.render_list(rows)

    def render_list(self, rows):
        self.list_box.clear_widgets()
        for row in rows:
            line = RoundedBox(orientation="horizontal", size_hint_y=None, height=dp(68), spacing=dp(8), padding=dp(8), bg_color=COLORS["surface_soft"])
            badge_color = COLORS["success"] if row["tipo"] == "receita" else COLORS["danger"]
            line.add_widget(Badge(row["tipo"], badge_color))
            details = BoxLayout(orientation="vertical", spacing=dp(2))
            details.add_widget(TitleLabel(text=row["descricao"] or "-", font_size=dp(14)))
            details.add_widget(MutedLabel(text=f"{row['data'] or '-'} | {row['categoria'] or '-'} | {money(row['valor'], get_settings(self.user())['moeda'])}", font_size=dp(12)))
            line.add_widget(details)
            edit = SoftButton(text="Editar", size_hint_x=None, width=dp(70))
            delete = DangerButton(text="X", size_hint_x=None, width=dp(48))
            edit.bind(on_release=lambda _, item=row: self.open_edit_popup(item))
            delete.bind(on_release=lambda _, item_id=row["id"]: self.delete_transaction(item_id))
            line.add_widget(edit)
            line.add_widget(delete)
            self.list_box.add_widget(line)

    def add_transaction(self, *_):
        usuario = self.user()
        if get_settings(usuario)["plano"] == "free" and count_month_transactions(usuario) >= 20:
            show_message("Limite", "Plano gratuito: 20 lancamentos por mes.")
            return
        try:
            valor = parse_amount(self.valor.text)
        except ValueError:
            show_message("Atencao", "Informe um valor valido.")
            return
        conn = conectar()
        c = conn.cursor()
        c.execute(
            """
            INSERT INTO movimentacoes (usuario, tipo, descricao, valor, categoria, conta, data, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (usuario, self.tipo.text, self.descricao.text.strip(), valor, self.categoria.text, self.conta.text, self.data.text.strip(), date.today().isoformat()),
        )
        conn.commit()
        conn.close()
        self.descricao.text = ""
        self.valor.text = ""
        self.data.text = ""
        self.refresh()

    def open_edit_popup(self, row):
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))
        tipo = Spinner(text=row["tipo"], values=["receita", "despesa"])
        descricao = TextInput(text=row["descricao"] or "", multiline=False)
        valor = TextInput(text=str(row["valor"]), multiline=False)
        data_input = TextInput(text=row["data"] or "", multiline=False)
        box.add_widget(tipo)
        box.add_widget(descricao)
        box.add_widget(valor)
        box.add_widget(data_input)
        save = Button(text="Salvar", size_hint_y=None, height=dp(44))
        box.add_widget(save)
        popup = Popup(title="Editar lancamento", content=box, size_hint=(0.9, 0.65))

        def do_save(*_):
            try:
                parsed = parse_amount(valor.text)
            except ValueError:
                show_message("Atencao", "Valor invalido.")
                return
            conn = conectar()
            c = conn.cursor()
            c.execute(
                """
                UPDATE movimentacoes
                SET tipo=?, descricao=?, valor=?, data=?
                WHERE id=? AND usuario=?
                """,
                (tipo.text, descricao.text.strip(), parsed, data_input.text.strip(), row["id"], self.user()),
            )
            conn.commit()
            conn.close()
            popup.dismiss()
            self.refresh()

        save.bind(on_release=do_save)
        popup.open()

    def delete_transaction(self, item_id):
        conn = conectar()
        c = conn.cursor()
        c.execute("DELETE FROM movimentacoes WHERE id=? AND usuario=?", (item_id, self.user()))
        conn.commit()
        conn.close()
        self.refresh()


class OrganizationScreen(BaseScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        with self.box.canvas.before:
            Color(*COLORS["bg"])
            self.org_bg = RoundedRectangle(pos=self.box.pos, size=self.box.size, radius=[0])
        self.box.bind(pos=lambda *_: setattr(self.org_bg, "pos", self.box.pos), size=lambda *_: setattr(self.org_bg, "size", self.box.size))
        self.box.add_widget(self.nav_bar())
        self.box.add_widget(TitleLabel(text="Categorias e contas", font_size=dp(20), size_hint_y=None, height=dp(36)))
        self.box.add_widget(self.forms())
        self.items = BoxLayout(orientation="vertical", spacing=dp(6), size_hint_y=None)
        self.items.bind(minimum_height=self.items.setter("height"))
        scroll = ScrollView()
        scroll.add_widget(self.items)
        self.box.add_widget(scroll)
        self.add_widget(self.box)

    def forms(self):
        panel = RoundedBox(orientation="vertical", padding=dp(12), spacing=dp(8), size_hint_y=None, height=dp(220))
        grid = GridLayout(cols=2, spacing=dp(6))
        self.cat_name = AppTextInput(hint_text="Nova categoria", multiline=False)
        self.cat_color = AppSpinner(text=FREE_ALLOWED_COLORS[0], values=FREE_ALLOWED_COLORS)
        self.acc_name = AppTextInput(hint_text="Nova conta", multiline=False)
        self.acc_type = AppSpinner(text="bank", values=ACCOUNT_TYPES)
        add_cat = PrimaryButton(text="Criar categoria")
        add_acc = SoftButton(text="Criar conta")
        add_cat.bind(on_release=self.add_category)
        add_acc.bind(on_release=self.add_account)
        for widget in [self.cat_name, self.cat_color, self.acc_name, self.acc_type]:
            grid.add_widget(widget)
        panel.add_widget(grid)
        action_grid = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=dp(48))
        action_grid.add_widget(add_cat)
        action_grid.add_widget(add_acc)
        panel.add_widget(action_grid)
        return panel

    def refresh(self):
        usuario = self.user()
        settings = get_settings(usuario)
        self.cat_color.values = PAID_ALLOWED_COLORS if settings["plano"] == "pro" else FREE_ALLOWED_COLORS
        self.items.clear_widgets()
        self.items.add_widget(TitleLabel(text="Categorias", size_hint_y=None, height=dp(32)))
        for item in fetch_items("categorias", usuario):
            self.items.add_widget(self.row("categorias", item))
        self.items.add_widget(TitleLabel(text="Contas", size_hint_y=None, height=dp(32)))
        for item in fetch_items("contas", usuario):
            self.items.add_widget(self.row("contas", item))

    def row(self, table, item):
        line = RoundedBox(orientation="horizontal", size_hint_y=None, height=dp(58), spacing=dp(6), padding=dp(8), bg_color=COLORS["surface_soft"])
        name = AppTextInput(text=item["nome"], multiline=False)
        line.add_widget(name)
        save = SoftButton(text="Salvar", size_hint_x=None, width=dp(82))
        delete = DangerButton(text="Excluir", size_hint_x=None, width=dp(82))
        save.bind(on_release=lambda *_: self.update_item(table, item["id"], name.text))
        delete.bind(on_release=lambda *_: self.delete_item(table, item["id"]))
        line.add_widget(save)
        line.add_widget(delete)
        return line

    def add_category(self, *_):
        usuario = self.user()
        settings = get_settings(usuario)
        if settings["plano"] == "free" and count_table("categorias", usuario) >= 10:
            show_message("Limite", "Plano gratuito: 10 categorias.")
            return
        color = self.cat_color.text
        if settings["plano"] == "free" and color not in FREE_ALLOWED_COLORS:
            color = FREE_ALLOWED_COLORS[0]
        self.insert_item("categorias", self.cat_name.text.strip(), color)
        self.cat_name.text = ""

    def add_account(self, *_):
        usuario = self.user()
        if get_settings(usuario)["plano"] == "free" and count_table("contas", usuario) >= 1:
            show_message("Limite", "Plano gratuito: 1 conta.")
            return
        self.insert_item("contas", self.acc_name.text.strip(), self.acc_type.text)
        self.acc_name.text = ""

    def insert_item(self, table, name, extra):
        if not name:
            show_message("Atencao", "Informe um nome.")
            return
        usuario = self.user()
        conn = conectar()
        c = conn.cursor()
        try:
            if table == "categorias":
                c.execute("INSERT INTO categorias (usuario, nome, cor) VALUES (?, ?, ?)", (usuario, name, extra))
            else:
                c.execute("INSERT INTO contas (usuario, nome, tipo) VALUES (?, ?, ?)", (usuario, name, extra))
            conn.commit()
        except sqlite3.IntegrityError:
            show_message("Atencao", "Esse nome ja existe.")
        conn.close()
        self.refresh()

    def update_item(self, table, item_id, name):
        usuario = self.user()
        if not name.strip():
            return
        conn = conectar()
        c = conn.cursor()
        column = "categoria" if table == "categorias" else "conta"
        id_column = "categoria_id" if table == "categorias" else "conta_id"
        c.execute(f"UPDATE {table} SET nome=? WHERE id=? AND usuario=?", (name.strip(), item_id, usuario))
        c.execute(f"UPDATE movimentacoes SET {column}=? WHERE usuario=? AND {id_column}=?", (name.strip(), usuario, item_id))
        conn.commit()
        conn.close()
        self.refresh()

    def delete_item(self, table, item_id):
        usuario = self.user()
        conn = conectar()
        c = conn.cursor()
        id_column = "categoria_id" if table == "categorias" else "conta_id"
        c.execute(f"UPDATE movimentacoes SET {id_column}=NULL WHERE usuario=? AND {id_column}=?", (usuario, item_id))
        c.execute(f"DELETE FROM {table} WHERE id=? AND usuario=?", (item_id, usuario))
        conn.commit()
        conn.close()
        self.refresh()


class PlansScreen(BaseScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
        with box.canvas.before:
            Color(*COLORS["bg"])
            self.plan_bg = RoundedRectangle(pos=box.pos, size=box.size, radius=[0])
        box.bind(pos=lambda *_: setattr(self.plan_bg, "pos", box.pos), size=lambda *_: setattr(self.plan_bg, "size", box.size))
        box.add_widget(self.nav_bar())
        box.add_widget(TitleLabel(text="Planos", font_size=dp(22), size_hint_y=None, height=dp(40)))
        free_card = RoundedBox(orientation="vertical", padding=dp(16), spacing=dp(8))
        free_card.add_widget(MutedLabel(text="PLANO ATUAL / TESTE", font_size=dp(11), size_hint_y=None, height=dp(18)))
        free_card.add_widget(TitleLabel(text="Gratuito"))
        free_card.add_widget(MutedLabel(text="20 lancamentos/mes\n1 conta\n10 categorias\nCores limitadas"))
        free = SoftButton(text="Usar gratuito")
        free_card.add_widget(free)
        pro_card = RoundedBox(orientation="vertical", padding=dp(16), spacing=dp(8))
        pro_card.add_widget(MutedLabel(text="PREMIUM", font_size=dp(11), size_hint_y=None, height=dp(18), color=COLORS["accent"]))
        pro_card.add_widget(TitleLabel(text="Pago"))
        pro_card.add_widget(MutedLabel(text="Lancamentos ilimitados\nMulticontas\nRelatorios, alertas, cores e temas"))
        pro = PrimaryButton(text="Usar pago")
        pro_card.add_widget(pro)
        free.bind(on_release=lambda *_: self.set_plan("free"))
        pro.bind(on_release=lambda *_: self.set_plan("pro"))
        box.add_widget(free_card)
        box.add_widget(pro_card)
        self.add_widget(box)

    def set_plan(self, plan):
        usuario = self.user()
        settings = get_settings(usuario)
        save_settings(usuario, settings["moeda"], settings["tema"], plan, settings["alerta_limite"])
        show_message("Plano", f"Plano alterado para {plan}.")


class SettingsScreen(BaseScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
        with box.canvas.before:
            Color(*COLORS["bg"])
            self.settings_bg = RoundedRectangle(pos=box.pos, size=box.size, radius=[0])
        box.bind(pos=lambda *_: setattr(self.settings_bg, "pos", box.pos), size=lambda *_: setattr(self.settings_bg, "size", box.size))
        box.add_widget(self.nav_bar())
        box.add_widget(TitleLabel(text="Configuracoes", font_size=dp(22), size_hint_y=None, height=dp(40)))
        card = RoundedBox(orientation="vertical", padding=dp(16), spacing=dp(10))
        self.currency = AppSpinner(text="BRL", values=CURRENCIES)
        self.theme = AppSpinner(text="ocean", values=FREE_THEMES)
        self.alert = AppTextInput(hint_text="Limite de alerta", multiline=False)
        save = PrimaryButton(text="Salvar")
        save.bind(on_release=self.save)
        card.add_widget(MutedLabel(text="Moeda"))
        card.add_widget(self.currency)
        card.add_widget(MutedLabel(text="Tema"))
        card.add_widget(self.theme)
        card.add_widget(MutedLabel(text="Alerta de gasto (pago)"))
        card.add_widget(self.alert)
        card.add_widget(save)
        box.add_widget(card)
        self.add_widget(box)

    def refresh(self):
        settings = get_settings(self.user())
        self.currency.values = CURRENCIES if settings["plano"] == "pro" else ["BRL"]
        self.theme.values = PAID_THEMES if settings["plano"] == "pro" else FREE_THEMES
        self.currency.text = settings["moeda"] if settings["moeda"] in self.currency.values else "BRL"
        self.theme.text = settings["tema"] if settings["tema"] in self.theme.values else "ocean"
        self.alert.text = str(settings["alerta_limite"])

    def save(self, *_):
        settings = get_settings(self.user())
        try:
            alert_value = parse_amount(self.alert.text) if self.alert.text else 0
        except ValueError:
            alert_value = 0
        save_settings(self.user(), self.currency.text, self.theme.text, settings["plano"], alert_value)
        show_message("Pronto", "Configuracoes salvas.")


class FinanceiroOfflineApp(App):
    usuario = None

    def build(self):
        criar_tabelas()
        manager = ScreenManager()
        manager.add_widget(LoginScreen(name="login"))
        manager.add_widget(DashboardScreen(name="dashboard"))
        manager.add_widget(OrganizationScreen(name="organization"))
        manager.add_widget(PlansScreen(name="plans"))
        manager.add_widget(SettingsScreen(name="settings"))
        return manager


if __name__ == "__main__":
    FinanceiroOfflineApp().run()
