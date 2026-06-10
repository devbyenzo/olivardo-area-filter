# Autor: Enzo
# Projeto: Leitura e filtro de uma planilha xlsx

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd

# ── Tema ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Paleta de cores ───────────────────────────────────────────────────────────
BG_APP      = "#0F1117"   # fundo geral
BG_SIDEBAR  = "#161B27"   # sidebar
BG_CARD     = "#1C2333"   # cards e cabeçalho de tabela
BG_ROW_ALT  = "#1A2030"   # linha alternada
BG_ROW      = "#161B27"   # linha normal
ACCENT      = "#4F8EF7"   # azul principal
ACCENT_DIM  = "#1E3A6E"   # azul escuro (hover / badge bg)
GOLD        = "#F59E0B"   # estrelas
BORDER      = "#2A3347"   # divisores
TEXT_PRI    = "#E8EDF5"   # texto principal
TEXT_SEC    = "#6B7A99"   # texto secundário
TEXT_MUTED  = "#3D4F6E"   # muted

# Larguras das colunas da tabela
COL_WIDTHS  = {"Nome": 160, "Email": 190, "Área": 130, "Nota": 90, "Comentário": 240}
COLS_ORDER  = ["Nome", "Email", "Área", "Nota", "Comentário"]

# ── Estado global ─────────────────────────────────────────────────────────────
df_global   = None
df_filtrado = None
star_value  = 0

# ── Janela ────────────────────────────────────────────────────────────────────
root = ctk.CTk()
root.title("Olivardo Avaliações - Filtro de Planilha")
root.geometry("1080x660")
root.resizable(False, False)
root.configure(fg_color=BG_APP)

# ── Funções ───────────────────────────────────────────────────────────────────
def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecionar planilha",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if not caminho:
        return
    global df_global
    try:
        df_global = pd.read_excel(caminho)
        # Lê áreas únicas dinamicamente e atualiza o OptionMenu
        areas = sorted(df_global["Área"].dropna().unique().tolist())
        filtro_area.configure(values=["todas as áreas"] + areas)
        filtro_area.set("todas as áreas")

        nome = caminho.split("/")[-1]
        label_arquivo.configure(text=nome, text_color=TEXT_PRI)
        btn_arquivo.configure(text="  trocar arquivo")
        aplicar_filtros()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{e}")


def set_star(valor):
    global star_value
    star_value = 0 if star_value == valor else valor
    for i, btn in enumerate(botoes_estrelas):
        ativo = (i + 1) <= star_value
        btn.configure(
            text_color=GOLD if ativo else TEXT_MUTED,
            fg_color=ACCENT_DIM if ativo else "transparent"
        )
    aplicar_filtros()


def aplicar_filtros():
    global df_filtrado
    if df_global is None:
        return
    df = df_global.copy()

    area = filtro_area.get()
    if area != "todas as áreas":
        df = df[df["Área"] == area]

    if star_value > 0:
        df = df[df["Nota"] >= star_value]

    df_filtrado = df
    atualizar_stats(df)
    atualizar_tabela(df)


def salvar_filtrado():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
        return
    caminho = filedialog.asksaveasfilename(
        title="Salvar planilha filtrada",
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")]
    )
    if not caminho:
        return
    try:
        df_filtrado.to_excel(caminho, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho}")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível salvar:\n{e}")


def atualizar_stats(df):
    n = len(df)
    label_badge.configure(text=f"{n} resultado{'s' if n != 1 else ''}")
    label_total.configure(text=str(n))
    if n > 0:
        label_media.configure(text=f"{df['Nota'].mean():.1f}")
        label_areas.configure(text=str(df["Área"].nunique()))
    else:
        label_media.configure(text="—")
        label_areas.configure(text="—")


def atualizar_tabela(df):
    for w in frame_tabela.winfo_children():
        w.destroy()

    if df.empty:
        ctk.CTkLabel(
            frame_tabela,
            text="Nenhum resultado.\nAjuste os filtros.",
            text_color=TEXT_SEC,
            font=("Arial", 13),
            justify="center",
            fg_color="transparent"
        ).pack(pady=60)
        return

    # Cabeçalho
    cab = ctk.CTkFrame(frame_tabela, fg_color=BG_CARD, corner_radius=0, height=36)
    cab.pack(fill="x")
    cab.pack_propagate(False)
    for i, col in enumerate(COLS_ORDER):
        ctk.CTkLabel(
            cab,
            text=col.upper(),
            width=COL_WIDTHS[col],
            font=("Arial", 10, "bold"),
            text_color=TEXT_MUTED,
            anchor="w",
            fg_color="transparent"
        ).grid(row=0, column=i, padx=(16 if i == 0 else 8), pady=0, sticky="w")

    # Separador
    ctk.CTkFrame(frame_tabela, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")

    for idx, (_, row) in enumerate(df.iterrows()):
        bg = BG_ROW if idx % 2 == 0 else BG_ROW_ALT
        linha = ctk.CTkFrame(frame_tabela, fg_color=bg, corner_radius=0, height=42)
        linha.pack(fill="x")
        linha.pack_propagate(False)

        for i, col in enumerate(COLS_ORDER):
            val = row[col]

            if col == "Nota":
                try:
                    n = int(val)
                    texto = "★" * n + "☆" * (5 - n)
                    cor = GOLD
                except Exception:
                    texto = str(val)
                    cor = TEXT_PRI
            elif col == "Área":
                texto = "" if pd.isna(val) else str(val)
                cor = ACCENT
            else:
                texto = "" if pd.isna(val) else str(val)
                cor = TEXT_PRI

            ctk.CTkLabel(
                linha,
                text=texto,
                width=COL_WIDTHS[col],
                font=("Arial", 12),
                text_color=cor,
                anchor="w",
                fg_color="transparent"
            ).grid(row=0, column=i, padx=(16 if i == 0 else 8), pady=0, sticky="w")

        # linha divisória
        ctk.CTkFrame(frame_tabela, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")


# ═══════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ═══════════════════════════════════════════════════════════════════════════════
frame_principal = ctk.CTkFrame(root, fg_color="transparent")
frame_principal.pack(fill="both", expand=True, padx=0, pady=0)

# ── Sidebar ────────────────────────────────────────────────────────────────────
sidebar = ctk.CTkFrame(frame_principal, width=230, fg_color=BG_SIDEBAR, corner_radius=0)
sidebar.pack(side="left", fill="y")
sidebar.pack_propagate(False)

# Logo
frame_logo = ctk.CTkFrame(sidebar, fg_color="transparent")
frame_logo.pack(fill="x", padx=20, pady=(24, 20))

ctk.CTkLabel(
    frame_logo,
    text="Olivardo Avaliações",
    font=("Arial", 17, "bold"),
    text_color=TEXT_PRI
).pack(anchor="w")
ctk.CTkLabel(
    frame_logo,
    text="Leitor & Filtro de Planilha",
    font=("Arial", 11),
    text_color=TEXT_SEC
).pack(anchor="w")

ctk.CTkFrame(sidebar, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")

# Seção: arquivo
def label_section(parent, texto):
    ctk.CTkLabel(
        parent, text=texto,
        font=("Arial", 9, "bold"),
        text_color=TEXT_MUTED,
        fg_color="transparent"
    ).pack(anchor="w", padx=20, pady=(18, 6))

label_section(sidebar, "ARQUIVO")

btn_arquivo = ctk.CTkButton(
    sidebar,
    text="  Selecionar Arquivo",
    command=selecionar_arquivo,
    fg_color="transparent",
    border_width=1,
    border_color=BORDER,
    text_color=TEXT_SEC,
    hover_color=BG_CARD,
    height=36,
    corner_radius=8,
    font=("Arial", 12)
)
btn_arquivo.pack(fill="x", padx=20, pady=(0, 6))

label_arquivo = ctk.CTkLabel(
    sidebar,
    text="Nenhum arquivo selecionado",
    font=("Arial", 10),
    text_color=TEXT_MUTED,
    fg_color="transparent",
    wraplength=185,
    justify="left"
)
label_arquivo.pack(anchor="w", padx=20, pady=(0, 4))

ctk.CTkFrame(sidebar, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x", pady=(10, 0))

# Seção: área
label_section(sidebar, "FILTRAR POR ÁREA")

filtro_area = ctk.CTkOptionMenu(
    sidebar,
    values=["todas as áreas", "senhor-vinho", "sitio", "salao-principal"],
    command=lambda _: aplicar_filtros(),
    fg_color=BG_CARD,
    button_color=BORDER,
    button_hover_color=ACCENT_DIM,
    text_color=TEXT_PRI,
    dropdown_fg_color=BG_CARD,
    dropdown_text_color=TEXT_PRI,
    dropdown_hover_color=ACCENT_DIM,
    height=36,
    corner_radius=8,
    font=("Arial", 12)
)
filtro_area.pack(fill="x", padx=20, pady=(0, 4))

ctk.CTkFrame(sidebar, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x", pady=(10, 0))

# Seção: avaliação mínima
label_section(sidebar, "NOTA MÍNIMA")

frame_estrelas = ctk.CTkFrame(sidebar, fg_color="transparent")
frame_estrelas.pack(anchor="w", padx=18)

botoes_estrelas = []
for i in range(5):
    b = ctk.CTkButton(
        frame_estrelas,
        text="★",
        width=34,
        height=34,
        fg_color="transparent",
        hover_color=BG_CARD,
        text_color=TEXT_MUTED,
        font=("Arial", 17),
        corner_radius=6,
        command=lambda v=i + 1: set_star(v)
    )
    b.pack(side="left", padx=2)
    botoes_estrelas.append(b)

# Botão limpar filtros
def limpar_filtros():
    global star_value
    star_value = 0
    for btn in botoes_estrelas:
        btn.configure(text_color=TEXT_MUTED, fg_color="transparent")
    filtro_area.set("todas as áreas")
    aplicar_filtros()

ctk.CTkFrame(sidebar, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x", pady=(16, 0))

btn_limpar = ctk.CTkButton(
    sidebar,
    text="Limpar Filtros",
    command=limpar_filtros,
    fg_color="transparent",
    text_color=TEXT_SEC,
    hover_color=BG_CARD,
    height=34,
    corner_radius=8,
    font=("Arial", 12),
    border_width=0
)
btn_limpar.pack(fill="x", padx=20, pady=(10, 0))

btn_salvar = ctk.CTkButton(
    sidebar,
    text="Salvar Arquivo Filtrado",
    command=salvar_filtrado,
    fg_color="transparent",
    border_width=1,
    border_color=BORDER,
    text_color=TEXT_SEC,
    hover_color=BG_CARD,
    height=36,
    corner_radius=8,
    font=("Arial", 12)
)
btn_salvar.pack(fill="x", padx=20, pady=(8, 0), side="bottom")

btn_aplicar = ctk.CTkButton(
    sidebar,
    text="Aplicar Filtros",
    command=aplicar_filtros,
    fg_color=ACCENT,
    hover_color="#3A7AE4",
    text_color="#ffffff",
    height=38,
    corner_radius=8,
    font=("Arial", 13, "bold")
)
btn_aplicar.pack(fill="x", padx=20, pady=(6, 24), side="bottom")

# ── Main ───────────────────────────────────────────────────────────────────────
main = ctk.CTkFrame(frame_principal, fg_color=BG_APP, corner_radius=0)
main.pack(side="left", fill="both", expand=True)

# Topbar
topbar = ctk.CTkFrame(main, fg_color=BG_SIDEBAR, corner_radius=0, height=58)
topbar.pack(fill="x")
topbar.pack_propagate(False)

ctk.CTkLabel(
    topbar,
    text="Resultados",
    font=("Arial", 14, "bold"),
    text_color=TEXT_PRI,
    fg_color="transparent"
).pack(side="left", padx=24, pady=0)

label_badge = ctk.CTkLabel(
    topbar,
    text="— Resultados",
    font=("Arial", 11),
    text_color=TEXT_SEC,
    fg_color=BG_CARD,
    corner_radius=6,
    padx=10,
    pady=3
)
label_badge.pack(side="right", padx=20)

ctk.CTkFrame(main, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")

# Cards de stats
frame_stats = ctk.CTkFrame(main, fg_color="transparent", height=82)
frame_stats.pack(fill="x", padx=20, pady=16)
frame_stats.pack_propagate(False)

def criar_stat(parent, rotulo, valor_inicial="—"):
    card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10)
    card.pack(side="left", expand=True, fill="both", padx=5)
    ctk.CTkLabel(card, text=rotulo, font=("Arial", 10), text_color=TEXT_SEC, fg_color="transparent").pack(anchor="w", padx=14, pady=(10, 2))
    lbl = ctk.CTkLabel(card, text=valor_inicial, font=("Arial", 22, "bold"), text_color=TEXT_PRI, fg_color="transparent")
    lbl.pack(anchor="w", padx=14, pady=(0, 10))
    return lbl

label_total  = criar_stat(frame_stats, "TOTAL DE REGISTROS")
label_media  = criar_stat(frame_stats, "MÉDIA DE NOTA (★)")
label_areas  = criar_stat(frame_stats, "AREAS ENCONTRADAS")

ctk.CTkFrame(main, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x", padx=20)

# Tabela com scroll
frame_scroll = ctk.CTkScrollableFrame(main, fg_color="transparent", scrollbar_button_color=BORDER)
frame_scroll.pack(fill="both", expand=True, padx=20, pady=(12, 16))

frame_tabela = ctk.CTkFrame(frame_scroll, fg_color="transparent", corner_radius=0)
frame_tabela.pack(fill="both", expand=True)

# Estado inicial
ctk.CTkLabel(
    frame_tabela,
    text="Selecione um arquivo .xlsx para começar.",
    text_color=TEXT_SEC,
    font=("Arial", 13),
    fg_color="transparent"
).pack(pady=60)

root.mainloop()