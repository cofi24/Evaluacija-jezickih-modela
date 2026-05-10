"""
Generisanje chartova za evaluaciju LLM-ova na srpskom jeziku
=============================================================
Čita Excel sa odgovorima modela i generiše vizualizacije.

Instalacija:
    pip install openpyxl matplotlib seaborn pandas numpy

Pokretanje:
    python pravi_chartove.py
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib import rcParams
import seaborn as sns
from pathlib import Path
import openpyxl
from openpyxl import load_workbook

# ─── Podešavanja ─────────────────────────────────────────────────────────────

# Promeni naziv fajla u tvoj stvarni izlazni fajl sa odgovorima
EXCEL_FAJL = ""  # ostaviti prazno = automatski pronalazi najnoviji fajl

IZLAZNI_FOLDER = "chartovi"  # folder gde se čuvaju slike

# Nazivi modela i boje
MODELI = {
    "Odg. Mistral": {"naziv": "Mistral Small",  "boja": "#FF7043"},
    "Odg. GPT-4":   {"naziv": "GPT-4",           "boja": "#10A37F"},
    "Odg. Claude":  {"naziv": "Claude Sonnet",  "boja": "#CC785C"},
}

# Kolone u sheetu Tačnost koje sadrže DA/NE ocenu
# Format: { naziv_modela: naziv_kolone_sa_DA_NE }
# Ako imaš jednu zajedničku kolonu "Tačno?" podesi ovde
# Kolone za ocenu tacnosti (DA/NE) po modelu
TACNOST_KOLONE = {
    "Mistral Small": "Tačno_Mistral (DA/NE)",
    "GPT-4":         "Tačno_GPT (DA/NE)",
    "Claude Sonnet": "Tačno_Claude (DA/NE)",
}

# Kolone za ocenu halucinacija (DA/NE) po modelu
HALUCINACIJE_KOLONE = {
    "Mistral Small": "Da li je halucinirao Mistral? (DA/NE)",
    "GPT-4":         "Da li je halucinirao GPT? (DA/NE)",
    "Claude Sonnet": "Da li je halucinirao Claude? (DA/NE)",
}

rcParams["font.family"] = "DejaVu Sans"
rcParams["axes.spines.top"]   = False
rcParams["axes.spines.right"]  = False

# ─── Boje i stil ─────────────────────────────────────────────────────────────
BOJA_POZADINE = "#0F1117"
BOJA_TEKSTA   = "#E8E8E8"
BOJA_GRID     = "#2A2A3A"
PALETE        = ["#4285F4", "#FF7043", "#10A37F", "#CC785C"]

plt.rcParams.update({
    "figure.facecolor":  BOJA_POZADINE,
    "axes.facecolor":    "#1A1A2E",
    "axes.labelcolor":   BOJA_TEKSTA,
    "xtick.color":       BOJA_TEKSTA,
    "ytick.color":       BOJA_TEKSTA,
    "text.color":        BOJA_TEKSTA,
    "grid.color":        BOJA_GRID,
    "grid.linewidth":    0.6,
    "axes.titlesize":    13,
    "axes.titleweight":  "bold",
    "axes.labelsize":    10,
})


# ═════════════════════════════════════════════════════════════════════════════
# UČITAVANJE PODATAKA
# ═════════════════════════════════════════════════════════════════════════════

def pronadji_excel() -> str:
    """Automatski pronalazi najnoviji Excel fajl sa odgovorima."""
    if EXCEL_FAJL and Path(EXCEL_FAJL).exists():
        return EXCEL_FAJL

    fajlovi = list(Path(".").glob("*sa_odgovorima*.xlsx"))
    if not fajlovi:
        fajlovi = list(Path(".").glob("*.xlsx"))

    if not fajlovi:
        print("❌ Nije pronađen Excel fajl! Postavi EXCEL_FAJL na vrhu skripte.")
        sys.exit(1)

    najnoviji = max(fajlovi, key=lambda f: f.stat().st_mtime)
    print(f"✅ Koristim fajl: {najnoviji}")
    return str(najnoviji)


def ucitaj_sheet(putanja: str, sheet: str) -> pd.DataFrame:
    """Učitava sheet iz Excela kao DataFrame."""
    try:
        df = pd.read_excel(putanja, sheet_name=sheet, header=1)
        df = df.dropna(how="all")
        return df
    except Exception as e:
        print(f"⚠️  Sheet '{sheet}' nije učitan: {e}")
        return pd.DataFrame()


def pronadji_kolone_modela(df: pd.DataFrame) -> dict:
    """Automatski pronalazi kolone modela u DataFrame-u."""
    rezultat = {}
    for col in df.columns:
        col_str = str(col).strip()
        for kljuc, info in MODELI.items():
            if kljuc.lower() in col_str.lower() or info["naziv"].lower() in col_str.lower():
                rezultat[info["naziv"]] = col_str
    return rezultat


def izracunaj_tacnost(df: pd.DataFrame, kolone_modela: dict) -> dict:
    """Racuna tacnost na osnovu specificnih DA/NE kolona po modelu."""
    tacnost = {}
    for naziv, col in kolone_modela.items():
        # Uzmi tacnu DA/NE kolonu za ovaj model
        da_ne_col = TACNOST_KOLONE.get(naziv)
        if da_ne_col and da_ne_col in df.columns:
            ocene = df[da_ne_col].dropna().astype(str).str.strip().str.upper()
            if ocene.empty:
                continue
            procenat = (ocene == "DA").sum() / len(ocene) * 100
        elif col in df.columns:
            # Fallback: procenat odgovora koji nisu greška
            odgovori = df[col].dropna().astype(str)
            if odgovori.empty:
                continue
            procenat = (~odgovori.str.startswith("GREŠKA")).sum() / len(odgovori) * 100
        else:
            continue
        tacnost[naziv] = round(procenat, 1)
    return tacnost


def izracunaj_halucinacije(df: pd.DataFrame, kolone_modela: dict) -> dict:
    """Racuna stopu halucinacija na osnovu specificnih DA/NE kolona po modelu."""
    hal_stopa = {}
    for naziv, col in kolone_modela.items():
        da_ne_col = HALUCINACIJE_KOLONE.get(naziv)
        if da_ne_col and da_ne_col in df.columns:
            ocene = df[da_ne_col].dropna().astype(str).str.strip().str.upper()
            if ocene.empty:
                continue
            stopa = (ocene == "DA").sum() / len(ocene) * 100
        else:
            stopa = 0.0
        hal_stopa[naziv] = round(stopa, 1)
    return hal_stopa


# Kolone za konzistentnost (DA/NE) po modelu
KONZISTENTNOST_KOLONE = {
    "Mistral Small": "Konzistentno Mistral? (DA/NE)",
    "GPT-4":         "Konzistentno GPT? (DA/NE)",
    "Claude Sonnet": "Konzistentno Claude? (DA/NE)",
}

def izracunaj_konzistentnost(df: pd.DataFrame, kolone_modela: dict) -> dict:
    """Racuna konzistentnost na osnovu specificnih DA/NE kolona po modelu."""
    konzistentnost = {}
    for naziv, col in kolone_modela.items():
        da_ne_col = KONZISTENTNOST_KOLONE.get(naziv)
        if da_ne_col and da_ne_col in df.columns:
            ocene = df[da_ne_col].dropna().astype(str).str.strip().str.upper()
            if ocene.empty:
                continue
            stopa = (ocene == "DA").sum() / len(ocene) * 100
        else:
            stopa = 0.0
        konzistentnost[naziv] = round(stopa, 1)
    return konzistentnost


# ═════════════════════════════════════════════════════════════════════════════
# CHARTOVI
# ═════════════════════════════════════════════════════════════════════════════

def sacuvaj(fig, naziv: str):
    """Čuva chart u folder."""
    Path(IZLAZNI_FOLDER).mkdir(exist_ok=True)
    putanja = f"{IZLAZNI_FOLDER}/{naziv}.png"
    fig.savefig(putanja, dpi=150, bbox_inches="tight",
                facecolor=BOJA_POZADINE)
    print(f"  💾 Sačuvan: {putanja}")
    plt.close(fig)


def chart_uporedni_bar(podaci: dict, naslov: str, naziv_fajla: str,
                        y_oznaka: str = "%", boje: list = None):
    """Horizontalni bar chart za poređenje modela."""
    if not podaci:
        print(f"  ⚠️  Nema podataka za: {naslov}")
        return

    modeli  = list(podaci.keys())
    vrednosti = list(podaci.values())
    boje = boje or PALETE[:len(modeli)]

    fig, ax = plt.subplots(figsize=(9, max(3, len(modeli) * 1.2)))
    bars = ax.barh(modeli, vrednosti, color=boje[:len(modeli)],
                   height=0.55, zorder=2)

    for bar, val in zip(bars, vrednosti):
        ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2,
                f"{val}{y_oznaka}", va="center", ha="left",
                fontsize=11, fontweight="bold", color=BOJA_TEKSTA)

    ax.set_xlim(0, max(vrednosti) * 1.18 if vrednosti else 100)
    ax.set_xlabel(y_oznaka, labelpad=8)
    ax.set_title(naslov, pad=14)
    ax.grid(axis="x", zorder=1)
    ax.invert_yaxis()
    fig.tight_layout()
    sacuvaj(fig, naziv_fajla)


def chart_radar(sve_metrike: dict, naziv_fajla: str):
    """Radar/spider chart za sve tri dimenzije."""
    kategorije = ["Tačnost", "Konzistentnost", "Bez halucinacija"]
    N = len(kategorije)
    uglovi = [n / float(N) * 2 * np.pi for n in range(N)]
    uglovi += uglovi[:1]

    fig, ax = plt.subplots(figsize=(7, 7),
                           subplot_kw=dict(polar=True),
                           facecolor=BOJA_POZADINE)
    ax.set_facecolor("#1A1A2E")
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_xticks(uglovi[:-1])
    ax.set_xticklabels(kategorije, size=11, color=BOJA_TEKSTA)
    ax.set_ylim(0, 100)
    ax.yaxis.set_tick_params(labelcolor=BOJA_GRID)
    ax.grid(color=BOJA_GRID, linewidth=0.7)
    ax.spines["polar"].set_color(BOJA_GRID)

    for i, (naziv, metrike) in enumerate(sve_metrike.items()):
        vrednosti = [
            metrike.get("tacnost", 0),
            metrike.get("konzistentnost", 0),
            100 - metrike.get("halucinacije", 0),
        ]
        vrednosti += vrednosti[:1]
        boja = PALETE[i % len(PALETE)]
        ax.plot(uglovi, vrednosti, linewidth=2, color=boja, label=naziv)
        ax.fill(uglovi, vrednosti, alpha=0.12, color=boja)

    ax.legend(loc="upper right", bbox_to_anchor=(1.35, 1.15),
              framealpha=0.2, labelcolor=BOJA_TEKSTA, fontsize=10)
    ax.set_title("Sveobuhvatna evaluacija modela", pad=20,
                 color=BOJA_TEKSTA, fontsize=14, fontweight="bold")
    fig.tight_layout()
    sacuvaj(fig, naziv_fajla)


def chart_heatmapa_kategorija(df_tacnost: pd.DataFrame,
                               kolone_modela: dict, naziv_fajla: str):
    """Heatmapa tačnosti po kategorijama."""
    kat_col = None
    for c in df_tacnost.columns:
        if "kateg" in str(c).lower():
            kat_col = c
            break

    if not kat_col or not kolone_modela:
        print("  ⚠️  Nema kategorija za heatmapu.")
        return

    # Napravi matricu: kategorija × model
    matrica = {}
    for naziv, col in kolone_modela.items():
        if col not in df_tacnost.columns:
            continue
        grupe = df_tacnost.groupby(kat_col)[col].apply(
            lambda x: (~x.dropna().astype(str).str.startswith("GREŠKA")).mean() * 100
        )
        matrica[naziv] = grupe

    if not matrica:
        return

    df_mat = pd.DataFrame(matrica).fillna(0)

    fig, ax = plt.subplots(figsize=(max(6, len(matrica) * 2), max(4, len(df_mat) * 0.7)))
    sns.heatmap(df_mat, ax=ax, annot=True, fmt=".0f", cmap="YlOrRd",
                linewidths=0.5, linecolor=BOJA_POZADINE,
                cbar_kws={"label": "Tačnost (%)"})
    ax.set_title("Tačnost po kategorijama pitanja", pad=14)
    ax.set_xlabel("")
    ax.set_ylabel("")
    plt.xticks(rotation=20, ha="right")
    plt.yticks(rotation=0)
    fig.tight_layout()
    sacuvaj(fig, naziv_fajla)


def chart_dashboard(tacnost: dict, halucinacije: dict,
                    konzistentnost: dict, naziv_fajla: str):
    """Glavni dashboard sa svim metrikama na jednoj slici."""
    modeli = list(set(list(tacnost.keys()) +
                      list(halucinacije.keys()) +
                      list(konzistentnost.keys())))
    if not modeli:
        return

    fig = plt.figure(figsize=(16, 9), facecolor=BOJA_POZADINE)
    gs  = gridspec.GridSpec(2, 3, figure=fig,
                            hspace=0.45, wspace=0.35,
                            left=0.07, right=0.97,
                            top=0.88, bottom=0.1)

    boje_map = {m: PALETE[i % len(PALETE)] for i, m in enumerate(modeli)}

    def mini_bar(ax, podaci, naslov, obrnuto=False):
        if not podaci:
            ax.set_visible(False)
            return
        m = list(podaci.keys())
        v = list(podaci.values())
        b = [boje_map.get(x, "#888") for x in m]
        bars = ax.barh(m, v, color=b, height=0.5, zorder=2)
        for bar, val in zip(bars, v):
            ax.text(bar.get_width() + 0.5,
                    bar.get_y() + bar.get_height() / 2,
                    f"{val}%", va="center", fontsize=9,
                    fontweight="bold", color=BOJA_TEKSTA)
        ax.set_xlim(0, 110)
        ax.set_title(naslov, fontsize=11, pad=8)
        ax.grid(axis="x", zorder=1)
        ax.invert_yaxis()
        if obrnuto:
            ax.invert_xaxis()

    # 1. Tačnost
    ax1 = fig.add_subplot(gs[0, 0])
    mini_bar(ax1, tacnost, "✅ Tačnost (%)")

    # 2. Halucinacije
    ax2 = fig.add_subplot(gs[0, 1])
    mini_bar(ax2, halucinacije, "🔴 Stopa halucinacija (%)")

    # 3. Konzistentnost
    ax3 = fig.add_subplot(gs[0, 2])
    mini_bar(ax3, konzistentnost, "🔁 Konzistentnost (%)")

    # 4. Rang tabela (donji levi)
    ax4 = fig.add_subplot(gs[1, :2])
    ax4.set_facecolor("#1A1A2E")
    rang_podaci = []
    for m in modeli:
        t = tacnost.get(m, 0)
        h = 100 - halucinacije.get(m, 0)
        k = konzistentnost.get(m, 0)
        ukupno = round((t + h + k) / 3, 1)
        rang_podaci.append((m, t, halucinacije.get(m, 0), k, ukupno))

    rang_podaci.sort(key=lambda x: x[4], reverse=True)

    kolone_tab = ["Model", "Tačnost", "Halucinacije", "Konzistentnost", "Ukupno"]
    tabela_vrednosti = [[r[0], f"{r[1]}%", f"{r[2]}%", f"{r[3]}%", f"{r[4]}%"]
                        for r in rang_podaci]

    tab = ax4.table(cellText=tabela_vrednosti,
                    colLabels=kolone_tab,
                    cellLoc="center", loc="center",
                    bbox=[0, 0, 1, 1])
    tab.auto_set_font_size(False)
    tab.set_fontsize(10)
    for (row, col), cell in tab.get_celld().items():
        cell.set_facecolor("#1A1A2E" if row > 0 else "#2A2A4E")
        cell.set_text_props(color=BOJA_TEKSTA)
        cell.set_edgecolor(BOJA_GRID)
        if row == 1:  # pobednik
            cell.set_facecolor("#2A3A2A")
    ax4.set_title("📊 Rang tabela modela", fontsize=11, pad=10)
    ax4.axis("off")

    # 5. Radar (donji desni)
    ax5 = fig.add_subplot(gs[1, 2], polar=True)
    ax5.set_facecolor("#1A1A2E")
    kategorije_r = ["Tačnost", "Konzist.", "Bez hal."]
    N = len(kategorije_r)
    uglovi = [n / float(N) * 2 * np.pi for n in range(N)]
    uglovi += uglovi[:1]
    ax5.set_theta_offset(np.pi / 2)
    ax5.set_theta_direction(-1)
    ax5.set_xticks(uglovi[:-1])
    ax5.set_xticklabels(kategorije_r, size=8, color=BOJA_TEKSTA)
    ax5.set_ylim(0, 100)
    ax5.grid(color=BOJA_GRID, linewidth=0.6)
    ax5.spines["polar"].set_color(BOJA_GRID)
    ax5.yaxis.set_tick_params(labelcolor=BOJA_GRID, labelsize=6)
    for naziv in modeli:
        v = [tacnost.get(naziv, 0),
             konzistentnost.get(naziv, 0),
             100 - halucinacije.get(naziv, 0)]
        v += v[:1]
        boja = boje_map[naziv]
        ax5.plot(uglovi, v, linewidth=1.5, color=boja, label=naziv)
        ax5.fill(uglovi, v, alpha=0.1, color=boja)
    ax5.set_title("Radar", fontsize=10, pad=12)

    fig.suptitle("Evaluacija LLM-ova na srpskom jeziku — Rezultati",
                 fontsize=16, fontweight="bold", color=BOJA_TEKSTA, y=0.96)

    sacuvaj(fig, naziv_fajla)


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    putanja = pronadji_excel()

    print("\n📊 Učitavam podatke iz Excela...")
    df_tacnost      = ucitaj_sheet(putanja, "Tačnost")
    df_halucinacije = ucitaj_sheet(putanja, "Halucinacije")
    df_konzist      = ucitaj_sheet(putanja, "Konzistentnost")

    print("\n🔍 Pronalazim kolone modela...")
    kol_t = pronadji_kolone_modela(df_tacnost)
    kol_h = pronadji_kolone_modela(df_halucinacije)
    kol_k = pronadji_kolone_modela(df_konzist)

    print(f"   Tačnost:       {list(kol_t.keys())}")
    print(f"   Halucinacije:  {list(kol_h.keys())}")
    print(f"   Konzistentnost:{list(kol_k.keys())}")

    print("\n📈 Računam metrike...")
    tacnost        = izracunaj_tacnost(df_tacnost, kol_t)
    halucinacije   = izracunaj_halucinacije(df_halucinacije, kol_h)
    konzistentnost = izracunaj_konzistentnost(df_konzist, kol_k)

    print(f"   Tačnost:        {tacnost}")
    print(f"   Halucinacije:   {halucinacije}")
    print(f"   Konzistentnost: {konzistentnost}")

    print(f"\n🎨 Generišem chartove u folder '{IZLAZNI_FOLDER}/'...")

    # Chart 1 — Tačnost
    chart_uporedni_bar(tacnost, "Tačnost modela (%)", "01_tacnost")

    # Chart 2 — Halucinacije
    chart_uporedni_bar(halucinacije, "Stopa halucinacija (%)", "02_halucinacije")

    # Chart 3 — Konzistentnost
    chart_uporedni_bar(konzistentnost, "Konzistentnost modela (%)", "03_konzistentnost")

    # Chart 4 — Heatmapa po kategorijama
    chart_heatmapa_kategorija(df_tacnost, kol_t, "04_heatmapa_kategorija")

    # Chart 5 — Radar
    sve_metrike = {}
    for naziv in set(list(tacnost.keys()) + list(halucinacije.keys()) + list(konzistentnost.keys())):
        sve_metrike[naziv] = {
            "tacnost":        tacnost.get(naziv, 0),
            "halucinacije":   halucinacije.get(naziv, 0),
            "konzistentnost": konzistentnost.get(naziv, 0),
        }
    chart_radar(sve_metrike, "05_radar")

    # Chart 6 — Dashboard (sve na jednoj slici)
    chart_dashboard(tacnost, halucinacije, konzistentnost, "06_dashboard")

    print(f"\n✅ Sve gotovo! Chartovi su u folderu: ./{IZLAZNI_FOLDER}/")
    print("   Možeš ih direktno ubaciti u diplomski rad.")
    print("\n--- DEBUG ---")
    print("Tačnost:", tacnost)
    print("Halucinacije:", halucinacije)
    print("Konzistentnost:", konzistentnost)
    print("-------------")

if __name__ == "__main__":
    
    main()
    