"""
export.py
---------
Génère un fichier Excel formaté à partir des résultats de prédiction batch.
"""

import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from datetime import date


# ── Couleurs ───────────────────────────────────────────────────────────────────
BLUE_FILL   = PatternFill("solid", fgColor="EEF4FF")
CORAL_FILL  = PatternFill("solid", fgColor="FDF2F2")
HEADER_FILL = PatternFill("solid", fgColor="111928")
WARN_FILL   = PatternFill("solid", fgColor="FFFBEB")
GRAY_FILL   = PatternFill("solid", fgColor="F9FAFB")

THIN_BORDER = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB"),
)


def _header_style(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = THIN_BORDER


def _data_style(cell, fill=None):
    cell.font = Font(size=10, name="Calibri")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill


def export_to_excel(results: list) -> io.BytesIO:
    """
    Crée un fichier Excel formaté avec les résultats de prédiction.
    Retourne un buffer BytesIO prêt à être envoyé.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Prédictions L-D"

    # ── Titre ──────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = f"Terminal L/D Predictor — Export du {date.today().strftime('%d/%m/%Y')}"
    title_cell.font  = Font(bold=True, size=13, color="111928", name="Calibri")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Sous-titre ─────────────────────────────────────────────────────────────
    ws.merge_cells("A2:I2")
    sub = ws["A2"]
    sub.value = "Ratios basés sur les 3 derniers mois d'activité par lane"
    sub.font  = Font(size=10, color="6B7280", name="Calibri")
    sub.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    ws.row_dimensions[3].height = 8  # Espace

    # ── En-têtes ───────────────────────────────────────────────────────────────
    headers = [
        "Lane", "Volume navire",
        "Load estimé", "% Load",
        "Discharge estimé", "% Discharge",
        "Période de référence", "N escales", "Statut"
    ]
    for col, h in enumerate(headers, start=1):
        _header_style(ws.cell(row=4, column=col), h)
    ws.row_dimensions[4].height = 22

    # ── Données ────────────────────────────────────────────────────────────────
    for row_idx, res in enumerate(results, start=5):
        unknown  = res.get("unknown", False)
        outdated = res.get("outdated", False)

        # Lane
        c = ws.cell(row=row_idx, column=1, value=res["lane"])
        c.font = Font(bold=True, size=10, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

        # Volume
        c = ws.cell(row=row_idx, column=2, value=res["volume"])
        _data_style(c)
        c.number_format = "#,##0"

        if unknown:
            # Lane inconnue → cellules grises
            for col in range(3, 10):
                c = ws.cell(row=row_idx, column=col, value="—" if col != 9 else "Lane inconnue")
                _data_style(c, GRAY_FILL)
        else:
            # Load
            c = ws.cell(row=row_idx, column=3, value=res["pred_L"])
            _data_style(c, BLUE_FILL)
            c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="1E40AF", name="Calibri")

            # % Load
            c = ws.cell(row=row_idx, column=4, value=res["pct_L"] / 100)
            _data_style(c, BLUE_FILL)
            c.number_format = "0.0%"
            c.font = Font(size=10, color="1E40AF", name="Calibri")

            # Discharge
            c = ws.cell(row=row_idx, column=5, value=res["pred_D"])
            _data_style(c, CORAL_FILL)
            c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="991B1B", name="Calibri")

            # % Discharge
            c = ws.cell(row=row_idx, column=6, value=res["pct_D"] / 100)
            _data_style(c, CORAL_FILL)
            c.number_format = "0.0%"
            c.font = Font(size=10, color="991B1B", name="Calibri")

            # Période
            period = f"{res['period_start']} → {res['last_date']}"
            c = ws.cell(row=row_idx, column=7, value=period)
            _data_style(c)
            c.alignment = Alignment(horizontal="left", vertical="center")

            # N escales
            c = ws.cell(row=row_idx, column=8, value=res["n_voyages"])
            _data_style(c)

            # Statut
            if outdated:
                c = ws.cell(row=row_idx, column=9, value="⚠ Données anciennes")
                _data_style(c, WARN_FILL)
                c.font = Font(size=10, color="92400E", name="Calibri")
            else:
                c = ws.cell(row=row_idx, column=9, value="OK")
                _data_style(c)
                c.font = Font(size=10, color="057A55", name="Calibri")

        ws.row_dimensions[row_idx].height = 20

    # ── Largeurs des colonnes ──────────────────────────────────────────────────
    col_widths = [12, 16, 16, 10, 18, 13, 30, 12, 20]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Freeze panes (figer l'en-tête) ────────────────────────────────────────
    ws.freeze_panes = "A5"

    # ── Sauvegarde dans buffer ─────────────────────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ── Couleurs F/M (ajoutées pour la fonctionnalité Full/Empty) ─────────────────
GREEN_FILL = PatternFill("solid", fgColor="ECFDF5")
AMBER_FILL = PatternFill("solid", fgColor="FFFBEB")


def export_fm_to_excel(results: list) -> io.BytesIO:
    """
    Exporte les résultats du pipeline complet (D/L + F/M) en Excel formaté.
    Colonnes : Lane, Volume, Discharge, Load, D-Full, D-Empty, L-Full, L-Empty.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Prédictions L-D + F-M"

    ws.merge_cells("A1:J1")
    t = ws["A1"]
    t.value = f"Terminal Predictor — L/D + Full/Empty — Export du {date.today().strftime('%d/%m/%Y')}"
    t.font  = Font(bold=True, size=13, color="111928", name="Calibri")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    s = ws["A2"]
    s.value = "Pipeline : Volume → Load/Discharge (ratios D/L) → Full/Empty (ratios F/M, 3 derniers mois)"
    s.font  = Font(size=10, color="6B7280", name="Calibri")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    headers = ["Lane", "Volume", "Discharge", "Load",
               "D-Full", "% D-Full", "D-Empty", "% D-Empty",
               "L-Full", "% L-Full", "L-Empty", "% L-Empty"]
    for col, h in enumerate(headers, start=1):
        _header_style(ws.cell(row=4, column=col), h)
    ws.row_dimensions[4].height = 22

    for row_idx, res in enumerate(results, start=5):
        unknown  = res.get("unknown", False)
        outdated = res.get("outdated", False)
        fm_ok    = res.get("fm_available", False)

        c = ws.cell(row=row_idx, column=1, value=res["lane"])
        c.font = Font(bold=True, size=10, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

        c = ws.cell(row=row_idx, column=2, value=res["volume"])
        _data_style(c); c.number_format = "#,##0"

        if unknown:
            for col in range(3, 13):
                c = ws.cell(row=row_idx, column=col, value="Lane inconnue" if col == 3 else "—")
                _data_style(c, GRAY_FILL)
        else:
            # Discharge total
            c = ws.cell(row=row_idx, column=3, value=res.get("pred_D"))
            _data_style(c, CORAL_FILL); c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="991B1B", name="Calibri")

            # Load total
            c = ws.cell(row=row_idx, column=4, value=res.get("pred_L"))
            _data_style(c, BLUE_FILL); c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="1E40AF", name="Calibri")

            if not fm_ok:
                for col in range(5, 13):
                    c = ws.cell(row=row_idx, column=col, value="N/D")
                    _data_style(c, GRAY_FILL)
            else:
                data_fm = [
                    (res.get("D_full"),  res.get("pct_D_full"),  CORAL_FILL, "991B1B"),
                    (res.get("D_empty"), res.get("pct_D_empty"), AMBER_FILL, "92400E"),
                    (res.get("L_full"),  res.get("pct_L_full"),  BLUE_FILL,  "1E40AF"),
                    (res.get("L_empty"), res.get("pct_L_empty"), GREEN_FILL, "065F46"),
                ]
                col = 5
                for val, pct, fill, color in data_fm:
                    c = ws.cell(row=row_idx, column=col, value=val)
                    _data_style(c, fill); c.number_format = "#,##0"
                    c.font = Font(bold=True, size=10, color=color, name="Calibri")
                    col += 1
                    c = ws.cell(row=row_idx, column=col, value=(pct or 0) / 100)
                    _data_style(c, fill); c.number_format = "0.0%"
                    c.font = Font(size=10, color=color, name="Calibri")
                    col += 1

        ws.row_dimensions[row_idx].height = 20

    for i, w in enumerate([10, 10, 12, 12, 10, 9, 10, 9, 10, 9, 10, 9], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A5"
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_teu_to_excel(results: list) -> io.BytesIO:
    """
    Exporte les résultats du pipeline complet (D/L + F/M + TEU) en Excel.
    Colonnes : Lane, Volume, D, L, D-Full, D-Empty, L-Full, L-Empty,
               D-Full TEU, D-Empty TEU, L-Full TEU, L-Empty TEU, Total TEU.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Prédictions TEU Complet"

    # Couleurs TEU
    teu_blue_fill  = PatternFill("solid", fgColor="DBEAFE")   # L-Full TEU
    teu_green_fill = PatternFill("solid", fgColor="D1FAE5")   # L-Empty TEU
    teu_red_fill   = PatternFill("solid", fgColor="FFE4E6")   # D-Full TEU
    teu_amber_fill = PatternFill("solid", fgColor="FEF3C7")   # D-Empty TEU
    total_fill     = PatternFill("solid", fgColor="EDE9FE")   # Total TEU

    ws.merge_cells("A1:M1")
    t = ws["A1"]
    t.value = f"Terminal Predictor — D/L + F/M + TEU — Export du {date.today().strftime('%d/%m/%Y')}"
    t.font  = Font(bold=True, size=13, color="111928", name="Calibri")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:M2")
    s = ws["A2"]
    s.value = "Pipeline : Volume → D/L → Full/Empty (containers) → TEU (20': ×1 · 40': ×2)"
    s.font  = Font(size=10, color="6B7280", name="Calibri")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    headers = [
        "Lane", "Volume",
        "Discharge", "Load",
        "D-Full", "D-Empty", "L-Full", "L-Empty",
        "D-Full TEU", "D-Empty TEU", "L-Full TEU", "L-Empty TEU",
        "Total TEU"
    ]
    for col, h in enumerate(headers, start=1):
        _header_style(ws.cell(row=4, column=col), h)
    ws.row_dimensions[4].height = 22

    for row_idx, res in enumerate(results, start=5):
        unknown  = res.get("unknown", False)
        outdated = res.get("outdated", False)
        fm_ok    = res.get("fm_available", False)
        teu_ok   = res.get("teu_available", False)

        c = ws.cell(row=row_idx, column=1, value=res.get("lane", "?"))
        c.font = Font(bold=True, size=10, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

        c = ws.cell(row=row_idx, column=2, value=res.get("volume"))
        _data_style(c); c.number_format = "#,##0"

        if unknown:
            for col in range(3, 14):
                c = ws.cell(row=row_idx, column=col, value="Lane inconnue" if col == 3 else "—")
                _data_style(c, GRAY_FILL)
        else:
            # D et L containers total
            c = ws.cell(row=row_idx, column=3, value=res.get("pred_D"))
            _data_style(c, CORAL_FILL); c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="991B1B", name="Calibri")

            c = ws.cell(row=row_idx, column=4, value=res.get("pred_L"))
            _data_style(c, BLUE_FILL); c.number_format = "#,##0"
            c.font = Font(bold=True, size=10, color="1E40AF", name="Calibri")

            # F/M containers
            fm_vals = [
                (res.get("D_full"),  CORAL_FILL, "991B1B"),
                (res.get("D_empty"), WARN_FILL,  "92400E"),
                (res.get("L_full"),  BLUE_FILL,  "1E40AF"),
                (res.get("L_empty"), GREEN_FILL, "065F46"),
            ]
            for i, (val, fill, color) in enumerate(fm_vals, start=5):
                c = ws.cell(row=row_idx, column=i, value=val if fm_ok else "N/D")
                _data_style(c, fill if fm_ok else GRAY_FILL)
                if fm_ok and val is not None:
                    c.number_format = "#,##0"
                    c.font = Font(bold=True, size=10, color=color, name="Calibri")

            # TEU par groupe
            teu_vals = [
                (res.get("D_full_teu"),  teu_red_fill,   "7F1D1D"),
                (res.get("D_empty_teu"), teu_amber_fill, "78350F"),
                (res.get("L_full_teu"),  teu_blue_fill,  "1E3A8A"),
                (res.get("L_empty_teu"), teu_green_fill, "065F46"),
            ]
            for i, (val, fill, color) in enumerate(teu_vals, start=9):
                c = ws.cell(row=row_idx, column=i, value=val if teu_ok else "N/D")
                _data_style(c, fill if teu_ok else GRAY_FILL)
                if teu_ok and val is not None:
                    c.number_format = "#,##0"
                    c.font = Font(bold=True, size=10, color=color, name="Calibri")

            # Total TEU
            c = ws.cell(row=row_idx, column=13, value=res.get("total_teu") if teu_ok else "N/D")
            _data_style(c, total_fill if teu_ok else GRAY_FILL)
            if teu_ok and res.get("total_teu") is not None:
                c.number_format = "#,##0"
                c.font = Font(bold=True, size=11, color="4C1D95", name="Calibri")

        ws.row_dimensions[row_idx].height = 20

    for i, w in enumerate([10, 10, 12, 12, 10, 10, 10, 10, 12, 13, 12, 13, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A5"
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
