# wt_stat.py

import os
import glob
from datetime import datetime

import pandas as pd
import matplotlib
matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from PyQt5.QtWidgets import (
    QWidget, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QComboBox, QTableWidget, QTableWidgetItem, QDateEdit,
    QListWidget, QListWidgetItem
)

from PyQt5.QtCore import Qt, QDate


# --------------------------
# FUNZIONI DI SUPPORTO
# --------------------------

def load_all_csv(folder):
    files = glob.glob(os.path.join(folder, "*.csv"))
    if not files:
        raise FileNotFoundError("Nessun file .csv trovato nella cartella.")

    df_list = []

    for f in files:
        try:
            df = pd.read_csv(
                f,
                sep=";",
                header=None,
                engine="python",
                dtype=str,
                skip_blank_lines=True
            )

            # Rimuove righe completamente vuote
            df = df.dropna(how="all")

            # Controlla che ci siano 7 colonne
            if df.shape[1] != 7:
                print(f"File {f} ignorato: colonne trovate = {df.shape[1]}")
                continue

            # Scarta l’ultima colonna vuota
            df = df.iloc[:, :6]

            df.columns = [
                "projectname", "date_start", "time_start",
                "date_end", "time_end", "timeelapsed"
            ]

            df_list.append(df)

        except Exception as e:
            print(f"Errore nel file {f}: {e}")
            continue

    if not df_list:
        raise ValueError("Nessun file valido trovato.")

    return pd.concat(df_list, ignore_index=True)






def preprocess(df):
    # Normalizza stringhe (toglie spazi, None, ecc.)
    for col in ["date_start", "time_start", "date_end", "time_end", "timeelapsed"]:
        df[col] = df[col].astype(str).str.strip()

    # Costruisce le stringhe datetime
    start_str = df["date_start"] + " " + df["time_start"]
    end_str   = df["date_end"]   + " " + df["time_end"]

    # Converte in datetime, ma senza esplodere: le righe errate diventano NaT
    df["start"] = pd.to_datetime(
        start_str,
        format="%d/%m/%Y %H:%M",
        errors="coerce"
    )
    df["end"] = pd.to_datetime(
        end_str,
        format="%d/%m/%Y %H:%M",
        errors="coerce"
    )

    # Converte la durata, righe non numeriche diventano NaN
    df["duration"] = pd.to_numeric(df["timeelapsed"], errors="coerce")

    # Rimuove tutte le righe con problemi gravi
    before = len(df)
    df = df.dropna(subset=["start", "end", "duration"])
    after = len(df)

    if after < before:
        print(f"Righe scartate per dati non validi: {before - after}")

    # Ordina per data di inizio
    df = df.sort_values("start").reset_index(drop=True)
    return df



def compute_weekly_stats(df):
    weekly = df.set_index("start").resample("W")["duration"].sum().reset_index()
    weekly["week"] = weekly["start"].dt.strftime("%Y-%W")
    return weekly[["week", "duration"]]


def compute_monthly_stats(df):
    monthly = df.set_index("start").resample("M")["duration"].sum().reset_index()
    monthly["month"] = monthly["start"].dt.strftime("%Y-%m")
    return monthly[["month", "duration"]]


def productivity_analysis(df):
    per_project = df.groupby("projectname")["duration"].sum().reset_index()
    per_project = per_project.sort_values("duration", ascending=False)

    per_day = df.set_index("start").resample("D")["duration"].sum().reset_index()
    avg_per_day = per_day["duration"].mean() if not per_day.empty else 0.0

    return per_project, avg_per_day


def generate_pdf_report(df, output_path):
    weekly = compute_weekly_stats(df)
    monthly = compute_monthly_stats(df)
    per_project, avg_per_day = productivity_analysis(df)

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    y = height - 50

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Report Attività / Produttività")
    y -= 40

    c.setFont("Helvetica", 10)
    c.drawString(50, y, f"Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore medie per giorno:")
    c.setFont("Helvetica", 10)
    c.drawString(200, y, f"{avg_per_day:.2f}")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per progetto")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in per_project.iterrows():
        c.drawString(60, y, f"{row['projectname']}: {row['duration']:.2f} h")
        y -= 15
        if y < 80:
            c.showPage()
            y = height - 50

    c.showPage()
    y = height - 50

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per settimana")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in weekly.iterrows():
        c.drawString(60, y, f"{row['week']}: {row['duration']:.2f} h")
        y -= 15

    c.showPage()
    y = height - 50

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per mese")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in monthly.iterrows():
        c.drawString(60, y, f"{row['month']}: {row['duration']:.2f} h")
        y -= 15

    c.save()


# --------------------------
# CANVAS MATPLOTLIB
# --------------------------

class MplCanvas(FigureCanvas):
    def __init__(self):
        fig = Figure(figsize=(5, 3), dpi=100)
        self.axes = fig.add_subplot(111)
        super().__init__(fig)


# --------------------------
# FINESTRA PRINCIPALE
# --------------------------

class WT_Stat(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Statistiche attività")
        self.resize(1400, 900)

        self.df = None
        self.filtered_df = None

        layout = QVBoxLayout(self)

        # --- CONTROLLI SUPERIORI ---
        top = QHBoxLayout()

        self.folder_label = QLabel("Cartella: (nessuna)")
        btn_folder = QPushButton("Apri cartella CSV")
        btn_folder.clicked.connect(self.load_folder)

        # Lista progetti da escludere
        self.exclude_list = QListWidget()
        self.exclude_list.setSelectionMode(QListWidget.MultiSelection)
        self.exclude_list.itemSelectionChanged.connect(self.apply_filters)

        # Date
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.dateChanged.connect(self.apply_filters)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.dateChanged.connect(self.apply_filters)

        btn_pdf = QPushButton("Esporta PDF")
        btn_pdf.clicked.connect(self.export_pdf)

        top.addWidget(self.folder_label)
        top.addWidget(btn_folder)
        top.addWidget(QLabel("Escludi progetti"))
        top.addWidget(self.exclude_list)
        top.addWidget(QLabel("Da"))
        top.addWidget(self.start_date)
        top.addWidget(QLabel("A"))
        top.addWidget(self.end_date)
        top.addWidget(btn_pdf)

        layout.addLayout(top)

        # --- GRAFICI ---
        self.canvas_act = MplCanvas()
        self.canvas_week = MplCanvas()
        self.canvas_month = MplCanvas()

        layout.addWidget(QLabel("Attività"))
        layout.addWidget(self.canvas_act, 2)
        layout.addWidget(QLabel("Statistiche settimanali"))
        layout.addWidget(self.canvas_week, 2)
        layout.addWidget(QLabel("Statistiche mensili"))
        layout.addWidget(self.canvas_month, 2)

    # --------------------------
    # LOGICA
    # --------------------------

    def load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleziona cartella CSV")
        if not folder:
            return

        try:
            df = load_all_csv(folder)
            df = preprocess(df)
            self.df = df
            self.filtered_df = df.copy()

            self.folder_label.setText(f"Cartella: {folder}")

            # Popola lista progetti
            self.exclude_list.clear()
            for p in sorted(df["projectname"].unique()):
                item = QListWidgetItem(p)
                item.setCheckState(Qt.Unchecked)
                self.exclude_list.addItem(item)

            # Imposta automaticamente l'anno corrente
            year = datetime.now().year
            start = QDate(year, 1, 1)
            end = QDate(year, 12, 31)
            self.start_date.setDate(start)
            self.end_date.setDate(end)

            self.apply_filters()

        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def apply_filters(self):
        if self.df is None:
            return

        df = self.df.copy()

        # Esclusione progetti
        excluded = [
            self.exclude_list.item(i).text()
            for i in range(self.exclude_list.count())
            if self.exclude_list.item(i).checkState() == Qt.Checked
        ]

        if excluded:
            df = df[~df["projectname"].isin(excluded)]

        # Filtro date
        s = self.start_date.date()
        e = self.end_date.date()
        s_dt = datetime(s.year(), s.month(), s.day())
        e_dt = datetime(e.year(), e.month(), e.day(), 23, 59, 59)

        df = df[(df["start"] >= s_dt) & (df["start"] <= e_dt)]

        self.filtered_df = df
        self.refresh_plots()

    def refresh_plots(self):
        df = self.filtered_df

        # Attività
        self.canvas_act.axes.clear()
        if df is not None and not df.empty:
            self.canvas_act.axes.bar(df["start"], df["duration"])
            self.canvas_act.axes.set_title("Tempo per attività")
            self.canvas_act.axes.tick_params(axis="x", labelrotation=45)
        self.canvas_act.draw()

        # Settimanali
        self.canvas_week.axes.clear()
        if df is not None and not df.empty:
            w = compute_weekly_stats(df)
            self.canvas_week.axes.bar(w["week"], w["duration"])
            self.canvas_week.axes.set_title("Ore per settimana")
            self.canvas_week.axes.tick_params(axis="x", labelrotation=45)
        self.canvas_week.draw()

        # Mensili
        self.canvas_month.axes.clear()
        if df is not None and not df.empty:
            m = compute_monthly_stats(df)
            self.canvas_month.axes.bar(m["month"], m["duration"])
            self.canvas_month.axes.set_title("Ore per mese")
            self.canvas_month.axes.tick_params(axis="x", labelrotation=45)
        self.canvas_month.draw()

    def export_pdf(self):
        if self.filtered_df is None or self.filtered_df.empty:
            QMessageBox.warning(self, "Attenzione", "Nessun dato da esportare.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Salva PDF", "", "PDF (*.pdf)")
        if not path:
            return

        try:
            generate_pdf_report(self.filtered_df, path)
            QMessageBox.information(self, "OK", f"PDF salvato in:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
