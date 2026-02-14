import os
import glob
import threading
import webbrowser
from datetime import datetime

import pandas as pd

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import dash
from dash import dcc, html
from dash.dependencies import Input, Output

import plotly.express as px

import tkinter as tk
from tkinter import filedialog, messagebox


# =========================
# CARICAMENTO E PREPROCESS
# =========================

def load_all_csv(folder):
    files = glob.glob(os.path.join(folder, "*.csv"))
    if not files:
        raise FileNotFoundError("Nessun file .csv trovato nella cartella.")

    df_list = []
    for f in files:
        # Adatta sep / header se necessario
        df = pd.read_csv(f, sep="\t", header=None)
        df.columns = [
            "projectname", "date_start", "time_start",
            "date_end", "time_end", "timeelapsed"
        ]
        df_list.append(df)

    df = pd.concat(df_list, ignore_index=True)
    return df


def preprocess(df):
    df["start"] = pd.to_datetime(
        df["date_start"] + " " + df["time_start"],
        format="%d/%m/%Y %H:%M"
    )
    df["end"] = pd.to_datetime(
        df["date_end"] + " " + df["time_end"],
        format="%d/%m/%Y %H:%M"
    )
    df["duration"] = df["timeelapsed"].astype(float)
    df = df.sort_values("start")
    return df


# =========================
# STATISTICHE E ANALISI
# =========================

def compute_weekly_stats(df):
    # Somma ore per settimana
    weekly = (
        df.set_index("start")
          .resample("W")["duration"]
          .sum()
          .reset_index()
    )
    weekly["week"] = weekly["start"].dt.strftime("%Y-%W")
    return weekly[["week", "duration"]]


def compute_monthly_stats(df):
    monthly = (
        df.set_index("start")
          .resample("M")["duration"]
          .sum()
          .reset_index()
    )
    monthly["month"] = monthly["start"].dt.strftime("%Y-%m")
    return monthly[["month", "duration"]]


def productivity_analysis(df):
    # Esempio semplice: ore per progetto e ore medie per giorno
    per_project = (
        df.groupby("projectname")["duration"]
          .sum()
          .reset_index()
          .sort_values("duration", ascending=False)
    )

    per_day = (
        df.set_index("start")
          .resample("D")["duration"]
          .sum()
          .reset_index()
    )
    avg_per_day = per_day["duration"].mean() if not per_day.empty else 0.0

    return per_project, avg_per_day


# =========================
# PDF REPORT
# =========================

def generate_pdf_report(df, output_path):
    weekly = compute_weekly_stats(df)
    monthly = compute_monthly_stats(df)
    per_project, avg_per_day = productivity_analysis(df)

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Report Attività / Produttività")
    y -= 30

    c.setFont("Helvetica", 10)
    c.drawString(50, y, f"Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 20

    # Produttività generale
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Produttività generale")
    y -= 20
    c.setFont("Helvetica", 10)
    c.drawString(60, y, f"Ore medie per giorno: {avg_per_day:.2f}")
    y -= 30

    # Ore per progetto
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per progetto")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in per_project.iterrows():
        line = f"{row['projectname']}: {row['duration']:.2f} h"
        c.drawString(60, y, line)
        y -= 15
        if y < 80:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 10)

    # Nuova pagina per settimanali/mensili
    c.showPage()
    y = height - 50

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per settimana")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in weekly.iterrows():
        line = f"Settimana {row['week']}: {row['duration']:.2f} h"
        c.drawString(60, y, line)
        y -= 15
        if y < 80:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 10)

    c.showPage()
    y = height - 50
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Ore per mese")
    y -= 20
    c.setFont("Helvetica", 10)
    for _, row in monthly.iterrows():
        line = f"Mese {row['month']}: {row['duration']:.2f} h"
        c.drawString(60, y, line)
        y -= 15
        if y < 80:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 10)

    c.save()


# =========================
# DASH APP
# =========================

def create_dash_app(df):
    app = dash.Dash(__name__)

    app.layout = html.Div([
        html.H1("Time Tracking Dashboard"),

        html.Div([
            html.Label("Filtra per progetto"),
            dcc.Dropdown(
                id="project-filter",
                options=[
                    {"label": p, "value": p}
                    for p in sorted(df["projectname"].unique())
                ],
                multi=True,
                placeholder="Seleziona uno o più progetti"
            ),
        ], style={"width": "40%", "display": "inline-block"}),

        html.Div([
            html.Label("Intervallo date"),
            dcc.DatePickerRange(
                id="date-filter",
                start_date=df["start"].min(),
                end_date=df["start"].max()
            ),
        ], style={"width": "40%", "display": "inline-block", "marginLeft": "40px"}),

        dcc.Tabs(id="tabs", value="tab-activities", children=[
            dcc.Tab(label="Attività", value="tab-activities"),
            dcc.Tab(label="Statistiche settimanali", value="tab-weekly"),
            dcc.Tab(label="Statistiche mensili", value="tab-monthly"),
            dcc.Tab(label="Produttività", value="tab-productivity"),
        ]),

        html.Div(id="tab-content")
    ])

    @app.callback(
        Output("tab-content", "children"),
        [
            Input("tabs", "value"),
            Input("project-filter", "value"),
            Input("date-filter", "start_date"),
            Input("date-filter", "end_date"),
        ]
    )
    def render_tab(tab, projects, start_date, end_date):
        filtered = df.copy()

        if projects:
            filtered = filtered[filtered["projectname"].isin(projects)]

        if start_date:
            filtered = filtered[filtered["start"] >= pd.to_datetime(start_date)]
        if end_date:
            filtered = filtered[filtered["start"] <= pd.to_datetime(end_date)]

        if filtered.empty:
            return html.Div("Nessun dato per i filtri selezionati.")

        if tab == "tab-activities":
            fig = px.bar(
                filtered,
                x="start",
                y="duration",
                color="projectname",
                title="Tempo per attività",
                hover_data=["projectname", "date_start", "time_start", "duration"]
            )
            return dcc.Graph(figure=fig)

        elif tab == "tab-weekly":
            weekly = compute_weekly_stats(filtered)
            fig = px.bar(
                weekly,
                x="week",
                y="duration",
                title="Ore per settimana"
            )
            return dcc.Graph(figure=fig)

        elif tab == "tab-monthly":
            monthly = compute_monthly_stats(filtered)
            fig = px.bar(
                monthly,
                x="month",
                y="duration",
                title="Ore per mese"
            )
            return dcc.Graph(figure=fig)

        elif tab == "tab-productivity":
            per_project, avg_per_day = productivity_analysis(filtered)
            fig = px.bar(
                per_project,
                x="projectname",
                y="duration",
                title=f"Ore per progetto (media giornaliera: {avg_per_day:.2f} h)",
            )
            fig.update_layout(xaxis_title="Progetto", yaxis_title="Ore")
            return dcc.Graph(figure=fig)

        return html.Div("Tab non riconosciuta.")

    return app


def run_dash_app(df):
    app = create_dash_app(df)
    # Avvia server su porta 8050
    app.run_server(debug=False, use_reloader=False)


# =========================
# INTERFACCIA DESKTOP (TK)
# =========================

class TimeTrackerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Tracking Dashboard Launcher")

        self.folder = tk.StringVar()
        self.df = None

        tk.Label(root, text="Cartella CSV:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(root, textvariable=self.folder, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Sfoglia", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)

        tk.Button(root, text="Carica dati", command=self.load_data).grid(row=1, column=0, padx=5, pady=10)
        tk.Button(root, text="Avvia dashboard", command=self.start_dashboard).grid(row=1, column=1, padx=5, pady=10)
        tk.Button(root, text="Esporta report PDF", command=self.export_pdf).grid(row=1, column=2, padx=5, pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder.set(folder_selected)

    def load_data(self):
        try:
            if not self.folder.get():
                messagebox.showwarning("Attenzione", "Seleziona una cartella.")
                return
            df = load_all_csv(self.folder.get())
            df = preprocess(df)
            self.df = df
            messagebox.showinfo("OK", f"Dati caricati. Righe totali: {len(df)}")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nel caricamento dati:\n{e}")

    def start_dashboard(self):
        if self.df is None:
            messagebox.showwarning("Attenzione", "Carica prima i dati.")
            return

        def run_server():
            run_dash_app(self.df)

        t = threading.Thread(target=run_server, daemon=True)
        t.start()

        webbrowser.open("http://127.0.0.1:8050")

    def export_pdf(self):
        if self.df is None:
            messagebox.showwarning("Attenzione", "Carica prima i dati.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            generate_pdf_report(self.df, output_path)
            messagebox.showinfo("OK", f"Report PDF salvato in:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nella generazione del PDF:\n{e}")


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    root = tk.Tk()
    gui = TimeTrackerGUI(root)
    root.mainloop()
