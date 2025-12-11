# Compara-planilhas
Comparar mais de uma coluna entre uma planilha e outra

<img width="1350" height="1022" alt="image" src="https://github.com/user-attachments/assets/342abc0e-6058-4e04-9743-06c4f24cd5d4" />


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import os
import tempfile
import shutil

# ====== Imports opcionais com fallback ======
try:
    import pandas as pd
except ImportError:
    pd = None

THEME = {
    "bg": "#1e1e1e",
    "card": "#2d2d2d",
    "fg": "#d4d4d4",
    "accent": "#569cd6",
    "accent2": "#4ec9b0",
    "success": "#73c990",
    "border": "#3c3c3c",
}

TITLE_TEXT = "Compare Sheets"
TITLE_FONT = ("Courier New", 20, "bold")

# ---------------- Helpers ----------------
def create_round_rectangle(canvas, x1, y1, x2, y2, radius=20, **kwargs):
    points = [
        x1 + radius, y1,
        x2 - radius, y1,
        x2, y1,
        x2, y1 + radius,
        x2, y2 - radius,
        x2, y2,
        x2 - radius, y2,
        x1 + radius, y2,
        x1, y2,
        x1, y2 - radius,
        x1, y1 + radius,
        x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)

def resolve_parent_bg(widget, fallback):
    try:
        return widget.cget("background")
    except Exception:
        return fallback

class PillButton(tk.Canvas):
    def __init__(self, master, text, width=220, height=48,
                 bg_fill=THEME["accent"], text_color=THEME["bg"],
                 command=None, parent_bg=None, variant="default", **kwargs):
        if parent_bg is None:
            parent_bg = resolve_parent_bg(master, THEME["card"])
        super().__init__(master, width=width, height=height,
                         highlightthickness=0, bg=parent_bg, **kwargs)
        self._bg_fill = bg_fill
        self._text_color = text_color
        self._text = text
        self._radius = height // 2
        self._command = command
        self._variant = variant
        self._render(self._bg_fill)
        if self._command is not None:
            self.bind("<Button-1>", lambda e: self._command())
            self.bind("<Enter>", lambda e: self._hover(True))
            self.bind("<Leave>", lambda e: self._hover(False))

    def _render(self, fill_color):
        self.delete("all")
        w = int(self["width"]); h = int(self["height"])
        create_round_rectangle(self, 2, 2, w - 2, h - 2,
                               radius=self._radius, fill=fill_color, outline="")
        if self._variant == "upload":
            self._draw_upload_content(w, h)
        else:
            self.create_text(w // 2, h // 2, text=self._text,
                             fill=self._text_color, font=("Segoe UI", 11, "bold"))

    def _draw_upload_content(self, w, h):
        icon_center_x = 22
        icon_center_y = h // 2
        icon_color = self._text_color
        arrow_width = 16
        arrow_height = 12
        arrow_points = [
            icon_center_x,                 icon_center_y - arrow_height // 2 - 1,
            icon_center_x - arrow_width//2, icon_center_y + arrow_height // 2 - 1,
            icon_center_x + arrow_width//2, icon_center_y + arrow_height // 2 - 1
        ]
        self.create_polygon(arrow_points, fill=icon_color, outline=icon_color)
        text_x = icon_center_x + 26
        self.create_text(text_x, h // 2, text=self._text,
                         fill=self._text_color, font=("Segoe UI", 11, "bold"), anchor="w")

    def _hover(self, on):
        if on:
            self._render(self._mix_color(self._bg_fill, "#ffffff", 0.12))
        else:
            self._render(self._bg_fill)

    def _mix_color(self, hex1, hex2, alpha):
        def h2rgb(h):
            h = h.lstrip("#")
            return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
        def rgb2h(rgb):
            return "#{:02x}{:02x}{:02x}".format(*rgb)
        c1 = h2rgb(hex1); c2 = h2rgb(hex2)
        mixed = tuple(min(255, max(0, int(c1[i]*(1-alpha) + c2[i]*alpha))) for i in range(3))
        return rgb2h(mixed)

class RoundedCard(tk.Canvas):
    def __init__(self, master, width, height, radius=26, fill=THEME["card"], **kwargs):
        parent_bg = resolve_parent_bg(master, THEME["bg"])
        super().__init__(master, width=width, height=height,
                         bg=parent_bg, highlightthickness=0, **kwargs)
        create_round_rectangle(self, 4, 4, width - 4, height - 4,
                               radius=radius, fill=fill, outline=THEME["border"])
        self.inner = tk.Frame(self, bg=fill)
        self.create_window(width // 2, height // 2, window=self.inner)

# ---------------- Worker (mesma lógica anterior: key Sheet1 + faixa sheet2) ----------------
class CompareWorker:
    def __init__(self, file1, file2, out_format, ui_queue,
                 key1, key2_opt, start2, end2):
        self.file1 = file1
        self.file2 = file2
        self.out_format = out_format.lower()
        self.ui_queue = ui_queue
        self.key1 = key1.strip()
        self.key2_opt = (key2_opt or '').strip()
        self.start2 = start2.strip()
        self.end2 = end2.strip()
        self.temp_output = None
        self._styles = None

    def run(self):
        try:
            if pd is None:
                raise ImportError("Pandas não está instalado. Instale com: pip install pandas openpyxl xlrd")
            self._emit_log("🟢 Iniciando comparação...")
            self._emit_progress(0)
            df1 = self._read_file(self.file1, name="Sheet1"); self._emit_progress(15)
            df2 = self._read_file(self.file2, name="sheet2"); self._emit_progress(30)
            df1 = self._normalize(df1); df2 = self._normalize(df2)
            self._emit_log("🔧 Normalização concluída."); self._emit_progress(45)
            result_df, styles = self._compare_policies(df1, df2)
            self._styles = styles
            self._emit_log("📊 Comparação concluída."); self._emit_progress(85)
            self.temp_output = self._write_output(result_df, self.out_format)
            self._emit_log(f"💾 Arquivo gerado: {self.temp_output}")
            self._emit_progress(100)
            self._emit_done(self.temp_output)
        except Exception as e:
            self._emit_error(str(e))

    def _emit_progress(self, pct): self.ui_queue.put({"type": "progress", "pct": int(pct)})
    def _emit_log(self, text): self.ui_queue.put({"type": "log", "text": text})
    def _emit_error(self, text): self.ui_queue.put({"type": "error", "text": text})
    def _emit_done(self, path): self.ui_queue.put({"type": "done", "path": path})

    def _read_file(self, path, name="Sheet"):
        ext = os.path.splitext(path)[1].lower()
        self._emit_log(f"📂 Lendo {name}: {os.path.basename(path)}")
        try:
            if ext == ".csv":
                return self._read_csv_with_fallback(path)
            elif ext in (".xlsx", ".xls"):
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                return pd.read_excel(path, engine=engine, dtype=str)
            else:
                raise ValueError(f"Formato não suportado: {ext}")
        except UnicodeDecodeError as ude:
            raise ValueError(f"Erro de codificação ao ler {name}: {ude}")
        except ImportError as ie:
            raise ImportError(f"Dependência ausente para Excel: {ie}. Instale com: pip install openpyxl xlrd")

    def _read_csv_with_fallback(self, path):
        encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1", "iso-8859-1", "utf-16"]
        last_err = None
        for enc in encodings:
            try:
                df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc, sep=None, engine='python')
                self._emit_log(f"✅ CSV carregado com encoding: {enc}")
                return df
            except Exception as e:
                last_err = e
        raise ValueError(f"Não foi possível ler o CSV. Tente salvar o arquivo como UTF-8. Último erro: {last_err}")

    def _normalize(self, df):
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        for c in df.columns:
            df[c] = df[c].astype(str).fillna("").str.strip()
        return df

    def _compare_policies(self, df1, df2):
        key1 = self.key1 or "Policy Name"
        key2 = self.key2_opt or key1
        start2 = self.start2
        end2 = self.end2

        if key1 not in df1.columns:
            raise ValueError(f"Sheet1 não possui a coluna-chave '{key1}'.")
        if key2 not in df2.columns:
            self._emit_log(f"⚠️ sheet2 não possui a coluna-chave '{key2}'. Será criada vazia.")
            df2[key2] = ""

        if start2 not in df2.columns:
            raise ValueError(f"sheet2 não possui a coluna inicial '{start2}'.")
        if end2 not in df2.columns:
            raise ValueError(f"sheet2 não possui a coluna final '{end2}'.")
        s_idx = list(df2.columns).index(start2)
        e_idx = list(df2.columns).index(end2)
        if e_idx < s_idx:
            raise ValueError("A coluna final em sheet2 aparece antes da coluna inicial.")
        compare_cols = list(df2.columns)[s_idx:e_idx+1]
        self._emit_log(f"🔎 Faixa de comparação (sheet2): {compare_cols}")

        for c in compare_cols:
            if c not in df1.columns:
                self._emit_log(f"⚠️ Sheet1 não possui a coluna '{c}'. Será considerada vazia na comparação.")
                df1[c] = ""

        out_rows = []
        styles = {}
        df2_index = {str(v): i for i, v in enumerate(df2[key2].tolist())}

        def parse_tokens(text):
            raw = [t for part in str(text).splitlines() for t in part.split(',')]
            return set([t.strip() for t in raw if t.strip()])

        for i, base_row in df1.iterrows():
            key_val = base_row.get(key1, "")
            out_row = {key1: key_val}
            status = "Igual"
            style_row_index = len(out_rows)

            match_idx = df2_index.get(str(key_val))
            target_row = df2.iloc[match_idx] if match_idx is not None else None

            if target_row is None:
                status = "Não existe."
                for c in compare_cols:
                    val = base_row.get(c, "")
                    out_row[c] = val
                    styles[(style_row_index, c)] = 'red'
                for c in df2.columns:
                    if c not in out_row:
                        out_row[c] = ""
            else:
                for c in compare_cols:
                    v1 = base_row.get(c, "")
                    v2 = target_row.get(c, "")
                    s1 = parse_tokens(v1)
                    s2 = parse_tokens(v2)
                    if s1 == s2:
                        out_row[c] = v1
                        styles[(style_row_index, c)] = 'green'
                    else:
                        out_row[c] = v1
                        styles[(style_row_index, c)] = 'red'
                        status = "Corrigir"
                for c in df2.columns:
                    if c not in out_row:
                        out_row[c] = target_row.get(c, "")

            out_row = {"Status": status, **out_row}
            out_rows.append(out_row)

        final_cols = ["Status", key1]
        final_cols += [c for c in compare_cols if c not in final_cols]
        final_cols += [c for c in df2.columns if c not in final_cols]
        result_df = pd.DataFrame(out_rows, columns=final_cols)
        self._emit_log(f"⚙️ Parâmetros: key1='{key1}', key2='{key2}', inicio(sheet2)='{start2}', fim(sheet2)='{end2}'")
        return result_df, styles

    def _write_output(self, df, out_format):
        temp_dir = tempfile.gettempdir()
        base_name = "comparison_report"
        if out_format == "csv":
            out_path = os.path.join(temp_dir, f"{base_name}.csv")
            df.to_csv(out_path, index=False)
        elif out_format == "xlsx":
            out_path = os.path.join(temp_dir, f"{base_name}.xlsx")
            self._write_xlsx_styled(df, self._styles or {}, out_path)
        elif out_format == "pdf":
            out_path = os.path.join(temp_dir, f"{base_name}.pdf")
            self._write_pdf(df, out_path, self._styles or {})
        else:
            raise ValueError("Formato inválido. Use: csv, xlsx ou pdf.")
        return out_path

    def _write_xlsx_styled(self, df, styles, out_path):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font
        except ImportError:
            raise ImportError("openpyxl não está instalado. Instale com: pip install openpyxl")
        wb = Workbook()
        ws = wb.active
        ws.title = "Comparison"
        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=j, value=col)
        for i, row in df.iterrows():
            excel_row = i + 2
            for j, col in enumerate(df.columns, start=1):
                val = row[col]
                cell = ws.cell(row=excel_row, column=j, value=val)
                style_key = (i, col)
                if style_key in styles:
                    color = styles[style_key]
                    if color == 'green':
                        cell.font = Font(color="008000")
                    elif color == 'red':
                        cell.font = Font(color="FF0000")
        wb.save(out_path)

    def _write_pdf(self, df, out_path, styles):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
        except ImportError:
            raise ImportError("Reportlab não está instalado. Instale com: pip install reportlab")
        styles_rl = getSampleStyleSheet()
        story = []
        doc = SimpleDocTemplate(out_path, pagesize=landscape(A4),
                                rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        title = Paragraph("<b>Comparison Report</b>", styles_rl['Title'])
        story.append(title); story.append(Spacer(1, 12))
        cols = list(df.columns)
        data = [cols] + df.values.tolist()
        table = Table(data, repeatRows=1)
        ts = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(THEME["accent"])),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor(THEME["bg"])),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor(THEME["border"]))
        ]
        for (row_idx, col_name), color in styles.items():
            j = cols.index(col_name)
            if color == 'green':
                ts.append(('TEXTCOLOR', (j, row_idx + 1), (j, row_idx + 1), colors.green))
            elif color == 'red':
                ts.append(('TEXTCOLOR', (j, row_idx + 1), (j, row_idx + 1), colors.red))
        table.setStyle(TableStyle(ts))
        story.append(table)
        doc.build(story)

# ---------------- App ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITLE_TEXT)
        self.geometry("1100x900")
        self.minsize(980, 780)
        self.configure(bg=THEME["bg"])  # fundo
        style = ttk.Style(self); style.theme_use("clam")
        style.configure("Dark.TLabel", background=THEME["bg"], foreground=THEME["fg"])  # labels dark
        style.configure("Card.TLabel", background=THEME["card"], foreground=THEME["fg"])  # labels dentro do card
        style.configure("Dark.TFrame", background=THEME["bg"])  # container dark
        style.configure("Card.TFrame", background=THEME["card"])  # card
        style.configure("Dark.TButton", background=THEME["accent"], foreground=THEME["bg"], borderwidth=0, padding=8)
        style.map(
            "Dark.TButton",
            background=[("active", THEME["accent2"])],
            foreground=[("active", THEME["bg"])],
        )
        style.configure("Dark.Horizontal.TProgressbar", troughcolor=THEME["card"], background=THEME["success"])  # barra
        style.configure("Dark.TCombobox", fieldbackground=THEME["card"], background=THEME["card"], foreground=THEME["fg"])  # combo
        style.map(
            "Dark.TCombobox",
            fieldbackground=[("readonly", THEME["card"])],
            foreground=[("readonly", THEME["fg"])],
        )
        # Estado
        self.file1 = None
        self.file2 = None
        self.var_ext = tk.StringVar(value="xlsx")
        self.var_key1 = tk.StringVar(value="Policy Name")
        self.var_key2_opt = tk.StringVar(value="")
        self.var_start2 = tk.StringVar(value="Source")
        self.var_end2 = tk.StringVar(value="Security Profiles")
        self.ui_queue = queue.Queue()
        self.worker_thread = None
        self._last_output_path = None

        self._build_ui()
        self.after(100, self._process_ui_queue)

    def _build_ui(self):
        # Header
        header = ttk.Frame(self, style="Dark.TFrame")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ttk.Label(header, text=TITLE_TEXT, style="Dark.TLabel", font=TITLE_FONT, anchor="center").pack(fill="x")

        # Card maior para acomodar tudo
        card = RoundedCard(self, width=1040, height=760, radius=26, fill=THEME["card"])
        card.pack(padx=16, pady=8)
        inner = card.inner

        # Empilhar seções verticalmente (uma abaixo da outra)
        # 1) Uploads
        uploads = ttk.Frame(inner, style="Card.TFrame")
        uploads.pack(fill="x", pady=(20, 8))
        top_labels = ttk.Frame(uploads, style="Card.TFrame")
        top_labels.pack()
        ttk.Label(top_labels, text="(Base) Sheet1", style="Card.TLabel").grid(row=0, column=0, padx=12)
        ttk.Label(top_labels, text="(Comparar) sheet2", style="Card.TLabel").grid(row=0, column=1, padx=12)
        btns = ttk.Frame(uploads, style="Card.TFrame")
        btns.pack()
        self.btn_sheet1 = PillButton(btns, text="Upload Sheet1", bg_fill=THEME["accent"], text_color=THEME["bg"],
                                     command=self._choose_file1, parent_bg=THEME["card"], variant="upload")
        self.btn_sheet2 = PillButton(btns, text="Upload sheet2", bg_fill=THEME["accent2"], text_color=THEME["bg"],
                                     command=self._choose_file2, parent_bg=THEME["card"], variant="upload")
        self.btn_sheet1.grid(row=0, column=0, padx=12, pady=6)
        self.btn_sheet2.grid(row=0, column=1, padx=12, pady=6)

        # 2) Parâmetros
        params = ttk.Frame(inner, style="Card.TFrame")
        params.pack(fill="x", padx=24, pady=(4, 8))
        for i in range(8): params.columnconfigure(i, weight=1)
        ttk.Label(params, text="Campo chave (Sheet1):", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(params, textvariable=self.var_key1).grid(row=0, column=1, sticky="ew", padx=(6,12))
        ttk.Label(params, text="Chave sheet2 (opcional):", style="Card.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(params, textvariable=self.var_key2_opt).grid(row=0, column=3, sticky="ew", padx=(6,12))
        ttk.Label(params, text="Início (sheet2):", style="Card.TLabel").grid(row=0, column=4, sticky="w")
        ttk.Entry(params, textvariable=self.var_start2).grid(row=0, column=5, sticky="ew", padx=(6,12))
        ttk.Label(params, text="Fim (sheet2):", style="Card.TLabel").grid(row=0, column=6, sticky="w")
        ttk.Entry(params, textvariable=self.var_end2).grid(row=0, column=7, sticky="ew")

        # 3) Botão iniciar
        actions_top = ttk.Frame(inner, style="Card.TFrame")
        actions_top.pack(fill="x", padx=24, pady=(0, 8))
        self.btn_start = ttk.Button(actions_top, text="Comparar", style="Dark.TButton",
                                    command=self._start_comparison, state="disabled")
        self.btn_start.pack(anchor="center")

        # 4) Progresso + status output
        progress_block = ttk.Frame(inner, style="Card.TFrame")
        progress_block.pack(fill="x", padx=24, pady=(4, 4))
        progress_block.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(progress_block, style="Dark.Horizontal.TProgressbar", mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.progress["maximum"] = 100
        self.progress["value"] = 0
        self.lbl_pct = ttk.Label(progress_block, text="0%", style="Card.TLabel")
        self.lbl_pct.grid(row=0, column=1, padx=(12, 0))

        output_frame = ttk.Frame(inner, style="Card.TFrame")
        output_frame.pack(fill="x", padx=24, pady=(2, 6))
        self.lbl_output = ttk.Label(output_frame, text="Output: (aguardando geração do arquivo)", style="Card.TLabel")
        self.lbl_output.pack(anchor="w")

        # 5) Log com scroll dentro do card
        log_container = ttk.Frame(inner, style="Card.TFrame")
        log_container.pack(fill="both", expand=True, padx=24, pady=(6, 12))
        # Scrollbars
        y_scroll = ttk.Scrollbar(log_container, orient="vertical")
        x_scroll = ttk.Scrollbar(log_container, orient="horizontal")
        self.txt_log = tk.Text(log_container, height=12, bg=THEME["card"], fg=THEME["fg"], relief="flat",
                               wrap='none', xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        y_scroll.config(command=self.txt_log.yview)
        x_scroll.config(command=self.txt_log.xview)
        # Grid positioning to make it resize nicely
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        # 6) Formato + ações (download/abrir pasta)
        bottom_frame = ttk.Frame(inner, style="Card.TFrame")
        bottom_frame.pack(fill="x", padx=24, pady=(0, 8))
        container = ttk.Frame(bottom_frame, style="Card.TFrame")
        container.pack(anchor="center")
        ttk.Label(container, text="Output format:", style="Card.TLabel").pack(anchor="center")
        self.cmb_ext = ttk.Combobox(container, textvariable=self.var_ext,
                                    values=["xlsx", "csv", "pdf"], state="readonly",
                                    width=10, style="Dark.TCombobox")
        self.cmb_ext.pack(pady=(6, 8))
        actions = ttk.Frame(container, style="Card.TFrame")
        actions.pack(pady=(0,8))
        self.btn_download = ttk.Button(actions, text="Download", style="Dark.TButton",
                                       state="disabled", command=self._download_output)
        self.btn_download.grid(row=0, column=0, padx=(0,8))
        self.btn_open_folder = ttk.Button(actions, text="Abrir pasta", style="Dark.TButton",
                                          state="disabled", command=self._open_output_folder)
        self.btn_open_folder.grid(row=0, column=1)

        # Footer
        footer = ttk.Frame(self, style="Dark.TFrame")
        footer.pack(fill="x", padx=16, pady=(8, 16))
        ttk.Label(footer, text="By @BlackMambaS", style="Dark.TLabel").pack(anchor="center")

    # ---------------- Eventos ----------------
    def _choose_file1(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha (Sheet1)",
            filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.file1 = path
            self._log(f"Sheet1: {os.path.basename(path)}")
            self._enable_start_if_ready()

    def _choose_file2(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha (sheet2)",
            filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.file2 = path
            self._log(f"sheet2: {os.path.basename(path)}")
            self._enable_start_if_ready()

    def _enable_start_if_ready(self):
        if self.file1 and self.file2:
            self.btn_start.config(state="normal")

    def _start_comparison(self):
        if pd is None:
            messagebox.showerror("Dependência ausente",
                                 "Pandas não está instalado.\n\nInstale com:\n  pip install pandas openpyxl xlrd")
            return
        out_format = self.var_ext.get().lower()
        if out_format == "pdf":
            try:
                import reportlab
            except ImportError:
                messagebox.showerror("Dependência ausente",
                                     "Reportlab não está instalado.\n\nInstale com:\n  pip install reportlab")
                return
        if out_format == "xlsx":
            try:
                import openpyxl
            except ImportError:
                messagebox.showerror("Dependência ausente",
                                     "openpyxl não está instalado.\n\nInstale com:\n  pip install openpyxl")
                return
        self.progress["value"] = 0
        self.lbl_pct.config(text="0%")
        self.btn_download.config(state="disabled")
        self.btn_open_folder.config(state="disabled")
        self.lbl_output.config(text="Output: processando…")
        self._last_output_path = None

        worker = CompareWorker(
            self.file1, self.file2, out_format, self.ui_queue,
            key1=self.var_key1.get(), key2_opt=self.var_key2_opt.get(),
            start2=self.var_start2.get(), end2=self.var_end2.get()
        )
        self.worker_thread = threading.Thread(target=worker.run, daemon=True)
        self.worker_thread.start()

    def _download_output(self):
        path = self._last_output_path
        self._log("🖫 Download acionado.")
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Arquivo não encontrado para download.")
            self._log("⚠️ _last_output_path ausente ou arquivo não existe.")
            return
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        initialdir = os.path.expanduser("~\\Downloads") if os.name == 'nt' else os.path.expanduser("~")
        save_path = filedialog.asksaveasfilename(
            title="Salvar relatório",
            initialdir=initialdir,
            defaultextension=f".{ext}",
            filetypes=[(ext.upper(), f"*.{ext}"), ("Todos os arquivos", "*.*")]
        )
        if save_path:
            try:
                shutil.copyfile(path, save_path)
                messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{save_path}")
                self._log(f"✅ Salvo em: {save_path}")
            except Exception as e:
                messagebox.showerror("Erro ao salvar", str(e))
                self._log(f"❌ Erro ao salvar: {e}")
        else:
            self._log("ℹ️ Salvamento cancelado pelo usuário.")

    def _open_output_folder(self):
        path = self._last_output_path
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Arquivo não encontrado.")
            return
        folder = os.path.dirname(path)
        try:
            if os.name == 'nt':
                os.startfile(folder)
            elif os.name == 'posix':
                import subprocess
                subprocess.Popen(['xdg-open', folder])
            else:
                messagebox.showinfo("Info", f"Pasta: {folder}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {e}")

    def _process_ui_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                if msg["type"] == "log":
                    self._log(msg["text"])
                elif msg["type"] == "progress":
                    pct = int(msg["pct"])
                    self.progress["value"] = pct
                    self.lbl_pct.config(text=f"{pct}%")
                elif msg["type"] == "error":
                    self._log(f"❌ ERRO: {msg['text']}")
                    messagebox.showerror("Erro", msg["text"])
                    self.lbl_output.config(text="Output: erro durante o processo")
                    self.btn_download.config(state="disabled")
                    self.btn_open_folder.config(state="disabled")
                    self.progress["value"] = 0
                    self.lbl_pct.config(text="0%")
                elif msg["type"] == "done":
                    path = msg["path"]
                    self._last_output_path = path
                    self.lbl_output.config(text=f"Output: {os.path.basename(path)} (gerado)")
                    self.btn_download.config(state="normal")
                    self.btn_open_folder.config(state="normal")
                    self._log("✅ Processo concluído.")
        except queue.Empty:
            pass
        self.after(100, self._process_ui_queue)

    def _log(self, text):
        self.txt_log.insert(tk.END, text + "\n")
        self.txt_log.see(tk.END)

if __name__ == "__main__":
    app = App()
    app.mainloop()
