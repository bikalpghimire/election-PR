import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import matplotlib
import requests
from bs4 import BeautifulSoup

# --- CRITICAL: mplcairo Initialization for Devanagari Font Shaping ---
try:
    import mplcairo
    matplotlib.use("module://mplcairo.tk")
except ImportError:
    matplotlib.use("TkAgg") 
    
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import matplotlib.colors as mcolors
import winsound
import sys
import os
import ctypes
import re
import subprocess
from io import BytesIO
from scraper import fetch_pr_votes

# --- Global Configuration ---
FONT_NAME = "Kokila"
MAROON_COLOR = "#800000" 
matplotlib.rcParams["axes.unicode_minus"] = False

def get_best_font_color(hex_color):
    """Refined contrast: Forces dark text on light blue and pastels [cite: 2026-03-07]."""
    rgb = mcolors.to_rgb(hex_color)
    # W3C Relative Luminance formula
    luminance = (0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2])
    # Threshold adjusted to 0.55 to catch 'light blue' and switch to maroon
    return 'white' if luminance < 0.55 else MAROON_COLOR

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def calculate_seats(df, total_seats, threshold_pct, manual_total_votes):
    threshold_limit = manual_total_votes * (threshold_pct / 100)
    eligible_df = df[df["Votes"] >= threshold_limit].copy()
    total_input_parties = len(df)
    
    if eligible_df.empty:
        return None, manual_total_votes, threshold_limit, 0, total_input_parties

    eligible_df["Seats"] = 0
    eligible_df["Votes"] = eligible_df["Votes"].astype(float)
    for _ in range(total_seats):
        eligible_df["Quotient"] = eligible_df["Votes"] / (2 * eligible_df["Seats"] + 1)
        winner_idx = eligible_df["Quotient"].idxmax()
        eligible_df.at[winner_idx, "Seats"] += 1
        
    return eligible_df[["Party", "Votes", "Seats"]].sort_values(by="Seats", ascending=False), \
           manual_total_votes, threshold_limit, len(eligible_df), total_input_parties

class ElectionApp:
    def __init__(self, root):

        self.root = root
        self.root.title("Nepal PR Seat Calculator Pro")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        w, h = 1380, 880 
        ws, hs = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{int(ws/2 - w/2)}+{int(hs/2 - h/2)}")
        self.root.configure(bg="#f8f9fa")

        self.result_df = None
        self.filepath = None
        self.current_figure = None
        self.total_input_votes = 0
        self.pan_start = None

        logo_path = resource_path("pumori_logo.png")
        icon_path = resource_path("pumori_icon.ico")
        try:
            icon_img = ImageTk.PhotoImage(Image.open(icon_path))
            self.root.iconphoto(True, icon_img)
            self._icon_ref = icon_img
            load = Image.open(logo_path)
            self.logo_img = ImageTk.PhotoImage(load.resize((180, 50), Image.Resampling.LANCZOS))
        except Exception:
            self.logo_img = None

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", font=(FONT_NAME, 14), rowheight=35)
        style.configure("Treeview.Heading", font=(FONT_NAME, 14, "bold"), background="#d1d8e0")

        header = tk.Frame(root, bg="#2c3e50", height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        if self.logo_img:
            logo_lbl = tk.Label(header, image=self.logo_img, bg="#2c3e50")
            logo_lbl.place(relx=0.0, rely=0.5, anchor="w", x=20)

        tk.Label(header, text="Nepal Election PR Seat Calculator", font=(FONT_NAME, 26, "bold"), 
                 fg="white", bg="#2c3e50").place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(header, text="Developed By: Pumori Engineering Services", font=(FONT_NAME, 12), 
                 fg="#bdc3c7", bg="#2c3e50").place(relx=1.0, rely=0.5, anchor="e", x=-30)

        ctrl_bar = tk.Frame(root, bg="white", pady=15, highlightbackground="#e1e4e8", highlightthickness=1)
        ctrl_bar.pack(fill="x", padx=25, pady=10)
        btn_config = {"font": (FONT_NAME, 12, "bold"), "fg": "white", "padx": 15, "relief": "flat"}

        tk.Button(ctrl_bar, text="📁 Load Excel", command=self.load_file, bg="#34495e", **btn_config).grid(row=0, column=0, padx=15)
        tk.Button(ctrl_bar, text="🌐 Load From Web", command=self.fetch_from_web, bg="#16a085", **btn_config).grid(row=0, column=1, padx=10)
        self.seats_entry = tk.Entry(ctrl_bar, width=8, font=(FONT_NAME, 12), justify="center", bd=2, relief="groove")
        self.seats_entry.insert(0, "110")
        self.seats_entry.grid(row=0, column=2, padx=10)
        tk.Button(ctrl_bar, text="⚡ Calculate", command=self.process, bg="#006994", **btn_config).grid(row=0, column=3, padx=10)
        tk.Button(ctrl_bar, text="💾 Export", command=self.export_file, bg="#2c3e50", **btn_config).grid(row=0, column=4, padx=10)
        tk.Button(ctrl_bar, text="🔄 Reset", command=self.reset_app, bg="#95a5a6", **btn_config).grid(row=0, column=5, padx=15)

        self.main_container = tk.PanedWindow(root, orient="horizontal", bg="#f8f9fa", sashwidth=6)
        self.main_container.pack(fill="both", expand=True, padx=25)

        self.table_frame = tk.Frame(self.main_container, bg="white")
        self.main_container.add(self.table_frame, width=500)

        self.tree = ttk.Treeview(self.table_frame, columns=("Party", "Votes", "Seats"), show="headings")
        self.tree.heading("Party", text="Political Party Name"); self.tree.column("Party", width=280, anchor="w")
        self.tree.heading("Votes", text="Votes"); self.tree.column("Votes", width=120, anchor="center")
        self.tree.heading("Seats", text="Seats"); self.tree.column("Seats", width=100, anchor="center") 
        self.tree.pack(fill="both", expand=True)

        self.right_pane = tk.PanedWindow(self.main_container, orient="vertical", bg="#f8f9fa", sashwidth=6)
        self.main_container.add(self.right_pane)

        self.chart_frame = tk.Frame(self.right_pane, bg="white", bd=1, relief="solid")
        self.right_pane.add(self.chart_frame, height=516) # 63% Chart Area

        self.legend_frame = tk.Frame(self.right_pane, bg="white", bd=1, relief="solid")
        self.right_pane.add(self.legend_frame, height=304) # 37% Legend Area

        tk.Label(self.legend_frame, text="   Legend", font=(FONT_NAME, 14, "bold"), bg="white", 
                 fg="#2c3e50").pack(anchor="w", padx=10, pady=(8, 2))

        self.legend_canvas = tk.Canvas(self.legend_frame, bg="white", highlightthickness=0)
        self.legend_scrollbar = tk.Scrollbar(self.legend_frame, orient="vertical", command=self.legend_canvas.yview)
        self.legend_canvas.configure(yscrollcommand=self.legend_scrollbar.set)
        self.legend_scrollbar.pack(side="right", fill="y")
        self.legend_canvas.pack(side="left", fill="both", expand=True)

        self.legend_container = tk.Frame(self.legend_canvas, bg="white")
        self.legend_window = self.legend_canvas.create_window((0, 0), window=self.legend_container, anchor="nw")
        self.legend_container.bind("<Configure>", lambda e: self.legend_canvas.configure(scrollregion=self.legend_canvas.bbox("all")))

        self.summary_frame = tk.Frame(root, bg="#ecf0f1", height=80)
        self.summary_frame.pack(fill="x", side="bottom", padx=25, pady=10)
        lbl_style = {"bg": "#ecf0f1", "font": (FONT_NAME, 14, "bold"), "fg": "#2c3e50"}
        self.lbl_total_votes = tk.Label(self.summary_frame, text="Total Valid Votes: -", **lbl_style); self.lbl_total_votes.pack(side="left", padx=40)
        self.lbl_threshold = tk.Label(self.summary_frame, text="3% Threshold: -", **lbl_style); self.lbl_threshold.pack(side="left", padx=40)
        self.lbl_qualified = tk.Label(self.summary_frame, text="Qualified Parties: -", **lbl_style); self.lbl_qualified.pack(side="left", padx=40)
        self.right_pane.bind("<Configure>", self.maintain_pane_ratio)

    def maintain_pane_ratio(self, event=None):
        try:
            total_height = self.right_pane.winfo_height()
            chart_height = int(total_height * 0.63)
            self.right_pane.sash_place(0, 0, chart_height)
        except:
            pass

    def on_closing(self):
        try:
            if self.current_figure: plt.close(self.current_figure)
            plt.close("all")
        except: pass
        self.root.destroy()
        os._exit(0)

    def clean_number(self, text):
        if pd.isna(text): return 0.0
        text = str(text).strip().replace(",", "").replace("\xa0", "")
        text = re.sub(r"[^\d०१२३४५६७८९.]", "", text)
        translation_table = str.maketrans("०१२३४५६७८९", "0123456789")
        try: return float(text.translate(translation_table))
        except: return 0.0

    def load_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if f: self.filepath = f; winsound.MessageBeep(winsound.MB_ICONASTERISK)
        
    def fetch_from_web(self):

        try:

            df, total_votes = fetch_pr_votes()

            self.total_input_votes = total_votes

            res, total, thresh, count, total_p = calculate_seats(
                df[["Party", "Votes"]],
                int(self.seats_entry.get()),
                3.0,
                total_votes
            )

            self.result_df = res

            self.lbl_total_votes.config(text=f"Total Valid Votes: {int(total):,}")
            self.lbl_threshold.config(text=f"3% Threshold: {int(thresh):,}")
            self.lbl_qualified.config(text=f"Qualified Parties: {count} out of {total_p}")

            for i in self.tree.get_children():
                self.tree.delete(i)

            if res is not None:
                for _, r in res.iterrows():
                    self.tree.insert(
                        "",
                        tk.END,
                        values=(r["Party"], f"{int(r['Votes']):,}", f"  {int(r['Seats'])}  ")
                    )

            self.update_chart()

            winsound.MessageBeep(winsound.MB_ICONASTERISK)

        except Exception as e:
            messagebox.showerror("Web Fetch Error", str(e))


    def clear_legend(self):
        for widget in self.legend_container.winfo_children(): widget.destroy()
        self.legend_canvas.yview_moveto(0)

    def update_legend(self, chart_data, colors):
        self.clear_legend()
        total_seats = int(chart_data["Seats"].sum())
        max_cols = 2
        for idx, (_, row) in enumerate(chart_data.iterrows()):
            box_row, box_col = idx // max_cols, idx % max_cols
            item = tk.Frame(self.legend_container, bg="white")
            item.grid(row=box_row, column=box_col, sticky="nw", padx=(24, 70), pady=3)
            color_hex = matplotlib.colors.to_hex(colors[idx % len(colors)])
            tk.Label(item, bg=color_hex, width=2, height=1, relief="solid", bd=1).pack(side="left", padx=(0, 12), anchor="n")
            pct = (row["Seats"] / total_seats * 100) if total_seats else 0
            text_block = f"{idx + 1}. {row['Party']}\n    Seats: {int(row['Seats'])}  |  Vote Share: {pct:.1f}%"
            tk.Label(item, text=text_block, font=(FONT_NAME, 11), bg="white", fg="#2c3e50", anchor="w", justify="left", wraplength=340).pack(side="left", anchor="w")
        for c in range(max_cols): self.legend_container.grid_columnconfigure(c, minsize=460)
        self.legend_canvas.configure(scrollregion=self.legend_canvas.bbox("all"))

    def update_chart(self):
        """Refined adaptive font colors for better visibility on light blue [cite: 2026-03-07]."""
        for w in self.chart_frame.winfo_children(): w.destroy()
        if self.current_figure:
            try: plt.close(self.current_figure)
            except: pass
            self.current_figure = None
        self.clear_legend()
        if self.result_df is None: return
        chart_data = self.result_df[self.result_df["Seats"] > 0].copy()
        if chart_data.empty: return

        fig, ax = plt.subplots(figsize=(10, 8), dpi=110, facecolor="white")
        self.current_figure = fig
        total_seats = int(chart_data["Seats"].sum())
        colors = list(matplotlib.colormaps.get_cmap("tab20").colors)
        slice_colors = colors[:len(chart_data)]
        numeric_labels = [str(i + 1) for i in range(len(chart_data))]

        def autopct_func(pct):
            seats = int(round(pct * total_seats / 100.0))
            return f"{pct:.1f}%\n{seats}"

        prop = fm.FontProperties(family=FONT_NAME)
        
        wedges, texts, autotexts = ax.pie(
            chart_data["Seats"], labels=numeric_labels, autopct=autopct_func, startangle=90, 
            counterclock=False, colors=slice_colors, labeldistance=1.06, pctdistance=0.76,
            wedgeprops={"width": 0.40, "edgecolor": "white", "linewidth": 2} 
        )
        
        for t in texts: t.set_fontproperties(prop); t.set_fontsize(13); t.set_fontweight("bold")
        
        # Apply logic to each slice to prevent 'White on Blue' blurriness [cite: 2026-03-07]
        for i, at in enumerate(autotexts): 
            slice_color = slice_colors[i % len(slice_colors)]
            at.set_fontproperties(prop)
            at.set_fontsize(9.5)
            at.set_color(get_best_font_color(slice_color)) 
            at.set_fontweight("bold")

        ax.text(0, 0.15, f"{total_seats}", ha="center", va="center", fontsize=30, fontweight="bold", color="#2c3e50")
        ax.text(0, 0.0, "Total Seats", ha="center", va="center", fontsize=12, fontweight="bold", color="#5d6d7e")
        ax.text(0, -0.15, f"{int(self.total_input_votes):,}", ha="center", va="center", fontsize=16, fontweight="bold", color="#2980b9")
        ax.text(0, -0.25, "Total Votes", ha="center", va="center", fontsize=10, fontweight="bold", color="#5d6d7e")

        fig.suptitle(
            "PR Seat Allocation Result",
            fontsize=12,
            y=0.99
        )
        ax.axis("equal")
        fig.subplots_adjust(top=0.90, bottom=0.04, left=0.04, right=0.96)

        # Pan/Zoom
        def on_scroll(event):
            if event.inaxes != ax: return
            base_scale = 1.1
            scale_factor = 1 / base_scale if event.button == 'up' else base_scale
            cur_xlim, cur_ylim = ax.get_xlim(), ax.get_ylim()
            new_width, new_height = (cur_xlim[1] - cur_xlim[0]) * scale_factor, (cur_ylim[1] - cur_ylim[0]) * scale_factor
            rel_x, rel_y = (cur_xlim[1] - event.xdata) / (cur_xlim[1] - cur_xlim[0]), (cur_ylim[1] - event.ydata) / (cur_ylim[1] - cur_ylim[0])
            ax.set_xlim([event.xdata - new_width * (1 - rel_x), event.xdata + new_width * rel_x])
            ax.set_ylim([event.ydata - new_height * (1 - rel_y), event.ydata + new_height * rel_y])
            fig.canvas.draw_idle()

        def on_press(event):
            if event.inaxes != ax: return
            self.pan_start = (event.xdata, event.ydata)

        def on_motion(event):
            if self.pan_start is None or event.inaxes != ax: return
            dx, dy = event.xdata - self.pan_start[0], event.ydata - self.pan_start[1]
            ax.set_xlim(ax.get_xlim() - dx); ax.set_ylim(ax.get_ylim() - dy); fig.canvas.draw_idle()

        fig.canvas.mpl_connect('scroll_event', on_scroll)
        fig.canvas.mpl_connect('button_press_event', on_press)
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('button_release_event', lambda e: setattr(self, 'pan_start', None))

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame); canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.update_legend(chart_data, colors)

    def process(self):
        if not self.filepath: return
        try:
            raw_df = pd.read_excel(self.filepath, header=None)
            df_clean = raw_df.dropna(how="all").reset_index(drop=True)
            manual_total = self.clean_number(df_clean.iloc[-1, 0])
            self.total_input_votes = manual_total
            party_data = []
            for i in range(1, len(df_clean) - 1, 2):
                p_name, p_votes = str(df_clean.iloc[i, 0]).strip(), self.clean_number(df_clean.iloc[i + 1, 0])
                if p_name and p_votes > 0: party_data.append({"Party": p_name, "Votes": p_votes})
            
            res, total, thresh, count, total_p = calculate_seats(pd.DataFrame(party_data), 
                                                                int(self.seats_entry.get()), 3.0, manual_total)
            self.result_df = res
            self.lbl_total_votes.config(text=f"Total Valid Votes: {int(total):,}")
            self.lbl_threshold.config(text=f"3% Threshold: {int(thresh):,}")
            self.lbl_qualified.config(text=f"Qualified Parties: {count} out of {total_p}")
            
            for i in self.tree.get_children(): self.tree.delete(i)
            if res is not None:
                for _, r in res.iterrows(): 
                    self.tree.insert("", tk.END, values=(r["Party"], f"{int(r['Votes']):,}", f"  {int(r['Seats'])}  "))
                self.update_chart(); winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception as e: messagebox.showerror("Error", str(e))

    def reset_app(self):
        self.filepath = self.result_df = None; self.total_input_votes = 0
        for i in self.tree.get_children(): self.tree.delete(i)
        for w in self.chart_frame.winfo_children(): w.destroy()
        self.clear_legend()
        if self.current_figure: plt.close(self.current_figure); self.current_figure = None
        self.lbl_total_votes.config(text="Total Valid Votes: -"); self.lbl_threshold.config(text="3% Threshold: -"); self.lbl_qualified.config(text="Qualified Parties: -")
        winsound.MessageBeep(winsound.MB_ICONASTERISK)

    def export_file(self):
        if self.result_df is not None:
            path = filedialog.asksaveasfilename(initialfile="Nepal_PR_result.xlsx", defaultextension=".xlsx")
            if path:
                writer = pd.ExcelWriter(path, engine="xlsxwriter")
                self.result_df.to_excel(writer, index=False, sheet_name="PR_Results")
                worksheet = writer.sheets["PR_Results"]
                max_row, max_col = self.result_df.shape
                columns = [{"header": col} for col in self.result_df.columns]
                worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": columns, "style": "Table Style Medium 9"})
                for i, col in enumerate(self.result_df.columns):
                    width = max(self.result_df[col].astype(str).str.len().max(), len(col)) + 2
                    worksheet.set_column(i, i, width)
                writer.close()
                messagebox.showinfo("Success", "Export Complete.")
                if sys.platform == "win32": os.startfile(os.path.dirname(path))

if __name__ == "__main__":
    try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"pumori.engineering.seatcalc.pro")
    except Exception: pass
    root = tk.Tk(); app = ElectionApp(root); root.mainloop()