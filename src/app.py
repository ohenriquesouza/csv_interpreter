import os
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox

from main import DEFAULT_THEME, process_csv


def resource_path(filename: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


def select_file(label: tk.Label, btn_process: tk.Button, state: dict) -> None:
    path = filedialog.askopenfilename(
        title="Selecionar arquivo CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    if path:
        state["csv_path"] = path
        label.config(text=os.path.basename(path), fg="#333333")
        btn_process.config(state=tk.NORMAL)


def pick_theme(state: dict, swatch: tk.Label) -> None:
    initial = f"#{state['theme_color']}"
    result = colorchooser.askcolor(color=initial, title="Selecionar cor do tema")
    if result and result[1]:
        hex_color = result[1].lstrip("#").upper()
        state["theme_color"] = hex_color
        swatch.config(bg=f"#{hex_color}")


def run_processing(
    csv_path: str, dest_dir: str, protect: bool, theme_color: str, null_treatment: str,
    btn_process: tk.Button, status: tk.Label,
) -> None:
    try:
        tmp_dir = tempfile.mkdtemp()
        tmp_csv = os.path.join(tmp_dir, os.path.basename(csv_path))
        shutil.copy2(csv_path, tmp_csv)

        out_xlsx = process_csv(
            tmp_csv,
            output_dir=dest_dir,
            protect=protect,
            theme_color=theme_color,
            null_treatment=null_treatment,
        )

        status.config(text="Arquivo gerado com sucesso!", fg="#2e7d32")
        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{out_xlsx}")
    except Exception as exc:
        status.config(text="Erro ao processar.", fg="#c62828")
        messagebox.showerror("Erro", str(exc))
    finally:
        btn_process.config(state=tk.NORMAL)
        shutil.rmtree(tmp_dir, ignore_errors=True)


def process(
    state: dict, protect_var: tk.BooleanVar, null_var: tk.StringVar,
    btn_process: tk.Button, status: tk.Label,
) -> None:
    csv_path = state.get("csv_path")
    if not csv_path:
        return

    dest_dir = filedialog.askdirectory(title="Escolha a pasta de destino do arquivo gerado")
    if not dest_dir:
        return

    btn_process.config(state=tk.DISABLED)
    status.config(text="Processando...", fg="#1565c0")

    thread = threading.Thread(
        target=run_processing,
        args=(
            csv_path, dest_dir, protect_var.get(),
            state["theme_color"], null_var.get(),
            btn_process, status,
        ),
        daemon=True,
    )
    thread.start()


def main() -> None:
    root = tk.Tk()
    root.title("CSV Interpreter")
    root.resizable(False, False)
    root.configure(bg="white")

    try:
        root.iconbitmap(resource_path("agrosynlogo.ico"))
    except Exception:
        pass

    window_w, window_h = 420, 295
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    root.geometry(f"{window_w}x{window_h}+{(screen_w - window_w) // 2}+{(screen_h - window_h) // 2}")

    state = {"csv_path": None, "theme_color": DEFAULT_THEME}

    tk.Label(root, text="CSV Interpreter", font=("Segoe UI", 14, "bold"),
             bg="white", fg="#1a1a1a").pack(pady=(20, 4))

    tk.Label(root, text="Selecione um arquivo CSV para gerar o relatório formatado.",
             font=("Segoe UI", 9), bg="white", fg="#555555").pack()

    file_label = tk.Label(root, text="Nenhum arquivo selecionado",
                          font=("Segoe UI", 9, "italic"), bg="white", fg="#aaaaaa")
    file_label.pack(pady=(12, 6))

    btn_select = tk.Button(
        root, text="Selecionar arquivo CSV",
        font=("Segoe UI", 10), bg="#2E75B6", fg="white",
        activebackground="#1f5490", activeforeground="white",
        relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
        command=lambda: select_file(file_label, btn_process, state),
    )
    btn_select.pack()

    # --- Theme selector ---
    theme_frame = tk.Frame(root, bg="white")
    theme_frame.pack(pady=(10, 0))

    swatch = tk.Label(
        theme_frame,
        bg=f"#{DEFAULT_THEME}",
        width=2, relief="groove", cursor="hand2",
    )
    swatch.pack(side=tk.LEFT)
    swatch.bind("<Button-1>", lambda e: pick_theme(state, swatch))

    btn_theme = tk.Button(
        theme_frame, text="Selecionar tema",
        font=("Segoe UI", 9), bg="white", fg="#2E75B6",
        activebackground="white", activeforeground="#1f5490",
        relief=tk.FLAT, cursor="hand2", borderwidth=0,
        command=lambda: pick_theme(state, swatch),
    )
    btn_theme.pack(side=tk.LEFT, padx=(6, 0))

    # --- Null treatment ---
    null_var = tk.StringVar(value="na")

    null_frame = tk.Frame(root, bg="white")
    null_frame.pack(pady=(10, 0))

    tk.Label(null_frame, text="Células vazias:",
             font=("Segoe UI", 9), bg="white", fg="#555555").pack(side=tk.LEFT)

    for label, val in [("N/A", "na"), ("Vazio", "vazio"), ("Zero", "zero")]:
        tk.Radiobutton(
            null_frame, text=label, variable=null_var, value=val,
            font=("Segoe UI", 9), bg="white", activebackground="white",
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(8, 0))

    # --- Process / protect ---
    protect_var = tk.BooleanVar(value=False)

    action_frame = tk.Frame(root, bg="white")
    action_frame.pack(pady=(10, 0))

    btn_process = tk.Button(
        action_frame, text="Processar e salvar",
        font=("Segoe UI", 10), bg="#e0e0e0", fg="#aaaaaa",
        relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
        state=tk.DISABLED,
        command=lambda: process(state, protect_var, null_var, btn_process, status_label),
    )
    btn_process.pack(side=tk.LEFT)

    chk_protect = tk.Checkbutton(
        action_frame, text="🔒", variable=protect_var,
        font=("Segoe UI", 12), bg="white", activebackground="white",
        cursor="hand2", borderwidth=0,
    )
    chk_protect.pack(side=tk.LEFT, padx=(8, 0))

    btn_process.bind("<Enter>", lambda e: btn_process.config(
        bg="#43a047", fg="white") if btn_process["state"] == tk.NORMAL else None)
    btn_process.bind("<Leave>", lambda e: btn_process.config(
        bg="#e0e0e0", fg="#aaaaaa") if btn_process["state"] == tk.DISABLED
        else btn_process.config(bg="#43a047"))

    status_label = tk.Label(root, text="", font=("Segoe UI", 9), bg="white")
    status_label.pack(pady=(12, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
