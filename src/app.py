import os
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from main import process_csv


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


def run_processing(csv_path: str, dest_dir: str, btn_process: tk.Button, status: tk.Label) -> None:
    try:
        tmp_dir = tempfile.mkdtemp()
        tmp_csv = os.path.join(tmp_dir, os.path.basename(csv_path))
        shutil.copy2(csv_path, tmp_csv)

        out_xlsx = process_csv(tmp_csv, output_dir=dest_dir)

        status.config(text=f"Arquivo gerado com sucesso!", fg="#2e7d32")
        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{out_xlsx}")
    except Exception as exc:
        status.config(text="Erro ao processar.", fg="#c62828")
        messagebox.showerror("Erro", str(exc))
    finally:
        btn_process.config(state=tk.NORMAL)
        shutil.rmtree(tmp_dir, ignore_errors=True)


def process(state: dict, btn_process: tk.Button, status: tk.Label) -> None:
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
        args=(csv_path, dest_dir, btn_process, status),
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

    window_w, window_h = 420, 220
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    root.geometry(f"{window_w}x{window_h}+{(screen_w - window_w) // 2}+{(screen_h - window_h) // 2}")

    state = {"csv_path": None}

    tk.Label(root, text="CSV Interpreter", font=("Segoe UI", 14, "bold"),
             bg="white", fg="#1a1a1a").pack(pady=(28, 4))

    tk.Label(root, text="Selecione um arquivo CSV para gerar o relatório formatado.",
             font=("Segoe UI", 9), bg="white", fg="#555555").pack()

    file_label = tk.Label(root, text="Nenhum arquivo selecionado",
                          font=("Segoe UI", 9, "italic"), bg="white", fg="#aaaaaa")
    file_label.pack(pady=(18, 6))

    btn_select = tk.Button(
        root, text="Selecionar arquivo CSV",
        font=("Segoe UI", 10), bg="#2E75B6", fg="white",
        activebackground="#1f5490", activeforeground="white",
        relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
        command=lambda: select_file(file_label, btn_process, state),
    )
    btn_select.pack()

    btn_process = tk.Button(
        root, text="Processar e salvar",
        font=("Segoe UI", 10), bg="#e0e0e0", fg="#aaaaaa",
        relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
        state=tk.DISABLED,
        command=lambda: process(state, btn_process, status_label),
    )
    btn_process.pack(pady=(10, 0))

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
