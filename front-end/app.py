"""Interface desktop do gerador de etiquetas ABB."""
from __future__ import annotations

import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from autoamte import (  # noqa: E402
    PADRAO,
    abrir,
    gerar_codigos_lote,
    imprimir,
    montar_documento,
    montar_lote,
    SAIDA,
)

ABB_RED = "#e30613"
ABB_DARK = "#20252b"
ABB_BLUE = "#0875c1"
ABB_LIGHT = "#f4f6f8"
ABB_LINE = "#d8dde3"
WHITE = "#ffffff"


class LabelApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ABB | Gerador de etiquetas")
        self.geometry("1020x900")
        self.minsize(900, 860)
        self.configure(bg=ABB_LIGHT)

        self.mode = tk.StringVar(value="lote")
        self.status = tk.StringVar(value="Pronto para gerar etiquetas")
        self.titulo = tk.StringVar(value=PADRAO["titulo"])
        self.rotulo = tk.StringVar(value=PADRAO["rotulo"])
        self.numero = tk.StringVar(value=PADRAO["numero"])
        self.codigo = tk.StringVar()
        self.prateleira = tk.StringVar(value="1")
        self.coluna_inicial = tk.StringVar(value="A")
        self.coluna_final = tk.StringVar(value="A")
        self.pasta_saida = tk.StringVar(value=str(SAIDA))
        self.ultimo_arquivo = None

        self._configure_styles()
        self._build_header()
        self._build_content()
        self._toggle_mode()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=ABB_LIGHT)
        style.configure("Card.TFrame", background=WHITE)
        style.configure("TLabel", background=WHITE, foreground=ABB_DARK, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=ABB_DARK, foreground=WHITE, font=("Segoe UI", 21, "bold"))
        style.configure("Subtitle.TLabel", background=ABB_DARK, foreground="#cbd2d9", font=("Segoe UI", 11))
        style.configure("Section.TLabel", background=WHITE, foreground=ABB_DARK, font=("Segoe UI", 12, "bold"))
        style.configure("Field.TLabel", background=WHITE, foreground="#59636e", font=("Segoe UI", 9, "bold"))
        style.configure("TEntry", fieldbackground=WHITE, padding=8)
        style.configure("TCombobox", fieldbackground=WHITE, padding=6)
        style.configure("Accent.TButton", background=ABB_RED, foreground=WHITE, borderwidth=0, padding=(16, 10), font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#b8040f")])
        style.configure("Secondary.TButton", background=ABB_DARK, foreground=WHITE, borderwidth=0, padding=(14, 10), font=("Segoe UI", 10, "bold"))
        style.map("Secondary.TButton", background=[("active", "#3a434c")])
        style.configure("Mode.TRadiobutton", background=WHITE, foreground=ABB_DARK, padding=(8, 6), font=("Segoe UI", 10, "bold"))
        style.map("Mode.TRadiobutton", foreground=[("selected", ABB_RED)])
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10), background=WHITE, fieldbackground=WHITE)
        style.configure("Treeview.Heading", background=ABB_DARK, foreground=WHITE, font=("Segoe UI", 9, "bold"))

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=ABB_DARK, height=116)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Frame(header, bg=ABB_RED, width=10).pack(side="left", fill="y")
        text = tk.Frame(header, bg=ABB_DARK)
        text.pack(side="left", padx=28, pady=20)
        ttk.Label(text, text="GERADOR DE ETIQUETAS", style="Title.TLabel").pack(anchor="w")
        ttk.Label(text, text="Almoxarifado  /  Controle de localização", style="Subtitle.TLabel").pack(anchor="w", pady=(5, 0))
        tk.Label(header, text="ABB", bg=ABB_DARK, fg=WHITE, font=("Segoe UI", 19, "bold")).pack(side="right", padx=30)

    def _build_content(self) -> None:
        body = ttk.Frame(self, padding=(28, 24))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(1, weight=1)

        form = ttk.Frame(body, style="Card.TFrame", padding=22)
        form.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 18))
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Nova geração", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(form, text="Escolha o tipo de impressão", style="Field.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(22, 4))
        modes = ttk.Frame(form, style="Card.TFrame")
        modes.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(modes, text="Lote", variable=self.mode, value="lote", command=self._toggle_mode, style="Mode.TRadiobutton").pack(side="left")
        ttk.Radiobutton(modes, text="Manual", variable=self.mode, value="manual", command=self._toggle_mode, style="Mode.TRadiobutton").pack(side="left", padx=(14, 0))

        self.batch_frame = ttk.Frame(form, style="Card.TFrame")
        self.batch_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        self.batch_frame.columnconfigure(1, weight=1)
        self.batch_frame.columnconfigure(3, weight=1)
        self._field(self.batch_frame, "Prateleira", self.prateleira, 0, 0, ["1", "2", "3"])
        self._field(self.batch_frame, "Coluna inicial", self.coluna_inicial, 0, 2, list("ABCDEFGHIJKLMNOPQR"))
        self._field(self.batch_frame, "Coluna final", self.coluna_final, 1, 0, list("ABCDEFGHIJKLMNOPQR"))

        self.manual_frame = ttk.Frame(form, style="Card.TFrame")
        self.manual_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        self.manual_frame.columnconfigure(1, weight=1)
        ttk.Label(self.manual_frame, text="Código da etiqueta", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(self.manual_frame, textvariable=self.codigo).grid(row=0, column=1, sticky="ew")

        base_row = 4
        ttk.Label(form, text="Dados da etiqueta", style="Section.TLabel").grid(row=base_row, column=0, columnspan=2, sticky="w", pady=(28, 14))
        self._entry(form, "Título", self.titulo, base_row + 1, 0)
        self._entry(form, "Rótulo do rodapé", self.rotulo, base_row + 1, 1)
        self._entry(form, "Número do rodapé", self.numero, base_row + 2, 0)

        output = ttk.Frame(form, style="Card.TFrame")
        output.grid(row=base_row + 11, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        output.columnconfigure(0, weight=1)
        ttk.Label(output, text="Pasta de saída", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(output, textvariable=self.pasta_saida).grid(row=1, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(output, text="Escolher pasta...", style="Secondary.TButton", command=self.choose_output_folder).grid(row=1, column=1, sticky="e", padx=(10, 0))

        actions = ttk.Frame(form, style="Card.TFrame")
        actions.grid(row=base_row + 13, column=0, columnspan=2, sticky="ew", pady=(22, 0))
        ttk.Button(actions, text="Gerar documento", style="Accent.TButton", command=self.generate).pack(side="left")
        ttk.Button(actions, text="Gerar e imprimir", style="Accent.TButton", command=self.generate_and_print).pack(side="left", padx=10)
        ttk.Button(actions, text="Limpar", style="Secondary.TButton", command=self.clear).pack(side="left", padx=10)

        preview = ttk.Frame(body, style="Card.TFrame", padding=22)
        preview.grid(row=0, column=1, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        ttk.Label(preview, text="Prévia do lote", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(preview, text="A sequência respeita prateleira, andar e coluna.", wraplength=270).grid(row=1, column=0, sticky="w", pady=(6, 16))
        self.tree = ttk.Treeview(preview, columns=("codigo",), show="headings", height=9)
        self.tree.heading("codigo", text="Códigos selecionados")
        self.tree.column("codigo", anchor="center", width=230)
        self.tree.grid(row=2, column=0, sticky="nsew")
        preview.rowconfigure(2, weight=1)
        self._refresh_preview()
        for variable in (self.prateleira, self.coluna_inicial, self.coluna_final):
            variable.trace_add("write", lambda *_: self._refresh_preview())

        status = tk.Frame(body, bg=ABB_LIGHT)
        status.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        tk.Label(status, textvariable=self.status, bg=ABB_LIGHT, fg="#59636e", font=("Segoe UI", 9)).pack(side="left")
        self.open_button = ttk.Button(status, text="Abrir último", style="Secondary.TButton", command=self.open_last, state="disabled")
        self.open_button.pack(side="right")
        self.print_button = ttk.Button(status, text="Imprimir último", style="Accent.TButton", command=self.print_last, state="disabled")
        self.print_button.pack(side="right", padx=(0, 8))

    def _field(self, parent, label, variable, row, column, values) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row * 2, column=column, sticky="w", padx=(0, 10), pady=(0, 4))
        ttk.Combobox(parent, textvariable=variable, values=values, state="readonly", width=8).grid(row=row * 2 + 1, column=column, sticky="ew", padx=(0, 16), pady=(0, 12))

    def _entry(self, parent, label, variable, row, column) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row * 2, column=column, sticky="w", padx=(0, 12), pady=(0, 4))
        ttk.Entry(parent, textvariable=variable).grid(row=row * 2 + 1, column=column, sticky="ew", padx=(0, 12), pady=(0, 10))

    def _toggle_mode(self) -> None:
        if self.mode.get() == "lote":
            self.batch_frame.grid()
            self.manual_frame.grid_remove()
        else:
            self.batch_frame.grid_remove()
            self.manual_frame.grid()
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            if self.mode.get() == "lote":
                codes = gerar_codigos_lote(self.prateleira.get(), self.coluna_inicial.get(), self.coluna_final.get())
            else:
                codes = [self.codigo.get()] if self.codigo.get() else []
        except ValueError:
            codes = []
        for code in codes[:30]:
            self.tree.insert("", "end", values=(code,))
        if len(codes) > 30:
            self.tree.insert("", "end", values=(f"... e mais {len(codes) - 30} etiquetas",))

    def generate(self, print_after: bool = False) -> None:
        title = self.titulo.get()
        footer_label = self.rotulo.get()
        footer_number = self.numero.get()
        mode = self.mode.get()
        try:
            if mode == "lote":
                codes = gerar_codigos_lote(self.prateleira.get(), self.coluna_inicial.get(), self.coluna_final.get())
            else:
                if not self.codigo.get().strip():
                    raise ValueError("informe o código da etiqueta")
                codes = [self.codigo.get().strip()]
        except ValueError as error:
            messagebox.showerror("Dados inválidos", str(error))
            return

        self.status.set("Gerando documento... aguarde")
        self._set_buttons(False)
        threading.Thread(
            target=self._generate_worker,
            args=(codes, title, footer_label, footer_number, mode, self.pasta_saida.get(), print_after),
            daemon=True,
        ).start()

    def generate_and_print(self) -> None:
        self.generate(print_after=True)

    def _generate_worker(self, codes, title, footer_label, footer_number, mode, output_folder, print_after) -> None:
        try:
            PADRAO.update(titulo=title, rotulo=footer_label, numero=footer_number)
            if len(codes) == 1 and mode == "manual":
                path = montar_documento(title, codes[0], footer_label, footer_number, Path(output_folder))
            else:
                path = montar_lote(title, codes, footer_label, footer_number, Path(output_folder))
            if print_after:
                imprimir(path)
            self.after(0, self._generation_finished, path, len(codes), print_after)
        except Exception as error:
            self.after(0, self._generation_failed, error)

    def _generation_finished(self, path: Path, count: int, printed: bool = False) -> None:
        self.ultimo_arquivo = path
        acao = " salva(s) e enviada(s) para impressão" if printed else " salva(s)"
        self.status.set(f"{count} etiqueta(s){acao} em {path.name}")
        self._set_buttons(True)
        mensagem = f"Arquivo criado em:\n{path}"
        if printed:
            mensagem += "\n\nO documento foi enviado para a impressora padrão."
        messagebox.showinfo("Documento gerado", mensagem)

    def _generation_failed(self, error: Exception) -> None:
        self.status.set("Não foi possível gerar o documento")
        self._set_buttons(True)
        messagebox.showerror("Erro ao gerar", str(error))

    def _set_buttons(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.open_button.configure(state=state if self.ultimo_arquivo else "disabled")
        self.print_button.configure(state=state if self.ultimo_arquivo else "disabled")

    def open_last(self) -> None:
        if self.ultimo_arquivo:
            abrir(self.ultimo_arquivo)

    def print_last(self) -> None:
        if self.ultimo_arquivo:
            try:
                imprimir(self.ultimo_arquivo)
                self.status.set("Documento enviado para a impressora padrão")
            except Exception as error:
                messagebox.showerror("Erro ao imprimir", str(error))

    def clear(self) -> None:
        self.mode.set("lote")
        self.titulo.set("")
        self.rotulo.set("")
        self.numero.set("")
        self.codigo.set("")
        self.prateleira.set("1")
        self.coluna_inicial.set("A")
        self.coluna_final.set("A")
        self.pasta_saida.set(str(SAIDA))
        self.ultimo_arquivo = None
        self.open_button.configure(state="disabled")
        self.print_button.configure(state="disabled")
        self._toggle_mode()
        self._refresh_preview()
        self.status.set("Campos limpos")

    def choose_output_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="Escolha onde salvar os documentos",
            initialdir=self.pasta_saida.get(),
        )
        if folder:
            self.pasta_saida.set(folder)
            self.status.set(f"Pasta de saída: {folder}")


if __name__ == "__main__":
    LabelApp().mainloop()
