import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime


class ChequeoArchivosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Chequeo de archivos vs. Excel")
        # Ícono de la ventana (funciona tanto en Python como dentro del .exe de PyInstaller)
        try:
            base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
            self.root.iconbitmap(os.path.join(base, "icono_chequeo.ico"))
        except Exception:
            pass  # si no encuentra el ícono, sigue con el de por defecto
        self.root.geometry("680x620")

        self.excel_path = tk.StringVar()
        self.folder_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.column_name = tk.StringVar()
        self.df = None

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # Excel
        frame_excel = ttk.LabelFrame(self.root, text="1) Archivo Excel")
        frame_excel.pack(fill="x", **pad)
        ttk.Entry(frame_excel, textvariable=self.excel_path, width=50).pack(
            side="left", padx=5, pady=8, fill="x", expand=True
        )
        ttk.Button(frame_excel, text="Buscar...", command=self.elegir_excel).pack(
            side="left", padx=5
        )
        ttk.Label(frame_excel, text="Fila de títulos:").pack(side="left", padx=(10, 2))
        self.fila_titulos = tk.StringVar(value="1")
        spin_fila = ttk.Spinbox(
            frame_excel, from_=1, to=100, textvariable=self.fila_titulos, width=4,
            command=self.cargar_columnas,
        )
        spin_fila.pack(side="left", padx=5)
        spin_fila.bind("<Return>", self.cargar_columnas)
        spin_fila.bind("<FocusOut>", self.cargar_columnas)

        # Hoja y columna
        frame_cols = ttk.LabelFrame(self.root, text="2) Hoja y columna con el número interno")
        frame_cols.pack(fill="x", **pad)

        ttk.Label(frame_cols, text="Hoja:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_sheet = ttk.Combobox(frame_cols, textvariable=self.sheet_name, state="readonly")
        self.combo_sheet.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.combo_sheet.bind("<<ComboboxSelected>>", self.cargar_columnas)

        ttk.Label(frame_cols, text="Columna:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_col = ttk.Combobox(frame_cols, textvariable=self.column_name, state="readonly")
        self.combo_col.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        frame_cols.columnconfigure(1, weight=1)

        ttk.Label(
            frame_cols, text="Columnas a incluir en el detalle\nde faltantes (Ctrl+clic para varias):"
        ).grid(row=2, column=0, padx=5, pady=5, sticky="nw")
        self.listbox_extra = tk.Listbox(frame_cols, selectmode="multiple", height=5, exportselection=False)
        self.listbox_extra.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Carpeta
        frame_folder = ttk.LabelFrame(self.root, text="3) Carpeta con los archivos")
        frame_folder.pack(fill="x", **pad)
        ttk.Entry(frame_folder, textvariable=self.folder_path, width=60).pack(
            side="left", padx=5, pady=8, fill="x", expand=True
        )
        ttk.Button(frame_folder, text="Buscar...", command=self.elegir_carpeta).pack(
            side="left", padx=5
        )

        # Opciones
        frame_opts = ttk.LabelFrame(self.root, text="Opciones")
        frame_opts.pack(fill="x", **pad)
        self.var_incluir_subcarpetas = tk.BooleanVar(value=False)
        self.var_mostrar_extra = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_opts, text="Incluir subcarpetas", variable=self.var_incluir_subcarpetas
        ).pack(anchor="w", padx=5)
        ttk.Checkbutton(
            frame_opts,
            text="Mostrar también archivos que no corresponden a ningún número del Excel",
            variable=self.var_mostrar_extra,
        ).pack(anchor="w", padx=5)

        frame_prefijo = ttk.Frame(frame_opts)
        frame_prefijo.pack(anchor="w", padx=5, pady=(4, 0), fill="x")
        ttk.Label(
            frame_prefijo, text="Prefijo a agregar al número del Excel antes de comparar:"
        ).pack(side="left")
        self.prefijo = tk.StringVar(value="I")
        ttk.Entry(frame_prefijo, textvariable=self.prefijo, width=6).pack(side="left", padx=5)
        ttk.Label(
            frame_prefijo, text="(dejar vacío si el Excel ya incluye el prefijo)"
        ).pack(side="left")

        # Botón ejecutar
        ttk.Button(self.root, text="Chequear", command=self.chequear).pack(pady=10)

        # Resultado
        frame_res = ttk.LabelFrame(self.root, text="Resultado")
        frame_res.pack(fill="both", expand=True, **pad)

        self.text_result = tk.Text(frame_res, wrap="word")
        scrollbar = ttk.Scrollbar(frame_res, command=self.text_result.yview)
        self.text_result.configure(yscrollcommand=scrollbar.set)
        self.text_result.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.btn_exportar = ttk.Button(
            self.root, text="Exportar reporte a Excel", command=self.exportar_reporte, state="disabled"
        )
        self.btn_exportar.pack(pady=6)

        self._faltantes = []
        self._extra = []

    # ---------- Acciones ----------
    def elegir_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm")],
        )
        if not path:
            return
        self.excel_path.set(path)
        try:
            xls = pd.ExcelFile(path)
            self.combo_sheet["values"] = xls.sheet_names
            self.combo_sheet.current(0)
            self.cargar_columnas()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")

    def cargar_columnas(self, event=None):
        if not self.excel_path.get() or not self.sheet_name.get():
            return
        try:
            fila = int(self.fila_titulos.get())
        except ValueError:
            fila = 1
        try:
            self.df = pd.read_excel(
                self.excel_path.get(),
                sheet_name=self.sheet_name.get(),
                header=fila - 1,  # pandas usa índice 0, la fila 1 del Excel = header 0
            )
            columnas = list(self.df.columns)
            self.combo_col["values"] = columnas
            # Solo elije una columna por defecto si el usuario todavía no había
            # seleccionado ninguna (para no pisar su elección al recargar)
            if columnas and self.column_name.get() not in columnas:
                self.combo_col.current(0)

            # Poblar el listado de columnas adicionales
            seleccion_previa = [self.listbox_extra.get(i) for i in self.listbox_extra.curselection()]
            self.listbox_extra.delete(0, "end")
            for c in columnas:
                self.listbox_extra.insert("end", c)

            if seleccion_previa and any(c in columnas for c in seleccion_previa):
                # Restaura la selección anterior (si esas columnas siguen existiendo)
                for i, c in enumerate(columnas):
                    if c in seleccion_previa:
                        self.listbox_extra.selection_set(i)
            else:
                # Primera carga: preseleccionar por posición B(1), C(2), I(8) si existen
                # (0 = columna A, 1 = columna B, etc., asumiendo que el Excel arranca en A)
                posiciones_default = [1, 2, 8]
                for pos in posiciones_default:
                    if pos < len(columnas):
                        self.listbox_extra.selection_set(pos)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron leer las columnas:\n{e}")

    def elegir_carpeta(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta de archivos")
        if path:
            self.folder_path.set(path)

    @staticmethod
    def _formatear_valor(valor):
        # Si es una fecha/hora (Timestamp de pandas o datetime), se muestra solo como dd/mm/aa
        if isinstance(valor, (pd.Timestamp, datetime)):
            return valor.strftime("%d/%m/%y")
        return valor

    @staticmethod
    def _limpiar_numero_interno(valor):
        # Pandas suele leer columnas numéricas como float (14725.0). Si el valor
        # es un número entero "disfrazado" de float, sacar el ".0" para que
        # coincida con el nombre real del archivo (I14725, no I14725.0)
        if pd.isna(valor):
            return ""
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()

    def chequear(self):
        if self.df is None or not self.column_name.get():
            messagebox.showwarning("Faltan datos", "Elegí el Excel y la columna con el número interno.")
            return
        if not self.folder_path.get():
            messagebox.showwarning("Faltan datos", "Elegí la carpeta con los archivos.")
            return

        col = self.column_name.get()
        extra_indices = self.listbox_extra.curselection()
        extra_cols = [self.listbox_extra.get(i) for i in extra_indices]

        # Se queda con el número interno + columnas adicionales, sin nulos en el número interno
        cols_a_usar = [col] + [c for c in extra_cols if c != col]
        df_valido = self.df[cols_a_usar].dropna(subset=[col]).copy()
        df_valido[col] = df_valido[col].apply(self._limpiar_numero_interno)

        # Si hay números repetidos, se queda con la primera fila de cada uno (para los datos extra)
        df_unico = df_valido.drop_duplicates(subset=[col], keep="first")

        if df_unico.empty:
            messagebox.showwarning(
                "Sin datos",
                f"La columna '{col}' no tiene ningún valor válido para comparar.\n"
                "Revisá que sea la columna correcta y que la fila de títulos esté bien indicada.",
            )
            return

        # Se agrega el prefijo (ej: "I") para que coincida con el nombre de archivo
        prefijo = self.prefijo.get().strip()

        # Listado de archivos en la carpeta
        folder = self.folder_path.get()
        archivos = []
        if self.var_incluir_subcarpetas.get():
            for root_dir, _, files in os.walk(folder):
                archivos.extend(files)
        else:
            archivos = [
                f for f in os.listdir(folder)
                if os.path.isfile(os.path.join(folder, f))
            ]

        # Chequeo: cada número interno (con prefijo), ¿tiene al menos un archivo que empiece con él?
        faltantes = []  # lista de dicts con numero + columnas extra
        archivos_usados = set()
        for _, fila in df_unico.iterrows():
            numero_original = fila[col]
            numero_buscado = prefijo + numero_original
            matches = [a for a in archivos if a.upper().startswith(numero_buscado.upper())]
            if matches:
                archivos_usados.update(matches)
            else:
                registro = {"Numero_interno": numero_original, "Buscado_como": numero_buscado}
                for c in extra_cols:
                    if c != col:
                        registro[c] = self._formatear_valor(fila[c])
                faltantes.append(registro)

        # Archivos que no matchean ningún número
        extra = [a for a in archivos if a not in archivos_usados]

        self._faltantes = faltantes
        self._extra = extra

        # Mostrar resultado
        self.text_result.delete("1.0", "end")
        self.text_result.insert("end", f"Total de números internos en el Excel: {len(df_unico)}\n")
        self.text_result.insert("end", f"Total de archivos en la carpeta: {len(archivos)}\n\n")

        self.text_result.insert("end", f"❌ FALTANTES ({len(faltantes)}):\n")
        if faltantes:
            for reg in faltantes:
                detalle = ", ".join(f"{k}: {v}" for k, v in reg.items())
                self.text_result.insert("end", f"   - {detalle}\n")
        else:
            self.text_result.insert("end", "   (ninguno, todos los números tienen archivo)\n")

        if self.var_mostrar_extra.get():
            self.text_result.insert("end", f"\n⚠️ ARCHIVOS SIN NÚMERO CORRESPONDIENTE ({len(extra)}):\n")
            if extra:
                for a in extra:
                    self.text_result.insert("end", f"   - {a}\n")
            else:
                self.text_result.insert("end", "   (ninguno)\n")

        self.btn_exportar.config(state="normal")

    def exportar_reporte(self):
        path = filedialog.asksaveasfilename(
            title="Guardar reporte",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"reporte_chequeo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        )
        if not path:
            return

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame(self._faltantes).to_excel(writer, sheet_name="Faltantes", index=False)
            pd.DataFrame({"Archivos_sin_numero": self._extra}).to_excel(
                writer, sheet_name="Archivos_extra", index=False
            )

        messagebox.showinfo("Listo", f"Reporte guardado en:\n{path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ChequeoArchivosApp(root)
    root.mainloop()
