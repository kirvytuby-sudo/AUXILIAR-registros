# ─────────────────────────────────────────────────────────────────────────────
#  modules/impuestos_pf.py  —  Módulo "Impuestos PF" para gui_conciliacion.py
#  Mixin: ImpuestosPFMixin
#  Versión 1.0  —  Oct 2026
# ─────────────────────────────────────────────────────────────────────────────

import os, threading, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from calc_impuestos_pf import calcular_impuestos, generar_excel
    _ENGINE_OK = True
except ImportError:
    _ENGINE_OK = False


class ImpuestosPFMixin:
    """
    Módulo de Impuestos Persona Física para WorkspaceWindow.
    Agrega la pestaña '  📊 Impuestos PF  ' con:
      - Configuración: RFC, nombre, año
      - Fuentes: carpetas XMLs emitidos/recibidos
      - Fuentes: PDFs BBVA, Excels Banorte, Excels Inbursa, Excels extra
      - Botón Generar → calcular_impuestos + generar_excel
      - Visor inline (primeras filas del Excel generado)
      - Botones Abrir en Excel / Guardar como
    """

    # ── Pestaña principal ────────────────────────────────────────────────────
    def _tab_impuestos_pf(self, nb):
        """Construye la pestaña dentro del Notebook principal (nb)."""
        COLOR_F  = getattr(self, 'COLOR_FONDO',         '#FFF0F5')
        COLOR_AZ = getattr(self, 'COLOR_FUCSIA',        '#1E3A8A')
        COLOR_BL = getattr(self, 'COLOR_BLANCO',        '#FFFFFF')
        COLOR_DK = getattr(self, 'COLOR_FUCSIA_OSCURO', '#0F172A')

        frame = ttk.Frame(nb)
        nb.add(frame, text='  📊 Impuestos PF  ')

        # ── Barra inferior (resultado + visor) ─── pack PRIMERO
        bar = tk.Frame(frame, bg=COLOR_DK, height=160)
        bar.pack(side='bottom', fill='x')
        bar.pack_propagate(False)

        lbl_arch = tk.Label(bar, text='Sin resultado', bg=COLOR_DK, fg='#94A3B8',
                            font=('Arial',8), anchor='w')
        lbl_arch.pack(side='top', fill='x', padx=8, pady=(4,0))

        btn_bar = tk.Frame(bar, bg=COLOR_DK)
        btn_bar.pack(side='top', fill='x', padx=6, pady=2)
        tk.Button(btn_bar, text='📂 Abrir en Excel', command=self._ipf_abrir,
                  bg='#1D4ED8', fg='white', font=('Arial',8,'bold'),
                  relief='flat', padx=6, pady=2).pack(side='left', padx=2)
        tk.Button(btn_bar, text='💾 Guardar como', command=self._ipf_guardar_como,
                  bg='#374151', fg='white', font=('Arial',8,'bold'),
                  relief='flat', padx=6, pady=2).pack(side='left', padx=2)

        cols = ('Hoja','Fila','Col A','Col B','Col C','Col D')
        tv = ttk.Treeview(bar, columns=cols, show='headings', height=4)
        for c in cols:
            tv.heading(c, text=c); tv.column(c, width=110, anchor='w')
        tv.pack(fill='x', padx=6, pady=(0,4))

        self._ipf_lbl_arch = lbl_arch
        self._ipf_tv       = tv
        self._ipf_resultado = None

        # ── Panel de configuración y controles ─── pack DESPUÉS
        scroll_c = tk.Frame(frame, bg=COLOR_F)
        scroll_c.pack(side='top', fill='both', expand=True)

        canvas = tk.Canvas(scroll_c, bg=COLOR_F, highlightthickness=0)
        sb = ttk.Scrollbar(scroll_c, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = tk.Frame(canvas, bg=COLOR_F)
        win_id = canvas.create_window((0,0), window=inner, anchor='nw')

        def _on_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind('<Configure>', _on_resize)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        PAD = dict(padx=8, pady=3)

        # ── Título
        tk.Label(inner, text='📊  CÁLCULO IMPUESTOS — PERSONA FÍSICA  📊',
                 bg=COLOR_AZ, fg='white', font=('Arial',11,'bold'),
                 pady=8).pack(fill='x', **PAD)

        # ── Configuración básica
        cf = tk.LabelFrame(inner, text=' Configuración ', bg=COLOR_F, font=('Arial',9,'bold'))
        cf.pack(fill='x', **PAD)

        def row(parent, lbl, var, w=30):
            r = tk.Frame(parent, bg=COLOR_F); r.pack(fill='x', padx=4, pady=2)
            tk.Label(r, text=lbl, bg=COLOR_F, font=('Arial',9), width=18, anchor='e').pack(side='left')
            e = tk.Entry(r, textvariable=var, width=w, font=('Arial',9)); e.pack(side='left', padx=4)
            return e

        self._ipf_rfc     = tk.StringVar(value='GOFC7909293X6')
        self._ipf_nombre  = tk.StringVar(value='')
        self._ipf_year    = tk.StringVar(value=str(datetime.date.today().year - 1))

        row(cf, 'RFC contribuyente:', self._ipf_rfc, 20)
        row(cf, 'Nombre/Razón social:', self._ipf_nombre, 40)
        row(cf, 'Año fiscal:', self._ipf_year, 6)

        # ── XMLs Emitidos
        def _section(parent, titulo, lista_var, btn_carpeta_cmd, btn_arch_cmd):
            sf = tk.LabelFrame(parent, text=f' {titulo} ', bg=COLOR_F, font=('Arial',9,'bold'))
            sf.pack(fill='x', **PAD)
            bf = tk.Frame(sf, bg=COLOR_F); bf.pack(fill='x', padx=4, pady=2)
            tk.Button(bf, text='📁 Agregar carpeta', command=btn_carpeta_cmd,
                      bg=COLOR_AZ, fg='white', font=('Arial',8), relief='flat', padx=6).pack(side='left', padx=2)
            tk.Button(bf, text='📄 Agregar archivos', command=btn_arch_cmd,
                      bg='#374151', fg='white', font=('Arial',8), relief='flat', padx=6).pack(side='left', padx=2)
            tk.Button(bf, text='✕ Quitar selección', command=lambda: _quitar(lb),
                      bg='#7F1D1D', fg='white', font=('Arial',8), relief='flat', padx=6).pack(side='left', padx=2)
            lb = tk.Listbox(sf, listvariable=lista_var, height=3, font=('Arial',8),
                            selectmode='extended', bg='#0D1117', fg='#94A3B8')
            lb.pack(fill='x', padx=4, pady=(0,4))
            return lb

        def _quitar(lb):
            for i in reversed(lb.curselection()):
                lb.delete(i)

        self._ipf_emit_list   = tk.Variable(value=[])
        self._ipf_rec_list    = tk.Variable(value=[])
        self._ipf_bbva_list   = tk.Variable(value=[])
        self._ipf_banorte_list= tk.Variable(value=[])
        self._ipf_inbursa_list= tk.Variable(value=[])

        _section(inner, 'XMLs Emitidos (facturas propias)',
                 self._ipf_emit_list,
                 lambda: self._ipf_add_folder(self._ipf_emit_list),
                 lambda: self._ipf_add_files(self._ipf_emit_list, [('XML','*.xml'),('Todos','*.*')]))

        _section(inner, 'XMLs Recibidos (gastos/proveedores)',
                 self._ipf_rec_list,
                 lambda: self._ipf_add_folder(self._ipf_rec_list),
                 lambda: self._ipf_add_files(self._ipf_rec_list, [('XML','*.xml'),('Todos','*.*')]))

        _section(inner, 'Estados de cuenta BBVA (PDFs)',
                 self._ipf_bbva_list,
                 lambda: self._ipf_add_folder(self._ipf_bbva_list, ext='*.pdf'),
                 lambda: self._ipf_add_files(self._ipf_bbva_list, [('PDF','*.pdf')]))

        _section(inner, 'Estados de cuenta Banorte (Excel)',
                 self._ipf_banorte_list,
                 lambda: self._ipf_add_folder(self._ipf_banorte_list, ext='*.xls*'),
                 lambda: self._ipf_add_files(self._ipf_banorte_list, [('Excel','*.xlsx;*.xls')]))

        _section(inner, 'Estados de cuenta Inbursa (Excel)',
                 self._ipf_inbursa_list,
                 lambda: self._ipf_add_folder(self._ipf_inbursa_list, ext='*.xls*'),
                 lambda: self._ipf_add_files(self._ipf_inbursa_list, [('Excel','*.xlsx;*.xls')]))

        # ── Verificar vigencia
        vv_frame = tk.Frame(inner, bg=COLOR_F); vv_frame.pack(fill='x', padx=8, pady=2)
        self._ipf_verificar = tk.BooleanVar(value=True)
        tk.Checkbutton(vv_frame, text='🔍 Verificar vigencia de CFDIs en el SAT (excluye cancelados del cálculo)',
                       variable=self._ipf_verificar, bg=COLOR_F,
                       font=('Arial',9), fg='#1E3A8A').pack(side='left')

        # ── Progress bar
        pb_frame = tk.Frame(inner, bg=COLOR_F); pb_frame.pack(fill='x', padx=8, pady=4)
        self._ipf_pb  = ttk.Progressbar(pb_frame, mode='indeterminate', length=300)
        self._ipf_pbl = tk.Label(pb_frame, text='', bg=COLOR_F, font=('Arial',8), fg=COLOR_AZ)
        self._ipf_pb.pack(side='left', padx=(0,6))
        self._ipf_pbl.pack(side='left')

        # ── Botón Generar
        tk.Button(inner, text='⚡  GENERAR CÁLCULO DE IMPUESTOS',
                  command=self._ipf_generar,
                  bg='#14532D', fg='white', font=('Arial',10,'bold'),
                  relief='flat', pady=8).pack(fill='x', padx=8, pady=8)

    # ── Helpers de listas ────────────────────────────────────────────────────
    def _ipf_add_folder(self, lista_var, ext='*'):
        d = filedialog.askdirectory(title='Seleccionar carpeta')
        if not d: return
        cur = list(lista_var.get())
        if d not in cur: cur.append(d)
        lista_var.set(cur)

    def _ipf_add_files(self, lista_var, ftypes):
        files = filedialog.askopenfilenames(filetypes=ftypes, title='Seleccionar archivos')
        if not files: return
        cur = list(lista_var.get())
        for f in files:
            if f not in cur: cur.append(f)
        lista_var.set(cur)

    # ── Generar ─────────────────────────────────────────────────────────────
    def _ipf_generar(self):
        if not _ENGINE_OK:
            messagebox.showerror('Error',
                'No se encontró calc_impuestos_pf.py\n'
                'Asegúrate de que esté en la carpeta _app/')
            return
        rfc    = self._ipf_rfc.get().strip().upper()
        nombre = self._ipf_nombre.get().strip() or rfc
        try:
            year = int(self._ipf_year.get().strip())
        except ValueError:
            messagebox.showerror('Error', 'Año inválido'); return
        if not rfc:
            messagebox.showerror('Error', 'Ingresa el RFC'); return

        emit  = list(self._ipf_emit_list.get())
        rec   = list(self._ipf_rec_list.get())
        bbva  = list(self._ipf_bbva_list.get())
        ban   = list(self._ipf_banorte_list.get())
        inb   = list(self._ipf_inbursa_list.get())

        if not (emit or rec):
            messagebox.showwarning('Aviso', 'Agrega al menos una carpeta/archivo XML'); return

        # destino
        res_dir = self._resultados_dir(f'Impuestos_PF_{rfc}_{year}')
        os.makedirs(res_dir, exist_ok=True)
        out_path = os.path.join(res_dir, f'Impuestos_{rfc}_{year}.xlsx')

        self._pb_iniciar(self._ipf_pb, self._ipf_pbl)
        self._ipf_pbl.config(text='Calculando...')

        verificar = self._ipf_verificar.get()

        def hilo():
            try:
                res, emit_log, ded_log, no_ded_log = calcular_impuestos(
                    rfc=rfc, year=year,
                    carpetas_xml_emit=emit,
                    carpetas_xml_rec=rec,
                    pdfs_bbva=bbva if bbva else None,
                    excels_banorte=ban if ban else None,
                    excels_inbursa=inb if inb else None,
                    verificar_vigencia=verificar,
                    log=lambda t: self.after(0, self._log, t))
                generar_excel(res, emit_log, ded_log, no_ded_log,
                              rfc, nombre, year, out_path,
                              log=lambda t: self.after(0, self._log, t))
                self.after(0, self._ipf_set_resultado, out_path)
            except Exception as ex:
                import traceback
                msg = traceback.format_exc()
                self.after(0, self._log, f'✗ Error Impuestos PF: {ex}', True)
                self.after(0, self._log, msg, True)
            finally:
                self.after(0, self._pb_detener, self._ipf_pb, self._ipf_pbl, 'Listo')

        threading.Thread(target=hilo, daemon=True).start()

    # ── Actualizar resultado ─────────────────────────────────────────────────
    def _ipf_set_resultado(self, path):
        self._ipf_resultado = path
        self._ipf_lbl_arch.config(text=os.path.basename(path), fg='#86EFAC')
        self._log(f'✔ Generado: {path}', ok=True)
        self._ipf_cargar_visor(path)

    # ── Visor inline ─────────────────────────────────────────────────────────
    def _ipf_cargar_visor(self, path):
        tv = self._ipf_tv
        for i in tv.get_children(): tv.delete(i)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            filas = 0
            for sh in wb.sheetnames[:3]:
                ws = wb[sh]
                for ri, row in enumerate(ws.iter_rows(max_row=10, values_only=True), 1):
                    vals = [str(v)[:30] if v is not None else '' for v in row[:5]]
                    tv.insert('', 'end', values=[sh] + vals)
                    filas += 1
                    if filas >= 50: break
                if filas >= 50: break
            wb.close()
        except Exception as e:
            tv.insert('', 'end', values=['Error', str(e), '', '', '', ''])

    # ── Abrir / Guardar ──────────────────────────────────────────────────────
    def _ipf_abrir(self):
        if not self._ipf_resultado:
            messagebox.showinfo('Sin resultado', 'Genera primero el cálculo.'); return
        try:
            os.startfile(self._ipf_resultado)
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _ipf_guardar_como(self):
        if not self._ipf_resultado:
            messagebox.showinfo('Sin resultado', 'Genera primero el cálculo.'); return
        dest = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel','*.xlsx'),('Todos','*.*')],
            initialfile=os.path.basename(self._ipf_resultado),
            title='Guardar cálculo de impuestos como')
        if not dest: return
        import shutil
        try:
            shutil.copy2(self._ipf_resultado, dest)
            self._log(f'✔ Guardado en: {dest}', ok=True)
        except Exception as e:
            messagebox.showerror('Error', str(e))
