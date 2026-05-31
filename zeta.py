import os
import sys
import subprocess
import re
import time
import threading
import queue
import unicodedata
import multiprocessing
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importar worker solo cuando NO estamos en proceso hijo
if __name__ == '__main__' or not hasattr(sys, '_MEIPASS'):
    pass  # import normal

SKIP_DIRS = {'.git','.godot','node_modules','__pycache__','.vscode','.idea',
             'vnum','$RECYCLE.BIN','System Volume Information','.Trash'}
BINARY_EXT = ('.png','.jpg','.jpeg','.gif','.bmp','.ico','.dll','.exe','.so','.dylib',
              '.wav','.ogg','.mp3','.flac','.zip','.rar','.7z','.tar','.gz','.bz2',
              '.mp4','.avi','.mkv','.mov','.pdf','.doc','.docx','.xls','.xlsx',
              '.pyc','.pyo','.class','.o','.obj','.bin','.pack','.idx','.db',
              '.sqlite','.lnk','.pdb','.lib','.exp','.woff','.woff2','.ttf','.otf')
MAX_SIZE = 5 * 1024 * 1024
CHUNK_SIZE = 200  # archivos por bloque enviado al pool
N_WORKERS = max(2, multiprocessing.cpu_count())

# ── helpers ──────────────────────────────────────────────────────────────────
def _normalize(text: str) -> str:
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode().lower()

def _normalize_name(name: str) -> str:
    return re.sub(r'[^a-z0-9]', '', _normalize(name))

def make_regex_pattern(word: str) -> str:
    char_map = {
        'a':'[aAáÁàÀäÄ]','e':'[eEéÉèÈëË]','i':'[iIíÍìÌïÏ]',
        'o':'[oOóÓòÒöÖ]','u':'[uUúÚùÙüÜ]','n':'[nNñÑ]','c':'[cCçÇ]'
    }
    norm = _normalize(word)
    return ''.join(char_map.get(c, re.escape(c)) for c in norm)

# ── pool-worker (debe ser pickleable = nivel de módulo) ───────────────────────
def _search_chunk(args):
    """Retorna (resultados, cantidad_archivos_procesados) para tracking preciso."""
    paths, kws_norm, whole_word = args
    results = []
    for path in paths:
        try:
            size = os.path.getsize(path)
            if size == 0 or size > MAX_SIZE:
                continue
            with open(path, 'rb') as f:
                raw = f.read()
            if b'\x00' in raw[:512]:
                continue
            text = unicodedata.normalize('NFKD', raw.decode('utf-8', errors='ignore')
                                         ).encode('ASCII', 'ignore').decode().lower()
            found = True
            for kw in kws_norm:
                if whole_word:
                    if not re.search(r'(?<![a-z0-9])' + re.escape(kw) + r'(?![a-z0-9])', text):
                        found = False; break
                else:
                    if kw not in text:
                        found = False; break
            if found:
                results.append((text.count(kws_norm[0]), path))
        except Exception:
            pass
    return results, len(paths)

# ── lógica de búsqueda ────────────────────────────────────────────────────────
class SearchEngine:
    def __init__(self):
        self.cancel = False
        self._pool = None

    def stop(self):
        self.cancel = True
        if self._pool:
            self._pool.terminate()
            self._pool = None

    # ── recorrer disco iterativo (sin recursión) ──────────────────────────────
    def _walk_files(self, root, text_only=True):
        """Generador iterativo ultra-rápido con os.scandir."""
        stack = [root]
        while stack:
            current = stack.pop()
            if self.cancel:
                return
            try:
                with os.scandir(current) as it:
                    for e in it:
                        if self.cancel:
                            return
                        if e.is_dir(follow_symlinks=False):
                            if e.name.lower() not in SKIP_DIRS:
                                stack.append(e.path)
                        elif text_only:
                            if not e.name.lower().endswith(BINARY_EXT):
                                yield e.path
                        else:
                            yield e
            except PermissionError:
                pass

    # ── búsqueda por nombre ───────────────────────────────────────────────────
    def search_name(self, root, keywords, exact, match_cb, scan_cb):
        """scan_cb(path, scanned, total, matches) - total=-1 mientras recolecta"""
        self.cancel = False
        kws = [kw.lower() for kw in keywords] if exact else [_normalize_name(kw) for kw in keywords]
        n = 0
        matches = 0
        for entry in self._walk_files(root, text_only=False):
            if self.cancel:
                break
            n += 1
            name = entry.name.lower() if exact else _normalize_name(entry.name)
            if all(k in name for k in kws):
                score = 100 if name == kws[0] else 10
                match_cb((score, entry.path))
                matches += 1
            if n % 2000 == 0:
                scan_cb(entry.path, n, -1, matches)
        scan_cb('', n, n, matches)

    # ── búsqueda por contenido con multiprocessing ────────────────────────────
    def search_content(self, root, keywords, mode_and, whole_word, match_cb, scan_cb):
        """
        Fase 1: recolectar TODOS los archivos (rápido, solo os.scandir).
        Fase 2: procesar con pool. imap_unordered retorna cada bloque apenas termina.
        scan_cb(path, scanned, total, matches)
        """
        self.cancel = False
        kws_norm = [_normalize(kw) for kw in keywords]

        # ── FASE 1: recolectar todos los archivos ──────────────────────────────
        all_files = []
        for path in self._walk_files(root, text_only=True):
            if self.cancel:
                return
            all_files.append(path)
            # Actualizar UI durante recolección cada 5000 archivos
            if len(all_files) % 5000 == 0:
                scan_cb(path, 0, len(all_files), 0)

        total = len(all_files)
        if total == 0 or self.cancel:
            scan_cb('', 0, 0, 0)
            return

        # Emitir total real inmediatamente
        scan_cb(all_files[-1], 0, total, 0)

        # ── FASE 2: procesar en parallel con imap_unordered ───────────────────
        # Armar chunks
        chunks_args = [
            (all_files[i:i+CHUNK_SIZE], kws_norm, whole_word)
            for i in range(0, total, CHUNK_SIZE)
        ]

        self._pool = multiprocessing.Pool(processes=N_WORKERS)
        pool = self._pool
        done = 0
        matches = 0
        last_path = all_files[-1]

        try:
            # imap_unordered: cada bloque terminado se entrega inmediatamente
            for results, chunk_len in pool.imap_unordered(_search_chunk, chunks_args, chunksize=1):
                if self.cancel:
                    break
                done += chunk_len
                matches += len(results)
                for match in results:
                    match_cb(match)
                # Reportar progreso real con total conocido
                scan_cb(last_path, done, total, matches)
        finally:
            pool.close()
            pool.join()
            self._pool = None
            if not self.cancel:
                scan_cb('', done, total, matches)


# ── GUI ───────────────────────────────────────────────────────────────────────
class ZetaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zeta")
        self.root.geometry("1100x650")
        self.root.minsize(900, 500)

        if sys.platform == "win32":
            import ctypes
            try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('zeta.search.app')
            except: pass

        base = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        ico = os.path.join(base, "buscador.ico")
        if os.path.exists(ico):
            self.root.iconbitmap(ico)

        self.engine = SearchEngine()
        self.q = queue.Queue()
        self.is_searching = False
        self.last_query = ""
        self.last_type = "nombre"
        self.start_time = 0
        self.stats = {'scanned': 0, 'matches': 0, 'last_path': ''}
        self.match_indices = []
        self.cur_match = 0
        self.matches_shown = 0

        self._build_ui()
        self.root.after(80, self._poll_queue)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        s = ttk.Style()
        s.configure("Status.TLabel", font=("Segoe UI", 10, "bold"))
        s.configure("Stats.TLabel", font=("Segoe UI", 9), foreground="#333")

        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Carpeta Base:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
        self.var_path = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_path, font=("Segoe UI", 10)).grid(row=0, column=1, sticky=tk.EW, padx=8)
        ttk.Button(top, text="Examinar…", command=self._browse).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar texto:", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=6)
        self.var_q = tk.StringVar()
        eq = ttk.Entry(top, textvariable=self.var_q, font=("Segoe UI", 10))
        eq.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=6)
        eq.bind("<Return>", lambda e: self._start())
        self.btn_search = ttk.Button(top, text="Buscar", command=self._start)
        self.btn_search.grid(row=1, column=2)

        opts = ttk.LabelFrame(top, text="Opciones", padding=8)
        opts.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=8)
        self.var_type = tk.StringVar(value="nombre")
        ttk.Radiobutton(opts, text="Archivo / Carpeta", variable=self.var_type, value="nombre").pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(opts, text="Solo Contenido",    variable=self.var_type, value="contenido").pack(side=tk.LEFT, padx=8)
        ttk.Separator(opts, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        self.var_exact = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Nombre Exacto",   variable=self.var_exact).pack(side=tk.LEFT, padx=8)
        self.var_whole = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Palabra Completa", variable=self.var_whole).pack(side=tk.LEFT, padx=8)
        ttk.Separator(opts, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        self.var_mode = tk.StringVar(value="and")
        ttk.Radiobutton(opts, text="Todas",  variable=self.var_mode, value="and").pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(opts, text="Alguna", variable=self.var_mode, value="or").pack(side=tk.LEFT, padx=4)

        # Status bar
        sb = ttk.Frame(self.root, padding=4)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        self.var_status = tk.StringVar(value="Listo.")
        ttk.Label(sb, textvariable=self.var_status, style="Status.TLabel").pack(side=tk.LEFT, padx=4)
        self.var_stats = tk.StringVar(value="")
        ttk.Label(sb, textvariable=self.var_stats, style="Stats.TLabel").pack(side=tk.RIGHT, padx=8)

        # Paned
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))

        left = ttk.Frame(paned)
        paned.add(left, weight=2)
        cols = ("nombre","ruta","tipo")
        self.tree = ttk.Treeview(left, columns=cols, show="headings")
        for c, w, t, stretch in (("nombre", 150, "Nombre", False),
                                 ("ruta", 250, "Ruta", True),
                                 ("tipo", 70, "Tipo", False)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=tk.W, stretch=stretch)
        self.tree.tag_configure("noresult", foreground="red")
        sb_tree = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb_tree.set)
        sb_tree.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right = ttk.LabelFrame(paned, text="Vista Previa", padding=4)
        paned.add(right, weight=1)
        nav = ttk.Frame(right)
        nav.pack(fill=tk.X, pady=(0,4))
        self.btn_prev = ttk.Button(nav, text="< Anterior", command=self._prev, state=tk.DISABLED)
        self.btn_prev.pack(side=tk.LEFT, padx=4)
        self.lbl_nav = ttk.Label(nav, text="0 / 0", font=("Segoe UI",9,"bold"))
        self.lbl_nav.pack(side=tk.LEFT, expand=True)
        self.btn_next = ttk.Button(nav, text="Siguiente >", command=self._next, state=tk.DISABLED)
        self.btn_next.pack(side=tk.RIGHT, padx=4)
        self.txt = tk.Text(right, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas",10), width=40)
        sb_txt = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb_txt.set)
        sb_txt.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.txt.tag_config("match",  background="#a8e6cf", foreground="black")
        self.txt.tag_config("active", background="#22c55e", foreground="white",
                            font=("Consolas",10,"bold"))

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self._open())
        self.tree.bind("<Button-3>", self._ctx_menu)

        cm = tk.Menu(self.root, tearoff=0)
        cm.add_command(label="Abrir",           command=self._open)
        cm.add_command(label="Abrir ubicación", command=self._open_loc)
        cm.add_command(label="Copiar ruta",     command=self._copy_path)
        self.cm = cm

    # ── helpers UI ────────────────────────────────────────────────────────────
    def _browse(self):
        d = filedialog.askdirectory()
        if d: self.var_path.set(d)

    def _get_sel_path(self):
        sel = self.tree.selection()
        if not sel: return None
        v = self.tree.item(sel[0])['values']
        return v[1] if v and v[1] else None

    # ── animación en status bar ───────────────────────────────────────────────
    def _animate(self):
        if not self.is_searching:
            return
        elapsed = time.time() - self.start_time
        sc    = self.stats['scanned']
        tot   = self.stats['total']
        mat   = self.stats['matches']
        phase = self.stats.get('phase', 'recolectando')

        if phase == 'recolectando':
            tot_str = f"{tot} detectados…"
        else:
            tot_str = f"/ {tot}" if tot > 0 else ''

        self.var_stats.set(f"⏱ {elapsed:.0f}s  │  Leídos: {sc} {tot_str}  │  Hallados: {mat}")
        self.root.after(200, self._animate)

    # ── poll queue ────────────────────────────────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "match":
                    score, path = data
                    if self.matches_shown < 10000:
                        self.tree.insert("", tk.END,
                            values=(os.path.basename(path), path,
                                    "Carpeta" if os.path.isdir(path) else "Archivo"))
                        self.matches_shown += 1
                elif kind == "scan":
                    path, sc, tot, mc = data
                    self.stats.update({
                        'last_path': path,
                        'scanned':   sc,
                        'total':     tot,
                        'matches':   mc,
                        # total=-1 significa que aún estamos en fase de recolección
                        'phase': 'recolectando' if tot < 0 else 'procesando'
                    })
                elif kind == "status":
                    self.var_status.set(data)
                elif kind == "done":
                    self._finish(data)
                elif kind == "preview_raw":
                    self._show_raw(data)
                elif kind == "preview_hl":
                    self._show_highlight(data)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_queue)

    # ── start / stop ──────────────────────────────────────────────────────────
    def _start(self):
        if self.is_searching:
            self.engine.stop()
            self.btn_search.config(text="Cancelando…", state=tk.DISABLED)
            return

        base = self.var_path.get().strip()
        qry  = self.var_q.get().strip()
        if not base or not os.path.isdir(base):
            messagebox.showwarning("Zeta", "Selecciona una carpeta válida.")
            return
        if not qry:
            messagebox.showwarning("Zeta", "Ingresa texto a buscar.")
            return

        self.last_query = qry
        self.last_type  = self.var_type.get()
        for i in self.tree.get_children(): self.tree.delete(i)
        self._show_raw("")
        self.is_searching = True
        self.start_time = time.time()
        self.stats = {'scanned': 0, 'matches': 0, 'total': 0, 'last_path': '', 'phase': 'recolectando'}
        self.matches_shown = 0
        self.btn_search.config(text="Cancelar")
        self._animate()

        kws  = [k.strip() for k in qry.split() if k.strip()]
        mode_and = self.var_mode.get() == "and"
        exact    = self.var_exact.get()
        whole    = self.var_whole.get()
        typ      = self.var_type.get()

        threading.Thread(target=self._run, args=(base, kws, typ, exact, whole, mode_and),
                         daemon=True).start()

    def _run(self, base, kws, typ, exact, whole, mode_and):
        self.q.put(("status", "Buscando…"))
        try:
            def match_cb(m):  self.q.put(("match", m))
            def scan_cb(p, sc, tot, mc): self.q.put(("scan", (p, sc, tot, mc)))

            if typ == "nombre":
                self.engine.search_name(base, kws, exact, match_cb, scan_cb)
            else:
                self.engine.search_content(base, kws, mode_and, whole, match_cb, scan_cb)
            self.q.put(("done", not self.engine.cancel))
        except Exception as e:
            self.q.put(("status", f"Error: {e}"))
            self.q.put(("done", False))

    def _finish(self, completed):
        self.is_searching = False
        self.btn_search.config(text="Buscar", state=tk.NORMAL)
        elapsed = time.time() - self.start_time
        m   = self.stats['matches']
        tot = self.stats['total']
        sc  = self.stats['scanned']
        if not completed:
            self.var_status.set("Búsqueda cancelada.")
        elif m == 0:
            self.tree.insert("", tk.END, values=("❌ Sin resultados","",""), tags=("noresult",))
            self.var_status.set("No se encontraron coincidencias.")
        else:
            self.var_status.set(f"✅ {m} resultado(s) en {elapsed:.1f}s")
        self.var_stats.set(f"Total detectados: {tot}  │  Leídos: {sc}  │  Hallados: {m}  │  Tiempo: {elapsed:.1f}s")

    # ── preview ───────────────────────────────────────────────────────────────
    def _reset_nav(self):
        self.match_indices = []
        self.cur_match = 0
        self.lbl_nav.config(text="0 / 0")
        self.btn_prev.config(state=tk.DISABLED)
        self.btn_next.config(state=tk.DISABLED)

    def _show_raw(self, text):
        self._reset_nav()
        self.txt.config(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.insert(tk.END, text)
        self.txt.config(state=tk.DISABLED)

    def _show_highlight(self, data):
        self._reset_nav()
        self.txt.config(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        if data.get('prefix'): self.txt.insert(tk.END, data['prefix'])
        self.txt.insert(tk.END, data['content'])

        for kw in data['keywords']:
            pat = make_regex_pattern(kw)
            pos = "1.0"
            while True:
                cv = tk.StringVar()
                pos = self.txt.search(pat, pos, stopindex=tk.END, regexp=True, nocase=True, count=cv)
                if not pos: break
                ln = cv.get()
                if not ln: break
                end = f"{pos}+{ln}c"
                self.txt.tag_add("match", pos, end)
                self.match_indices.append((pos, end))
                pos = end

        self.txt.config(state=tk.DISABLED)
        self.match_indices.sort(key=lambda x: [int(n) for n in x[0].split('.')])

        if self.match_indices:
            self.cur_match = 0
            self._update_nav()
            self._hl_current()

    def _update_nav(self):
        n = len(self.match_indices)
        self.lbl_nav.config(text=f"{self.cur_match+1} / {n}")
        st = tk.NORMAL if n > 1 else tk.DISABLED
        self.btn_prev.config(state=st)
        self.btn_next.config(state=st)

    def _hl_current(self):
        if not self.match_indices: return
        self.txt.tag_remove("active","1.0",tk.END)
        s, e = self.match_indices[self.cur_match]
        self.txt.tag_add("active", s, e)
        self.txt.see(s)

    def _prev(self):
        if not self.match_indices: return
        self.cur_match = (self.cur_match - 1) % len(self.match_indices)
        self._update_nav(); self._hl_current()

    def _next(self):
        if not self.match_indices: return
        self.cur_match = (self.cur_match + 1) % len(self.match_indices)
        self._update_nav(); self._hl_current()

    def _on_select(self, _):
        path = self._get_sel_path()
        if not path or not os.path.exists(path):
            self._show_raw(""); return
        self._show_raw("Cargando…")
        threading.Thread(target=self._load_preview, args=(path,), daemon=True).start()

    def _load_preview(self, path):
        try:
            if os.path.isdir(path):
                items = os.listdir(path)
                txt = f"📁 {os.path.basename(path)}  ({len(items)} elementos)\n\n"
                for it in items[:40]: txt += f"  {it}\n"
                if len(items) > 40: txt += f"\n  … y {len(items)-40} más."
                self.q.put(("preview_raw", txt)); return

            if path.lower().endswith(('.dll','.exe','.pyc','.png','.jpg','.mp3','.zip','.rar')):
                self.q.put(("preview_raw", f"Archivo binario: {os.path.basename(path)}")); return

            size = os.path.getsize(path)
            kws  = [k.strip() for k in self.last_query.split() if k.strip()] if self.last_query else []

            # Encontrar offset del primer match
            offset = 0
            if kws:
                kw_norm = unicodedata.normalize('NFKD', kws[0]).encode('ASCII','ignore').decode().lower()
                try:
                    with open(path,'rb') as f:
                        raw = f.read(MAX_SIZE)
                    txt_norm = unicodedata.normalize('NFKD', raw.decode('utf-8','ignore')
                                                     ).encode('ASCII','ignore').decode().lower()
                    idx = txt_norm.find(kw_norm)
                    if idx > 0:
                        # Aproximar byte offset (relación ~1:1 para ASCII)
                        offset = max(0, idx - 500)
                except Exception:
                    pass

            start = max(0, offset - 8000)
            with open(path,'rb') as f:
                f.seek(start)
                content = f.read(50000).decode('utf-8', errors='ignore')

            prefix = "…\n" if start > 0 else ""
            suffix = "\n\n… (truncado) …" if start + 50000 < size else ""

            if kws:
                self.q.put(("preview_hl", {"content": content+suffix, "keywords": kws, "prefix": prefix}))
            else:
                self.q.put(("preview_raw", prefix + content + suffix))
        except Exception as ex:
            self.q.put(("preview_raw", f"Error: {ex}"))

    # ── context menu / open ───────────────────────────────────────────────────
    def _ctx_menu(self, e):
        row = self.tree.identify_row(e.y)
        if row:
            self.tree.selection_set(row)
            self.cm.tk_popup(e.x_root, e.y_root)

    def _open(self):
        p = self._get_sel_path()
        if p and os.path.exists(p):
            os.startfile(p) if sys.platform == "win32" else subprocess.run(['open', p])

    def _open_loc(self):
        p = self._get_sel_path()
        if p and os.path.exists(p):
            subprocess.Popen(['explorer', '/select,', os.path.normpath(p)])

    def _copy_path(self):
        p = self._get_sel_path()
        if p:
            self.root.clipboard_clear()
            self.root.clipboard_append(p)
            self.var_status.set(f"Copiado: {p}")


if __name__ == '__main__':
    multiprocessing.freeze_support()   # obligatorio para PyInstaller + multiprocessing
    root = tk.Tk()
    ZetaApp(root)
    root.mainloop()
