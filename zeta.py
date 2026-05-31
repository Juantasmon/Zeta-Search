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

TRANSLATIONS = {
    'es': {
        'title': 'Zeta',
        'lbl_base': 'Carpeta Base:',
        'lbl_search': 'Buscar texto:',
        'btn_browse': 'Examinar…',
        'btn_search': 'Buscar',
        'btn_cancel': 'Cancelar',
        'btn_canceling': 'Cancelando…',
        'lbl_options': 'Opciones',
        'opt_type_name': 'Archivo / Carpeta',
        'opt_type_content': 'Solo Contenido',
        'chk_exact': 'Nombre Exacto',
        'chk_whole': 'Palabra Completa',
        'opt_mode_all': 'Todas',
        'opt_mode_any': 'Alguna',
        'lbl_preview': 'Vista Previa',
        'btn_prev': '< Anterior',
        'btn_next': 'Siguiente >',
        'col_name': 'Nombre',
        'col_path': 'Ruta',
        'col_type': 'Tipo',
        'type_folder': 'Carpeta',
        'type_file': 'Archivo',
        'status_ready': 'Listo.',
        'status_searching': 'Buscando…',
        'status_canceled': 'Búsqueda cancelada.',
        'status_no_results': 'No se encontraron coincidencias.',
        'status_results': '✅ {count} resultado(s) en {time:.1f}s',
        'stats_done': 'Total detectados: {total}  │  Leídos: {scanned}  │  Hallados: {matches}  │  Tiempo: {time:.1f}s',
        'stats_running': '⏱ {elapsed:.0f}s  │  Leídos: {scanned} {total_str}  │  Hallados: {matches}',
        'stats_detecting': '{total} detectados…',
        'msg_valid_folder': 'Selecciona una carpeta válida.',
        'msg_valid_query': 'Ingresa texto a buscar.',
        'menu_open': 'Abrir',
        'menu_open_loc': 'Abrir ubicación',
        'menu_copy_path': 'Copiar ruta',
        'status_copied': 'Copiado: {path}',
        'preview_loading': 'Cargando…',
        'preview_folder': '📁 {name}  ({count} elementos)\n\n',
        'preview_more': '\n  … y {count} más.',
        'preview_binary': 'Archivo binario: {name}',
        'preview_truncated': '\n\n… (truncado) …',
        'preview_error': 'Error: {error}',
        'tree_no_results': '❌ Sin resultados'
    },
    'en': {
        'title': 'Zeta',
        'lbl_base': 'Base Folder:',
        'lbl_search': 'Search text:',
        'btn_browse': 'Browse…',
        'btn_search': 'Search',
        'btn_cancel': 'Cancel',
        'btn_canceling': 'Canceling…',
        'lbl_options': 'Options',
        'opt_type_name': 'File / Folder',
        'opt_type_content': 'Content Only',
        'chk_exact': 'Exact Name',
        'chk_whole': 'Whole Word',
        'opt_mode_all': 'All',
        'opt_mode_any': 'Any',
        'lbl_preview': 'Preview',
        'btn_prev': '< Previous',
        'btn_next': 'Next >',
        'col_name': 'Name',
        'col_path': 'Path',
        'col_type': 'Type',
        'type_folder': 'Folder',
        'type_file': 'File',
        'status_ready': 'Ready.',
        'status_searching': 'Searching…',
        'status_canceled': 'Search canceled.',
        'status_no_results': 'No matches found.',
        'status_results': '✅ {count} match(es) in {time:.1f}s',
        'stats_done': 'Total detected: {total}  │  Read: {scanned}  │  Found: {matches}  │  Time: {time:.1f}s',
        'stats_running': '⏱ {elapsed:.0f}s  │  Read: {scanned} {total_str}  │  Found: {matches}',
        'stats_detecting': '{total} detected…',
        'msg_valid_folder': 'Select a valid folder.',
        'msg_valid_query': 'Enter text to search.',
        'menu_open': 'Open',
        'menu_open_loc': 'Open location',
        'menu_copy_path': 'Copy path',
        'status_copied': 'Copied: {path}',
        'preview_loading': 'Loading…',
        'preview_folder': '📁 {name}  ({count} items)\n\n',
        'preview_more': '\n  … and {count} more.',
        'preview_binary': 'Binary file: {name}',
        'preview_truncated': '\n\n… (truncated) …',
        'preview_error': 'Error: {error}',
        'tree_no_results': '❌ No results'
    }
}

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

        self.language = "es"
        self.theme = "light"
        self._load_settings()

        # Load custom theme image
        base = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        self.theme_img_path = os.path.join(base, "theme.png")
        self.theme_photo_light = None
        self.theme_photo_dark = None
        self._load_theme_images()

        self._build_ui()
        self.root.after(80, self._poll_queue)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        self.lbl_base_folder = ttk.Label(top, text="Carpeta Base:", font=("Segoe UI", 10, "bold"))
        self.lbl_base_folder.grid(row=0, column=0, sticky=tk.W)
        self.var_path = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_path, font=("Segoe UI", 10)).grid(row=0, column=1, sticky=tk.EW, padx=8)
        self.btn_browse = ttk.Button(top, text="Examinar…", command=self._browse)
        self.btn_browse.grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        self.lbl_search_text = ttk.Label(top, text="Buscar texto:", font=("Segoe UI", 10, "bold"))
        self.lbl_search_text.grid(row=1, column=0, sticky=tk.W, pady=6)
        self.var_q = tk.StringVar()
        eq = ttk.Entry(top, textvariable=self.var_q, font=("Segoe UI", 10))
        eq.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=6)
        eq.bind("<Return>", lambda e: self._start())
        self.btn_search = ttk.Button(top, text="Buscar", command=self._start)
        self.btn_search.grid(row=1, column=2)

        self.lbl_opts = ttk.LabelFrame(top, text="Opciones", padding=8)
        self.lbl_opts.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=8)
        self.var_type = tk.StringVar(value="nombre")
        self.rb_type_name = ttk.Radiobutton(self.lbl_opts, text="Archivo / Carpeta", variable=self.var_type, value="nombre")
        self.rb_type_name.pack(side=tk.LEFT, padx=8)
        self.rb_type_content = ttk.Radiobutton(self.lbl_opts, text="Solo Contenido",    variable=self.var_type, value="contenido")
        self.rb_type_content.pack(side=tk.LEFT, padx=8)
        ttk.Separator(self.lbl_opts, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        self.var_exact = tk.BooleanVar()
        self.chk_exact = ttk.Checkbutton(self.lbl_opts, text="Nombre Exacto",   variable=self.var_exact)
        self.chk_exact.pack(side=tk.LEFT, padx=8)
        self.var_whole = tk.BooleanVar()
        self.chk_whole = ttk.Checkbutton(self.lbl_opts, text="Palabra Completa", variable=self.var_whole)
        self.chk_whole.pack(side=tk.LEFT, padx=8)
        ttk.Separator(self.lbl_opts, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        self.var_mode = tk.StringVar(value="and")
        self.rb_mode_all = ttk.Radiobutton(self.lbl_opts, text="Todas",  variable=self.var_mode, value="and")
        self.rb_mode_all.pack(side=tk.LEFT, padx=4)
        self.rb_mode_any = ttk.Radiobutton(self.lbl_opts, text="Alguna", variable=self.var_mode, value="or")
        self.rb_mode_any.pack(side=tk.LEFT, padx=4)

        # Theme button (small, modern)
        self.btn_theme = ttk.Button(self.lbl_opts, command=self._toggle_theme)
        self.btn_theme.pack(side=tk.RIGHT, padx=8)
        
        # Language Selector Combobox
        lang_display = "Español" if self.language == "es" else "English"
        self.var_lang = tk.StringVar(value=lang_display)
        self.cb_lang = ttk.Combobox(self.lbl_opts, textvariable=self.var_lang, values=["Español", "English"], width=8, state="readonly")
        self.cb_lang.pack(side=tk.RIGHT, padx=8)
        self.cb_lang.bind("<<ComboboxSelected>>", lambda e: self._on_language_change())

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

        self.frame_preview = ttk.LabelFrame(paned, text="Vista Previa", padding=4)
        paned.add(self.frame_preview, weight=1)
        nav = ttk.Frame(self.frame_preview)
        nav.pack(fill=tk.X, pady=(0,4))
        self.btn_prev = ttk.Button(nav, text="< Anterior", command=self._prev, state=tk.DISABLED)
        self.btn_prev.pack(side=tk.LEFT, padx=4)
        self.lbl_nav = ttk.Label(nav, text="0 / 0", font=("Segoe UI",9,"bold"))
        self.lbl_nav.pack(side=tk.LEFT, expand=True)
        self.btn_next = ttk.Button(nav, text="Siguiente >", command=self._next, state=tk.DISABLED)
        self.btn_next.pack(side=tk.RIGHT, padx=4)
        self.txt = tk.Text(self.frame_preview, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas",10), width=40)
        sb_txt = ttk.Scrollbar(self.frame_preview, orient=tk.VERTICAL, command=self.txt.yview)
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

        self._apply_theme()
        self._update_language()

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
        t = TRANSLATIONS[self.language]
        elapsed = time.time() - self.start_time
        sc    = self.stats['scanned']
        tot   = self.stats['total']
        mat   = self.stats['matches']
        phase = self.stats.get('phase', 'recolectando')

        if phase == 'recolectando':
            tot_str = t['stats_detecting'].format(total=tot)
        else:
            tot_str = f"/ {tot}" if tot > 0 else ''

        self.var_stats.set(t['stats_running'].format(elapsed=elapsed, scanned=sc, total_str=tot_str, matches=mat))
        self.root.after(200, self._animate)

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "match":
                    score, path = data
                    if self.matches_shown < 10000:
                        t = TRANSLATIONS[self.language]
                        self.tree.insert("", tk.END,
                            values=(os.path.basename(path), path,
                                    t['type_folder'] if os.path.isdir(path) else t['type_file']))
                        self.matches_shown += 1
                elif kind == "scan":
                    path, sc, tot, mc = data
                    self.stats.update({
                        'last_path': path,
                        'scanned':   sc,
                        'total':     tot,
                        'matches':   mc,
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

    def _start(self):
        if self.is_searching:
            self.engine.stop()
            t = TRANSLATIONS[self.language]
            self.btn_search.config(text=t['btn_canceling'], state=tk.DISABLED)
            return

        base = self.var_path.get().strip()
        qry  = self.var_q.get().strip()
        t = TRANSLATIONS[self.language]
        if not base or not os.path.isdir(base):
            messagebox.showwarning("Zeta", t['msg_valid_folder'])
            return
        if not qry:
            messagebox.showwarning("Zeta", t['msg_valid_query'])
            return

        self.last_query = qry
        self.last_type  = self.var_type.get()
        for i in self.tree.get_children(): self.tree.delete(i)
        self._show_raw("")
        self.is_searching = True
        self.start_time = time.time()
        self.stats = {'scanned': 0, 'matches': 0, 'total': 0, 'last_path': '', 'phase': 'recolectando', 'elapsed_time': 0}
        self.matches_shown = 0
        self.btn_search.config(text=t['btn_cancel'])
        self._animate()

        kws  = [k.strip() for k in qry.split() if k.strip()]
        mode_and = self.var_mode.get() == "and"
        exact    = self.var_exact.get()
        whole    = self.var_whole.get()
        typ      = self.var_type.get()

        threading.Thread(target=self._run, args=(base, kws, typ, exact, whole, mode_and),
                         daemon=True).start()

    def _run(self, base, kws, typ, exact, whole, mode_and):
        t = TRANSLATIONS[self.language]
        self.q.put(("status", t['status_searching']))
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
        t = TRANSLATIONS[self.language]
        self.btn_search.config(text=t['btn_search'], state=tk.NORMAL)
        elapsed = time.time() - self.start_time
        self.stats['elapsed_time'] = elapsed
        m   = self.stats['matches']
        tot = self.stats['total']
        sc  = self.stats['scanned']
        if not completed:
            self.var_status.set(t['status_canceled'])
        elif m == 0:
            self.tree.insert("", tk.END, values=(t['tree_no_results'],"",""), tags=("noresult",))
            self.var_status.set(t['status_no_results'])
        else:
            self.var_status.set(t['status_results'].format(count=m, time=elapsed))
        self.var_stats.set(t['stats_done'].format(total=tot, scanned=sc, matches=m, time=elapsed))

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
        t = TRANSLATIONS[self.language]
        self._show_raw(t['preview_loading'])
        threading.Thread(target=self._load_preview, args=(path,), daemon=True).start()

    def _load_preview(self, path):
        t = TRANSLATIONS[self.language]
        try:
            if os.path.isdir(path):
                items = os.listdir(path)
                txt = t['preview_folder'].format(name=os.path.basename(path), count=len(items))
                for it in items[:40]: txt += f"  {it}\n"
                if len(items) > 40: txt += t['preview_more'].format(count=len(items)-40)
                self.q.put(("preview_raw", txt)); return

            if path.lower().endswith(('.dll','.exe','.pyc','.png','.jpg','.mp3','.zip','.rar')):
                self.q.put(("preview_raw", t['preview_binary'].format(name=os.path.basename(path)))); return

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
            suffix = t['preview_truncated'] if start + 50000 < size else ""

            if kws:
                self.q.put(("preview_hl", {"content": content+suffix, "keywords": kws, "prefix": prefix}))
            else:
                self.q.put(("preview_raw", prefix + content + suffix))
        except Exception as ex:
            self.q.put(("preview_raw", t['preview_error'].format(error=ex)))

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
            t = TRANSLATIONS[self.language]
            self.var_status.set(t['status_copied'].format(path=p))

    # ── theme and language helpers ────────────────────────────────────────────
    def _on_language_change(self):
        lang = "es" if self.var_lang.get() == "Español" else "en"
        self.language = lang
        self._update_language()
        self._save_settings()

    def _toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self._apply_theme()
        self._save_settings()

    def _load_settings(self):
        import json
        self.config_dir = os.path.join(os.path.expanduser('~'), '.zeta_search')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        self.language = "es"
        self.theme = "light"
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    self.language = cfg.get('language', 'es')
                    self.theme = cfg.get('theme', 'light')
        except Exception:
            pass

    def _save_settings(self):
        import json
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({'language': self.language, 'theme': self.theme}, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def _load_theme_images(self):
        from PIL import Image, ImageOps, ImageTk
        try:
            if os.path.exists(self.theme_img_path):
                img = Image.open(self.theme_img_path).convert("RGBA")
                # Light mode: original color (usually black/dark icon)
                img_light = img.resize((20, 20), Image.Resampling.LANCZOS)
                self.theme_photo_light = ImageTk.PhotoImage(img_light)
                
                # Dark mode: inverted colors (white/light icon)
                r, g, b, a = img.split()
                rgb_img = Image.merge("RGB", (r, g, b))
                inverted_rgb = ImageOps.invert(rgb_img)
                ir, ig, ib = inverted_rgb.split()
                img_dark = Image.merge("RGBA", (ir, ig, ib, a)).resize((20, 20), Image.Resampling.LANCZOS)
                self.theme_photo_dark = ImageTk.PhotoImage(img_dark)
        except Exception:
            pass

    def _update_language(self):
        lang = "es" if self.var_lang.get() == "Español" else "en"
        self.language = lang
        t = TRANSLATIONS[lang]
        
        self.root.title(t['title'])
        self.lbl_base_folder.config(text=t['lbl_base'])
        self.btn_browse.config(text=t['btn_browse'])
        self.lbl_search_text.config(text=t['lbl_search'])
        
        if self.is_searching:
            curr_text = self.btn_search.cget('text')
            if curr_text in (TRANSLATIONS['es']['btn_canceling'], TRANSLATIONS['en']['btn_canceling']):
                self.btn_search.config(text=t['btn_canceling'])
            else:
                self.btn_search.config(text=t['btn_cancel'])
        else:
            self.btn_search.config(text=t['btn_search'])
            
        self.lbl_opts.config(text=t['lbl_options'])
        self.rb_type_name.config(text=t['opt_type_name'])
        self.rb_type_content.config(text=t['opt_type_content'])
        self.chk_exact.config(text=t['chk_exact'])
        self.chk_whole.config(text=t['chk_whole'])
        self.rb_mode_all.config(text=t['opt_mode_all'])
        self.rb_mode_any.config(text=t['opt_mode_any'])
        
        curr_status = self.var_status.get()
        if curr_status in (TRANSLATIONS['es']['status_ready'], TRANSLATIONS['en']['status_ready']):
            self.var_status.set(t['status_ready'])
        elif curr_status in (TRANSLATIONS['es']['status_searching'], TRANSLATIONS['en']['status_searching']):
            self.var_status.set(t['status_searching'])
        elif curr_status in (TRANSLATIONS['es']['status_canceled'], TRANSLATIONS['en']['status_canceled']):
            self.var_status.set(t['status_canceled'])
        elif curr_status in (TRANSLATIONS['es']['status_no_results'], TRANSLATIONS['en']['status_no_results']):
            self.var_status.set(t['status_no_results'])
        elif "resultado(s) en" in curr_status or "match(es) in" in curr_status:
            m = self.stats.get('matches', 0)
            elapsed = self.stats.get('elapsed_time', 0)
            self.var_status.set(t['status_results'].format(count=m, time=elapsed))
        elif curr_status.startswith("Copiado: ") or curr_status.startswith("Copied: "):
            path = curr_status.split(": ", 1)[1] if ": " in curr_status else ""
            self.var_status.set(t['status_copied'].format(path=path))
            
        if self.is_searching:
            elapsed = time.time() - self.start_time
            sc = self.stats['scanned']
            tot = self.stats['total']
            mat = self.stats['matches']
            phase = self.stats.get('phase', 'recolectando')
            if phase == 'recolectando':
                tot_str = t['stats_detecting'].format(total=tot)
            else:
                tot_str = f"/ {tot}" if tot > 0 else ''
            self.var_stats.set(t['stats_running'].format(elapsed=elapsed, scanned=sc, total_str=tot_str, matches=mat))
        else:
            m = self.stats.get('matches', 0)
            tot = self.stats.get('total', 0)
            sc = self.stats.get('scanned', 0)
            elapsed = self.stats.get('elapsed_time', 0)
            self.var_stats.set(t['stats_done'].format(total=tot, scanned=sc, matches=m, time=elapsed))
            
        self.frame_preview.config(text=t['lbl_preview'])
        self.btn_prev.config(text=t['btn_prev'])
        self.btn_next.config(text=t['btn_next'])
        
        self.tree.heading("nombre", text=t['col_name'])
        self.tree.heading("ruta", text=t['col_path'])
        self.tree.heading("tipo", text=t['col_type'])
        
        self.cm.entryconfigure(0, label=t['menu_open'])
        self.cm.entryconfigure(1, label=t['menu_open_loc'])
        self.cm.entryconfigure(2, label=t['menu_copy_path'])

    def _apply_theme(self):
        theme_colors = {
            'light': {
                'bg': '#f3f4f6',
                'fg': '#1f2937',
                'field_bg': '#ffffff',
                'border': '#d1d5db',
                'btn_bg': '#e5e7eb',
                'btn_active': '#d1d5db',
                'btn_disabled': '#f3f4f6',
                'fg_disabled': '#9ca3af',
                'accent': '#2563eb',
                'sel_fg': '#ffffff',
                'txt_match_bg': '#a8e6cf',
                'txt_match_fg': 'black',
                'txt_match_active_bg': '#22c55e',
                'txt_match_active_fg': 'white',
                'theme_icon': '🌙'
            },
            'dark': {
                'bg': '#1f2937',
                'fg': '#f9fafb',
                'field_bg': '#111827',
                'border': '#374151',
                'btn_bg': '#374151',
                'btn_active': '#4b5563',
                'btn_disabled': '#1f2937',
                'fg_disabled': '#6b7280',
                'accent': '#3b82f6',
                'sel_fg': '#ffffff',
                'txt_match_bg': '#1e3a8a',
                'txt_match_fg': '#f9fafb',
                'txt_match_active_bg': '#2563eb',
                'txt_match_active_fg': '#ffffff',
                'theme_icon': '☀️'
            }
        }
        
        c = theme_colors[self.theme]
        
        self.root.config(bg=c['bg'])
        self.txt.config(
            bg=c['field_bg'], 
            fg=c['fg'], 
            insertbackground=c['fg'], 
            selectbackground=c['accent'], 
            selectforeground=c['sel_fg']
        )
        self.txt.tag_config("match", background=c['txt_match_bg'], foreground=c['txt_match_fg'])
        self.txt.tag_config("active", background=c['txt_match_active_bg'], foreground=c['txt_match_active_fg'])
        self.cm.config(bg=c['bg'], fg=c['fg'], activebackground=c['accent'], activeforeground=c['sel_fg'])
        
        self.style.configure('.', background=c['bg'], foreground=c['fg'])
        self.style.configure('TFrame', background=c['bg'])
        self.style.configure('TLabel', background=c['bg'], foreground=c['fg'])
        self.style.configure('TLabelframe', background=c['bg'], foreground=c['fg'], bordercolor=c['border'])
        self.style.configure('TLabelframe.Label', background=c['bg'], foreground=c['fg'])
        
        self.style.configure('TButton', background=c['btn_bg'], foreground=c['fg'], bordercolor=c['border'], focuscolor='', lightcolor=c['btn_bg'], darkcolor=c['btn_bg'])
        self.style.map('TButton', 
            background=[('active', c['btn_active']), ('disabled', c['btn_disabled'])], 
            foreground=[('disabled', c['fg_disabled'])],
            bordercolor=[('active', c['border']), ('disabled', c['border'])]
        )
        
        self.style.configure('TRadiobutton', background=c['bg'], foreground=c['fg'], focuscolor='', indicatorbackground=c['field_bg'], indicatorforeground=c['fg'])
        self.style.map('TRadiobutton', 
            background=[('active', c['bg'])], 
            indicatorbackground=[('selected', c['accent']), ('!selected', c['field_bg'])],
            foreground=[('active', c['fg'])]
        )
        self.style.configure('TCheckbutton', background=c['bg'], foreground=c['fg'], focuscolor='', indicatorbackground=c['field_bg'], indicatorforeground=c['fg'])
        self.style.map('TCheckbutton', 
            background=[('active', c['bg'])], 
            indicatorbackground=[('selected', c['accent']), ('!selected', c['field_bg'])],
            foreground=[('active', c['fg'])]
        )
        
        self.style.configure('TEntry', fieldbackground=c['field_bg'], foreground=c['fg'], bordercolor=c['border'], lightcolor=c['field_bg'], darkcolor=c['field_bg'])
        self.style.map('TEntry', bordercolor=[('focus', c['accent']), ('!focus', c['border'])])
        
        self.style.configure('Treeview', background=c['field_bg'], foreground=c['fg'], fieldbackground=c['field_bg'], bordercolor=c['border'])
        self.style.configure('Treeview.Heading', background=c['btn_bg'], foreground=c['fg'], bordercolor=c['border'], lightcolor=c['btn_bg'], darkcolor=c['btn_bg'])
        self.style.map('Treeview.Heading', background=[('active', c['btn_active'])])
        self.style.map('Treeview', background=[('selected', c['accent'])], foreground=[('selected', c['sel_fg'])])
        
        self.style.configure('TCombobox', fieldbackground=c['field_bg'], foreground=c['fg'], background=c['btn_bg'], bordercolor=c['border'], arrowcolor=c['fg'])
        self.style.map('TCombobox', 
            fieldbackground=[('readonly', c['field_bg'])], 
            foreground=[('readonly', c['fg'])],
            arrowcolor=[('active', c['accent'])]
        )
        
        self.style.configure("Status.TLabel", font=("Segoe UI", 10, "bold"), background=c['bg'], foreground=c['fg'])
        self.style.configure("Stats.TLabel", font=("Segoe UI", 9), background=c['bg'], foreground=c['fg'] if self.theme == 'dark' else '#333333')
        
        # Update theme button icon or text
        if self.theme == 'light' and self.theme_photo_light:
            self.btn_theme.config(image=self.theme_photo_light, text="")
        elif self.theme == 'dark' and self.theme_photo_dark:
            self.btn_theme.config(image=self.theme_photo_dark, text="")
        else:
            self.btn_theme.config(image="", text=c['theme_icon'])
            
        self.tree.tag_configure("noresult", foreground="red" if self.theme == 'light' else "#ff6b6b")


if __name__ == '__main__':
    multiprocessing.freeze_support()   # obligatorio para PyInstaller + multiprocessing
    root = tk.Tk()
    ZetaApp(root)
    root.mainloop()
