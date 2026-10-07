"""pdf2md_gui.py — window for the PDF / Word -> Markdown converter (tkinter, part of standard Python).
  py pdf2md_gui.py [files or folders to add]"""
import os, sys, json, threading, queue, traceback, subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf2md as P

CFG = os.path.join(P.HERE, "pdf2md_gui.cfg")
def load_cfg():
    try: return json.load(open(CFG, encoding="utf-8"))
    except Exception: return {}

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(f"pdf2md {P.__version__} — PDF / Word to Markdown"); self.geometry("840x620"); self.minsize(720, 520)
        cfg = load_cfg(); self.files = []; self.q = queue.Queue(); self.busy = False
        f = ttk.Frame(self, padding=10); f.pack(fill="both", expand=True)
        ttk.Label(f, text="Files (PDF, DOCX):").pack(anchor="w")
        self.lb = tk.Listbox(f, height=8, selectmode="extended"); self.lb.pack(fill="x")
        b = ttk.Frame(f); b.pack(fill="x", pady=4)
        ttk.Button(b, text="Add files…", command=self.add_files).pack(side="left")
        ttk.Button(b, text="Add folder…", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(b, text="Remove selected", command=self.del_sel).pack(side="left")
        ttk.Button(b, text="Clear", command=self.clear).pack(side="left", padx=6)
        o = ttk.LabelFrame(f, text="Options", padding=8); o.pack(fill="x", pady=(8, 0))
        B = lambda k, d: tk.BooleanVar(value=cfg.get(k, d))
        self.v_img, self.v_hf, self.v_sep, self.v_toc, self.v_num = B("images", True), B("strip_hf", True), B("page_sep", False), B("strip_toc", True), B("numbering", True)
        ttk.Checkbutton(o, text="Save pictures (folder <name>_images)", variable=self.v_img).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(o, text="Remove page headers / footers", variable=self.v_hf).grid(row=0, column=1, sticky="w", padx=20)
        ttk.Checkbutton(o, text="Mark page breaks with --- (PDF)", variable=self.v_sep).grid(row=1, column=0, sticky="w")
        ttk.Label(o, text="Picture quality (PDF):").grid(row=1, column=1, sticky="w", padx=20)
        self.v_dpi = ttk.Combobox(o, values=["100", "150", "200", "300"], width=6, state="readonly"); self.v_dpi.set(cfg.get("dpi", "150")); self.v_dpi.grid(row=1, column=2, sticky="w")
        ttk.Checkbutton(o, text="Remove table of contents", variable=self.v_toc).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(o, text="Restore Word heading numbers (1, 1.1 …)", variable=self.v_num).grid(row=2, column=1, columnspan=2, sticky="w", padx=20)
        r = ttk.Frame(f); r.pack(fill="x", pady=(8, 0))
        ttk.Label(r, text="Save to:").pack(side="left"); self.v_out = tk.StringVar(value=cfg.get("out", P.OUT))
        ttk.Entry(r, textvariable=self.v_out).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(r, text="Change…", command=self.pick_out).pack(side="left")
        a = ttk.Frame(f); a.pack(fill="x", pady=8)
        self.run_btn = ttk.Button(a, text="▶ Convert", command=self.run); self.run_btn.pack(side="left")
        ttk.Button(a, text="Open output folder", command=self.open_out).pack(side="left", padx=10)
        self.pb = ttk.Progressbar(a, mode="determinate"); self.pb.pack(side="left", fill="x", expand=True)
        ttk.Label(f, text="Log:").pack(anchor="w"); self.log = tk.Text(f, height=10, state="disabled"); self.log.pack(fill="both", expand=True)
        for p in P.collect(sys.argv[1:]): self._add(p)          # files / folders dragged onto run_gui.bat
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(100, self.pump)

    # the worker thread talks to the window through a queue (tkinter is not thread-safe)
    def logw(self, s): self.q.put(("log", s))
    def pump(self):
        while not self.q.empty():
            k, v = self.q.get()
            if k == "log": self.log.configure(state="normal"); self.log.insert("end", v + "\n"); self.log.see("end"); self.log.configure(state="disabled")
            elif k == "prog": self.pb["value"] = v
            elif k == "done": self.busy = False; self.run_btn.state(["!disabled"])
        self.after(100, self.pump)

    def _add(self, p):
        p = os.path.abspath(p)
        if p.lower().endswith(P.SUPPORTED) and not os.path.basename(p).startswith("~$") and p not in self.files: self.files.append(p); self.lb.insert("end", p)
    def add_files(self):
        for p in filedialog.askopenfilenames(filetypes=[("PDF and Word", "*.pdf *.docx"), ("PDF", "*.pdf"), ("Word", "*.docx")], initialdir=P.IN): self._add(p)
    def add_folder(self):
        d = filedialog.askdirectory(initialdir=P.IN)
        if d:
            for p in P.collect([d]): self._add(p)
    def del_sel(self):
        for i in reversed(self.lb.curselection()): self.lb.delete(i); del self.files[i]
    def clear(self): self.lb.delete(0, "end"); self.files.clear()
    def pick_out(self):
        d = filedialog.askdirectory(initialdir=self.v_out.get())
        if d: self.v_out.set(d)
    def open_out(self):
        d = self.v_out.get(); os.makedirs(d, exist_ok=True)
        if hasattr(os, "startfile"): os.startfile(d)
        else: subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", d])

    def run(self):
        if self.busy: return
        if not self.files: messagebox.showinfo("No files", "Add at least one PDF or DOCX file."); return
        self.busy = True; self.run_btn.state(["disabled"]); self.pb["maximum"] = len(self.files); self.pb["value"] = 0
        opts = dict(out_dir=self.v_out.get(), images=self.v_img.get(), strip_hf=self.v_hf.get(), strip_toc=self.v_toc.get(),
                    number_headings=self.v_num.get(), page_sep=self.v_sep.get(), dpi=int(self.v_dpi.get()))
        threading.Thread(target=self._run, args=(list(self.files), opts), daemon=True).start()

    def _run(self, files, opts):
        ok = 0; names = P.output_names(files)
        for i, f in enumerate(files, 1):
            self.logw(f"→ {os.path.basename(f)}")
            try:
                out = P.convert_any(f, log=self.logw, name=names[f], **opts)
                self.logw(f"OK  {os.path.basename(out)}  ({max(1, os.path.getsize(out) // 1024)} KB)"); ok += 1
            except Exception as e:
                self.logw(f"ERROR {os.path.basename(f)}: {e}"); traceback.print_exc()
            self.q.put(("prog", i))
        self.logw(f"Done: {ok} of {len(files)} — see {opts['out_dir']}"); self.q.put(("done", None))

    def close(self):
        cfg = {"out": self.v_out.get(), "images": self.v_img.get(), "strip_hf": self.v_hf.get(), "page_sep": self.v_sep.get(),
               "strip_toc": self.v_toc.get(), "numbering": self.v_num.get(), "dpi": self.v_dpi.get()}
        try: json.dump(cfg, open(CFG, "w", encoding="utf-8"), indent=1)
        except Exception: pass
        self.destroy()

if __name__ == "__main__":
    if "--version" in sys.argv: sys.exit(f"pdf2md {P.__version__}")
    App().mainloop()
