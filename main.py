import os, re, sys, threading, queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP = "Convertidor de Actas Viales PRO"

def metadata(text):
    text = re.sub(r"\s+", " ", text or "")
    upper = text.upper()
    domain = None
    for pat in (r"\b([A-Z]{2}\s?\d{3}\s?[A-Z]{2})\b", r"\b([A-Z]{3}\s?\d{3})\b"):
        m = re.search(pat, upper)
        if m:
            domain = re.sub(r"\s+", "", m.group(1))
            break
    acta = None
    pats = (
        r"(?:ACTA\s*(?:VIAL)?|N[°ºO]\s*DE?\s*ACTA|NUMERO\s*DE?\s*ACTA)\s*(?:N[°ºO]\s*)?[:\-]?\s*([0-9][0-9./-]{2,})",
        r"\bACTA\s+VIAL\s*[:\-]?\s*([0-9][0-9./-]{2,})",
    )
    for pat in pats:
        m = re.search(pat, text, re.I)
        if m:
            acta = m.group(1).strip(" ./-")
            break
    return domain, acta

class Word:
    def __init__(self):
        import pythoncom, win32com.client
        self.com = pythoncom
        self.client = win32com.client
    def _app(self):
        w = self.client.DispatchEx("Word.Application")
        w.Visible = False
        w.DisplayAlerts = 0
        return w
    def read(self, path):
        self.com.CoInitialize()
        w = d = None
        try:
            w = self._app()
            d = w.Documents.Open(os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, Visible=False)
            return d.Content.Text
        finally:
            if d is not None: d.Close(False)
            if w is not None: w.Quit()
            self.com.CoUninitialize()
    def pdf(self, source, target):
        self.com.CoInitialize()
        w = d = None
        try:
            w = self._app()
            d = w.Documents.Open(os.path.abspath(source), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, Visible=False)
            d.SaveAs2(os.path.abspath(target), FileFormat=17)
        finally:
            if d is not None: d.Close(False)
            if w is not None: w.Quit()
            self.com.CoUninitialize()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP); self.geometry("1120x700"); self.minsize(900,560)
        self.configure(bg="#f3f5f8")
        self.consultas=[]; self.targets=[]; self.events=queue.Queue()
        self.output=tk.StringVar(); self.status=tk.StringVar(value="Seleccioná los documentos para comenzar.")
        self._ui(); self.after(150,self._poll)
    def _ui(self):
        s=ttk.Style(self)
        try:s.theme_use("clam")
        except tk.TclError:pass
        s.configure("Treeview",rowheight=30,font=("Segoe UI",9))
        s.configure("Treeview.Heading",font=("Segoe UI",9,"bold"))
        s.configure("Title.TLabel",font=("Segoe UI",19,"bold"),background="#f3f5f8",foreground="#17243a")
        s.configure("Sub.TLabel",background="#f3f5f8",foreground="#536176")
        f=ttk.Frame(self,padding=18); f.pack(fill="both",expand=True)
        ttk.Label(f,text=APP,style="Title.TLabel").pack(anchor="w")
        ttk.Label(f,text="Cargá los Word de consulta y los Word que se convertirán. Revisá cada asociación antes de procesar.",style="Sub.TLabel").pack(anchor="w",pady=(4,14))
        bar=ttk.Frame(f);bar.pack(fill="x",pady=(0,10))
        ttk.Button(bar,text="1. Agregar Word de consulta",command=self.add_queries).pack(side="left")
        ttk.Button(bar,text="2. Agregar Word a convertir",command=self.add_targets).pack(side="left",padx=8)
        ttk.Button(bar,text="Quitar seleccionado",command=self.remove).pack(side="left")
        cols=("query","domain","act","target","status")
        self.tree=ttk.Treeview(f,columns=cols,show="headings")
        titles=("ARCHIVO DE CONSULTA","DOMINIO","ACTA VIAL","DOCUMENTO A CONVERTIR","ESTADO")
        widths=(210,120,110,340,120)
        for c,t,w in zip(cols,titles,widths):
            self.tree.heading(c,text=t);self.tree.column(c,width=w,anchor="w")
        self.tree.pack(fill="both",expand=True)
        self.tree.bind("<Double-1>",self.edit)
        ttk.Label(f,text="Doble clic en una fila para cambiar el documento de consulta asociado. Los originales no se modifican.",style="Sub.TLabel").pack(anchor="w",pady=9)
        out=ttk.Frame(f);out.pack(fill="x",pady=5)
        ttk.Label(out,text="Carpeta de salida:").pack(side="left")
        ttk.Entry(out,textvariable=self.output).pack(side="left",fill="x",expand=True,padx=8)
        ttk.Button(out,text="Examinar...",command=self.pick_output).pack(side="left")
        self.progress=ttk.Progressbar(f,mode="determinate");self.progress.pack(fill="x",pady=(8,5))
        bottom=ttk.Frame(f);bottom.pack(fill="x")
        ttk.Label(bottom,textvariable=self.status).pack(side="left")
        self.run=ttk.Button(bottom,text="CONVERTIR Y RENOMBRAR PDF",command=self.start);self.run.pack(side="right")
    def add_queries(self):
        fs=filedialog.askopenfilenames(title="Word de consulta: contienen dominio y acta",filetypes=[("Word","*.doc *.docx")])
        if fs:
            self.status.set("Leyendo documentos con Microsoft Word...")
            threading.Thread(target=self.read_queries,args=(list(fs),),daemon=True).start()
    def read_queries(self,fs):
        try:
            w=Word(); rows=[]
            for p in fs:
                d,a=metadata(w.read(p));rows.append((p,d,a))
            self.events.put(("queries",rows))
        except Exception as e:self.events.put(("error",str(e)))
    def add_targets(self):
        fs=filedialog.askopenfilenames(title="Word que se convertirán a PDF",filetypes=[("Word","*.doc *.docx")])
        existing={x["path"] for x in self.targets}
        for p in fs:
            if p not in existing:self.targets.append({"path":p,"query":None})
        self.defaults();self.refresh()
    def defaults(self):
        q=[x[0] for x in self.consultas]
        for i,x in enumerate(self.targets):
            if x["query"] not in q:x["query"]=q[min(i,len(q)-1)] if q else None
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        md={x[0]:x[1:] for x in self.consultas}
        for i,x in enumerate(self.targets):
            q=x["query"];d,a=md.get(q,(None,None))
            self.tree.insert("","end",iid=str(i),values=(Path(q).name if q else "Sin asignar",d or "NO DETECTADO",a or "NO DETECTADA",Path(x["path"]).name,"Listo" if d and a else "Revisar datos"))
    def edit(self,event=None):
        sel=self.tree.selection()
        if not sel or not self.consultas:return
        i=int(sel[0]);opts=[x[0] for x in self.consultas]
        win=tk.Toplevel(self);win.title("Asociar documento");win.transient(self);win.grab_set()
        ttk.Label(win,text="Seleccioná el documento que contiene dominio y acta:").pack(padx=14,pady=10)
        var=tk.StringVar(value=self.targets[i]["query"] or opts[0])
        cb=ttk.Combobox(win,values=opts,textvariable=var,state="readonly",width=85);cb.pack(padx=14,pady=5)
        def save():
            self.targets[i]["query"]=var.get();self.refresh();win.destroy()
        ttk.Button(win,text="Guardar asociación",command=save).pack(pady=12)
    def remove(self):
        sel=self.tree.selection()
        if sel:del self.targets[int(sel[0])];self.refresh()
    def pick_output(self):
        p=filedialog.askdirectory(title="Carpeta donde guardar los PDF")
        if p:self.output.set(p)
    def start(self):
        if not self.consultas or not self.targets:
            messagebox.showwarning(APP,"Agregá documentos de consulta y documentos a convertir.");return
        if not self.output.get():self.pick_output()
        if not self.output.get():return
        md={x[0]:x[1:] for x in self.consultas}
        bad=[Path(x["path"]).name for x in self.targets if not all(md.get(x["query"],(None,None)))]
        if bad:
            messagebox.showerror(APP,"Faltan dominio o número de acta en estas asociaciones. No se procesarán:\n\n"+"\n".join(bad));return
        self.run.configure(state="disabled");self.progress.configure(maximum=len(self.targets),value=0)
        self.status.set("Convirtiendo con Microsoft Word...")
        threading.Thread(target=self.convert,daemon=True).start()
    def convert(self):
        md={x[0]:x[1:] for x in self.consultas};out=Path(self.output.get());ok=[];err=[]
        try:
            w=Word()
            for i,x in enumerate(self.targets,1):
                d,a=md[x["query"]];stem=f"{d} - ACTA VIAL {a}";dest=out/(stem+".pdf");n=2
                while dest.exists():dest=out/(f"{stem} ({n}).pdf");n+=1
                try:w.pdf(x["path"],str(dest));ok.append(dest.name)
                except Exception as e:err.append(f"{Path(x['path']).name}: {e}")
                self.events.put(("progress",i))
        except Exception as e:err.append(str(e))
        self.events.put(("done",ok,err))
    def _poll(self):
        try:
            while True:
                e=self.events.get_nowait()
                if e[0]=="queries":
                    old={x[0] for x in self.consultas};self.consultas.extend(x for x in e[1] if x[0] not in old)
                    self.defaults();self.refresh();self.status.set(f"{len(self.consultas)} documento(s) de consulta leídos.")
                    missing=[Path(p).name for p,d,a in e[1] if not d or not a]
                    if missing:messagebox.showwarning(APP,"No se detectaron ambos datos en:\n"+"\n".join(missing))
                elif e[0]=="error":messagebox.showerror(APP,e[1])
                elif e[0]=="progress":self.progress.configure(value=e[1])
                elif e[0]=="done":
                    ok,err=e[1],e[2];self.run.configure(state="normal")
                    self.status.set(f"Terminado: {len(ok)} PDF generado(s), {len(err)} error(es).")
                    msg=f"PDF generados: {len(ok)}"
                    if ok:msg+="\n\n"+"\n".join(ok[:10])
                    if err:msg+="\n\nErrores:\n"+"\n".join(err[:8])
                    messagebox.showinfo(APP,msg)
        except queue.Empty:pass
        self.after(150,self._poll)

if __name__=="__main__":
    try:App().mainloop()
    except Exception as e:messagebox.showerror(APP,str(e))
