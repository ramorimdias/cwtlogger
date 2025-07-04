#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gpp_tk_gui.py – GPP-4323 live resistance dashboard
v1.2  (compact freq-selector, visible version label)

• 2×2 buttons (Start/Stop, Check/Stop Check, Save XLSX, Clear Cache)
• Log frequency radio row: 5 s 15 s 30 s 1 min 5 min
• Version tag shown top-left
• Hourly Excel, live change interval, Y-axis 8-15 Ω
"""

__version__ = "v1.2"

# ──── stdlib ─────────────────────────────────────────────────────
import csv, math, time, threading, datetime as dt, sys
from pathlib import Path
from collections import deque
import tkinter as tk
from tkinter import ttk, messagebox

# ──── 3rd-party ──────────────────────────────────────────────────
try:
    import numpy as np
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import pyvisa
    import pyvisa.constants as C
    import xlsxwriter
except ModuleNotFoundError as e:
    print(f"Missing required module: {e.name}", file=sys.stderr)
    sys.exit(1)

# ──── constants ─────────────────────────────────────────────────
PORT, BAUD = "/dev/ttyUSB0", 115200
VSET        = 5.0
WINDOW_H    = 48.0                    # hours shown
EXCEL_INT_H = 1.0                     # h
MAX_POINTS  = 20_000
Y_MIN_DFLT, Y_MAX_DFLT = 8.0, 15.0
FREQ_OPTIONS = [("5 s", 5), ("15 s", 15), ("30 s", 30),
                ("1 min", 60), ("5 min", 300)]
CHAN_LABELS = ["CH1", "CH2", "CH3", "CH4"]
LOG_DIR = Path.home() / "gpp_logs"; LOG_DIR.mkdir(exist_ok=True)
RAW_CSV = LOG_DIR / "raw.csv"

# ──── VISA helpers ───────────────────────────────────────────────
def open_psu():
    rm  = pyvisa.ResourceManager("@py")
    psu = rm.open_resource(f"ASRL{PORT}::INSTR", timeout=3000)
    psu.baud_rate, psu.data_bits = BAUD, 8
    psu.stop_bits, psu.parity    = C.StopBits.one, C.Parity.none
    psu.write_termination, psu.read_termination = "\r\n", "\n"
    psu.write("SYST:REM"); return psu

def chan_on(psu, ch, ilim):
    psu.write(f":SOUR{ch}:VOLT {VSET}")
    psu.write(f":SOUR{ch}:CURR {ilim}")
    psu.write(f":OUTP{ch}:STAT ON")

def chan_off(psu, ch): psu.write(f":OUTP{ch}:STAT OFF")

def safe_R(psu, ch):
    try:
        v,i,_ = map(float, psu.query(f":MEAS{ch}:ALL?").split(","))
        return np.inf if abs(i)<1e-6 else v/i
    except Exception: return np.nan

# ──── CSV / Excel helpers ────────────────────────────────────────
def ensure_raw():
    if RAW_CSV.exists():
        return
    with RAW_CSV.open("w", newline="") as f:
        f.write("#xlsx:\n")
        csv.writer(f).writerow(["time","rel_h",*CHAN_LABELS])

def prompt_existing_csv():
    if not RAW_CSV.exists() or RAW_CSV.stat().st_size == 0:
        return
    with RAW_CSV.open() as f:
        lines = (ln for ln in f if not ln.startswith("#"))
        next(lines, None)  # skip header if present
        if not any(ln.strip() for ln in lines):
            return
    if messagebox.askyesno(
            "Existing Log",
            "raw.csv contains previous data. Continue using it?\n"
            "Choose 'No' to start fresh."):
        return
    if messagebox.askyesno(
            "Confirm Delete",
            "Delete existing raw.csv and start with a blank log?"):
        RAW_CSV.unlink()

def current_xlsx():
    with RAW_CSV.open() as f:
        first=f.readline().strip()
    return Path(first[6:]) if first.startswith("#xlsx:") and first[6:] else None

def set_xlsx(path: Path):
    tmp = RAW_CSV.with_suffix(".tmp")
    with RAW_CSV.open() as fin, tmp.open("w") as fout:
        fout.write(f"#xlsx:{path}\n"); next(fin)
        for ln in fin: fout.write(ln)
    tmp.replace(RAW_CSV)

def csv_to_xlsx(csv_path: Path, xlsx_path: Path):
    with csv_path.open() as f:
        rows=[r.rstrip("\n").split(",") for r in f if not r.startswith("#")]
    wb=xlsxwriter.Workbook(xlsx_path)
    ws_d,ws_c=wb.add_worksheet("Data"),wb.add_worksheet("Chart")
    chart=wb.add_chart({"type":"line"})
    for r,row in enumerate(rows):
        for c,val in enumerate(row):
            if r==0: ws_d.write(r,c,val)
            elif val and (c==0 or math.isfinite(float(val))):
                ws_d.write(r,c,val if c==0 else float(val))
    max_row=len(rows)-1
    for col,label in enumerate(CHAN_LABELS,2):
        col_letter=xlsxwriter.utility.xl_col_to_name(col)
        chart.add_series({
            "name":label,
            "categories":f"=Data!$B$2:$B${max_row+1}",
            "values":   f"=Data!${col_letter}$2:${col_letter}${max_row+1}",
        })
    chart.set_x_axis({"name":"Hours from start"})
    chart.set_y_axis({"name":"R (Ω)"})
    ws_c.insert_chart("B2", chart, {"x_scale":1.5,"y_scale":1.3})
    wb.close()

# ──── GUI class ────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        ttk.Style(self).configure(".", font=("Helvetica",14))
        self.title("GPP-4323 · Resistance monitor")

        # ► Start fullscreen
        self.full=True
        self.attributes("-fullscreen", True)

        self.columnconfigure(0,weight=1); self.rowconfigure(3,weight=1)

        # top-bar (version + toggle)
        top=ttk.Frame(self); top.grid(row=0,column=0,sticky="ew")
        ttk.Label(top,text=__version__,font=("Helvetica",10,"italic"))\
            .pack(side="left",padx=6)
        ttk.Button(top,text="Full / Window",
                   command=self._toggle_full).pack(side="right",padx=6,pady=2)

        # control frame
        ctrl=ttk.LabelFrame(self,text="Controls")
        ctrl.grid(row=1,column=0,sticky="ew",padx=8,pady=(0,4))
        ctrl.columnconfigure(9,weight=2)

        self.chk=[tk.BooleanVar() for _ in CHAN_LABELS]
        for i,(v,label) in enumerate(zip(self.chk, CHAN_LABELS),1):
            ttk.Checkbutton(ctrl,text=label,variable=v)\
               .grid(row=0,column=i-1,padx=8,sticky="w")

        def spin(r,label,u,init,step,name):
            ttk.Label(ctrl,text=label).grid(row=r,column=0,sticky="e")
            setattr(self,name,tk.DoubleVar(value=init))
            ttk.Entry(ctrl,width=8,textvariable=getattr(self,name))\
               .grid(row=r,column=1,sticky="w")
            for txt,d,col in ((" – ",-step,2),(" + ",step,3)):
                ttk.Button(ctrl,text=txt,width=3,
                           command=lambda n=name,dd=d:self._bump(n,dd))\
                           .grid(row=r,column=col)
            ttk.Label(ctrl,text=u).grid(row=r,column=4,sticky="w")
        spin(1,"Max current:","A",0.100,0.010,"i_var")
        spin(2,"Y-min:","Ω",Y_MIN_DFLT,1.0,"ymin_var")
        spin(3,"Y-max:","Ω",Y_MAX_DFLT,1.0,"ymax_var")
        ttk.Button(ctrl,text="Set Y-axis",command=self.apply_y)\
            .grid(row=2,column=5,rowspan=2,padx=(10,0),pady=8)

        # freq selector
        ttk.Label(ctrl,text="Log every:").grid(row=0,column=5,sticky="e")
        self.sample_int=tk.IntVar(value=5)
        freq_frame=ttk.Frame(ctrl); freq_frame.grid(row=0,column=6,columnspan=3,sticky="w")
        for lbl,val in FREQ_OPTIONS:
            ttk.Radiobutton(freq_frame,text=lbl,variable=self.sample_int,
                            value=val).pack(side="left",padx=2)

        # buttons (2×2)
        btns=ttk.Frame(ctrl); btns.grid(row=1,column=6,columnspan=4,rowspan=4,padx=(20,0))
        self.start_btn=ttk.Button(btns,text="Start",width=16,command=self.start_log)
        self.check_btn=ttk.Button(btns,text="Check",width=16,command=self.check_toggle)
        self.save_btn =ttk.Button(btns,text="Save XLSX",width=16,command=self.save_xlsx)
        self.clear_btn=ttk.Button(btns,text="Clear Cache",width=16,command=self.clear_cache)
        for w,r,c in ((self.start_btn,0,0),(self.check_btn,0,1),
                      (self.save_btn ,1,0),(self.clear_btn ,1,1)):
            w.grid(row=r,column=c,padx=12,pady=8,sticky="ew")

        # plot
        fig,self.ax=plt.subplots(figsize=(8,5))
        self.ax.set_xlabel("time"); self.ax.set_ylabel("R (Ω)")
        self.ax.set_ylim(Y_MIN_DFLT,Y_MAX_DFLT)
        self.locator = mdates.AutoDateLocator(minticks=5, maxticks=10)
        self.locator.intervald.setdefault(mdates.MINUTELY,[1,2,5,10,15,30])
        self.ax.xaxis.set_major_locator(self.locator)
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b\n%H:%M"))
        fig.subplots_adjust(left=0.05,right=0.75,top=0.97,bottom=0.3)
        self.lines=[self.ax.plot([],[],label=lbl)[0] for lbl in CHAN_LABELS]
        self.leg=self.ax.legend(loc="center left",
                                bbox_to_anchor=(1,0.5),fontsize=16)
        self.canvas=FigureCanvasTkAgg(fig,master=self)
        self.canvas.get_tk_widget().grid(row=2,column=0,sticky="nsew",
                                         padx=8,pady=4)

        # runtime
        prompt_existing_csv()
        ensure_raw()
        self.t,self.r=self._load_cache()
        self.psu=None; self.mode=None
        self.thread=None; self.stop_evt=threading.Event()
        self.after_id=None; self.next_xlsx=None; self.t0=None
        self.redraw()

    # helpers
    def _toggle_full(self):
        self.full=not self.full; self.attributes("-fullscreen",self.full)

    def _bump(self,name,d):
        try:v=getattr(self,name).get()+d; getattr(self,name).set(round(v,3))
        except tk.TclError: pass

    def apply_y(self):
        ymin,ymax=self.ymin_var.get(),self.ymax_var.get()
        if ymax<=ymin: messagebox.showerror("Input","Y-max must be > Y-min"); return
        self.ax.set_ylim(ymin,ymax); self.canvas.draw_idle()

    def _load_cache(self):
        t=deque(maxlen=MAX_POINTS)
        r=[deque(maxlen=MAX_POINTS) for _ in CHAN_LABELS]
        with RAW_CSV.open() as f:
            rdr=csv.reader(l for l in f if not l.startswith("#"))
            next(rdr,None)
            for row in rdr:
                t.append(mdates.datestr2num(row[0]))
                for i,val in enumerate(row[2:2+len(CHAN_LABELS)]):
                    r[i].append(float(val) if val else np.nan)
        return t,r

    # worker -----------------------------------------------------------------
    def worker(self,chans,ilim):
        excel=current_xlsx()
        if excel is None:
            excel=LOG_DIR/f"gpp_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            set_xlsx(excel)
        self.next_xlsx=dt.datetime.now()+dt.timedelta(hours=EXCEL_INT_H)
        with RAW_CSV.open("a",newline="") as f:
            w=csv.writer(f)
            while not self.stop_evt.is_set():
                now=dt.datetime.now(); rel=(now-self.t0).total_seconds()/3600
                self.t.append(mdates.date2num(now))
                for idx,ch in enumerate(range(1,len(CHAN_LABELS)+1)):
                    self.r[idx].append(safe_R(self.psu,ch)
                                       if ch in chans else np.nan)
                w.writerow([now.strftime("%Y-%m-%d %H:%M:%S"),
                            f"{rel:.4f}"]+[ "" if np.isnan(self.r[i][-1])
                                            else f"{self.r[i][-1]:.4f}"
                                            for i in range(len(CHAN_LABELS))])
                f.flush()
                if now>=self.next_xlsx:
                    csv_to_xlsx(RAW_CSV,excel)
                    self.next_xlsx+=dt.timedelta(hours=EXCEL_INT_H)
                time.sleep(self.sample_int.get())          # live interval

    # redraw ---------------------------------------------------------------
    def redraw(self):
        if self.t:
            now=mdates.date2num(dt.datetime.now()); cut=now-WINDOW_H/24
            i0=next((i for i,tt in enumerate(self.t) if tt>=cut),len(self.t)-1)
            x=np.array(list(self.t)[i0:])
            live=[]
            for idx,ln in enumerate(self.lines):
                y=np.array(list(self.r[idx])[i0:],float); y[np.isinf(y)]=np.nan
                ln.set_data(x,y)
                v=y[-1] if len(y) else np.nan
                label=CHAN_LABELS[idx]
                live.append(f"{label}: ---" if np.isnan(v)
                            else f"{label}: {v:.3f} Ω")
            for txt,new in zip(self.leg.get_texts(),live): txt.set_text(new)
            if len(x)>1: self.ax.set_xlim(x[0],x[-1])
            self.canvas.draw_idle()
        if self.mode: self.after_id=self.after(1000,self.redraw)

    # start / stop --------------------------------------------------------
    def _start(self,mode):
        chans=[i+1 for i,v in enumerate(self.chk) if v.get()]
        if not chans:
            messagebox.showwarning("Select","Tick at least one channel"); return False
        ilim=self.i_var.get()
        try:self.psu=open_psu()
        except Exception as e: messagebox.showerror("PSU",e); return False
        for ch in chans: chan_on(self.psu,ch,ilim)
        self.t0=dt.datetime.now(); self.mode=mode; self.stop_evt.clear()
        self.thread=threading.Thread(target=self.worker,args=(chans,ilim),daemon=True)
        self.thread.start()
        if not self.after_id:
            self.after_id=self.after(1000,self.redraw)
        return True

    def start_log(self):
        if self.mode: messagebox.showinfo("Busy","Stop current run first"); return
        if self._start("log"):
            self.start_btn.config(text="Stop",command=self.stop_run)
            self.check_btn.state(["disabled"])

    def check_toggle(self):
        if self.mode is None:
            if self._start("check"):
                self.check_btn.config(text="Stop Check")
                self.start_btn.state(["disabled"])
        else:
            self.stop_run()

    def stop_run(self):
        if not self.mode: return
        if not messagebox.askyesno("Stop","Stop current operation?"): return
        self.stop_evt.set(); self.thread.join()
        try:
            for ch in range(1, len(CHAN_LABELS)+1):
                chan_off(self.psu,ch)
            self.psu.close()
        except: pass
        if self.after_id:
            self.after_cancel(self.after_id); self.after_id=None
        self.mode=None
        self.start_btn.config(text="Start",command=self.start_log)
        self.check_btn.config(text="Check")
        self.start_btn.state(["!disabled"]); self.check_btn.state(["!disabled"])

    # misc buttons --------------------------------------------------------
    def save_xlsx(self):
        excel=current_xlsx()
        if excel is None:
            excel=LOG_DIR/f"gpp_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            set_xlsx(excel)
        csv_to_xlsx(RAW_CSV,excel)

    def clear_cache(self):
        if not messagebox.askyesno("Clear","Delete raw cache?"): return
        RAW_CSV.unlink(missing_ok=True); ensure_raw()
        self.t.clear(); [q.clear() for q in self.r]; self.canvas.draw_idle()

    def quit_safe(self):
        if self.mode and not messagebox.askyesno("Quit","Stop current run and quit?"):
            return
        if self.mode: self.stop_run()
        self.destroy()

# ─ run ─
if __name__=="__main__":
    app=App()
    app.protocol("WM_DELETE_WINDOW",app.quit_safe)
    app.mainloop()
