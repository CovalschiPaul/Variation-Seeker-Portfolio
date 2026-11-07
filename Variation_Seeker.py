"""Offline portfolio demonstration. All bundled examples are fictional."""
import argparse
from datetime import datetime
import logging
from pathlib import Path
import queue
import sys
import threading
from variation_core import DataError, Result, export_result, load_table

ROOT=Path(__file__).resolve().parent
LOG=logging.getLogger('variation_seeker')

def run(csi,rtip,output,overwrite=False):
    if any(Path(output).resolve()==Path(p).resolve() for p in (csi,rtip)):
        raise DataError('Output must be different from both input files.')
    result=Result(csi=load_table(csi,'CSI'),rtip=load_table(rtip,'RTIP'),metadata={'Mode':'Offline comparison; bundled sample data is synthetic'})
    path=export_result(result,output,overwrite)
    LOG.info('Saved %s; CSI rows=%s; RTIP rows=%s',path,len(result.csi),len(result.rtip))
    return path


def gui():
    import tkinter as tk
    from tkinter import ttk,filedialog,messagebox
    root=tk.Tk();root.title('Variation Seeker — Portfolio Demo');root.geometry('820x440');root.minsize(640,400)
    panel=ttk.Frame(root,padding=24);panel.pack(fill='both',expand=True);panel.columnconfigure(1,weight=1)
    ttk.Label(panel,text='Variation Seeker',font=('Segoe UI',24,'bold')).grid(row=0,column=0,columnspan=3,sticky='w')
    ttk.Label(panel,text='Catalog relationship reconciliation · Offline portfolio demonstration').grid(row=1,column=0,columnspan=3,sticky='w',pady=(8,20))
    fields=[]
    for row,label,default in [(2,'CSI input',ROOT/'samples'/'csi.csv'),(3,'RTIP input',ROOT/'samples'/'rtip.csv'),(4,'Excel output',Path.cwd()/f'Variation_Demo_{datetime.now():%Y%m%d_%H%M%S}.xlsx')]:
        variable=tk.StringVar(value=str(default));fields.append(variable)
        ttk.Label(panel,text=label).grid(row=row,column=0,sticky='w',pady=8)
        ttk.Entry(panel,textvariable=variable).grid(row=row,column=1,sticky='ew',padx=12)
        def browse(v=variable,save=row==4):
            path=filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')]) if save else filedialog.askopenfilename(filetypes=[('Data','*.csv *.xlsx *.html *.htm')])
            if path:v.set(path)
        ttk.Button(panel,text='Browse',command=browse).grid(row=row,column=2)
    ttk.Label(panel,text='The included files contain fictional products, sellers and identifiers.\nNo account, browser automation or network connection is used.').grid(row=5,column=0,columnspan=3,sticky='w',pady=18)
    status=tk.StringVar(value='Ready — sample files are selected.')
    ttk.Label(panel,textvariable=status,wraplength=730).grid(row=7,column=0,columnspan=3,sticky='w',pady=16)
    events=queue.Queue();worker=None;closing=False
    def start():
        nonlocal worker
        paths=[x.get().strip() for x in fields]
        if not all(paths):messagebox.showerror('Missing path','Choose both input files and an output path.');return
        overwrite=Path(paths[2]).exists()
        if overwrite and not messagebox.askyesno('Replace workbook?',f'Replace {paths[2]}?'):return
        button.configure(state='disabled');status.set('Comparing relationships…')
        def work():
            try:events.put(('ok',str(run(*paths,overwrite))))
            except Exception as exc:events.put(('error',str(exc)))
        worker=threading.Thread(target=work,daemon=False);worker.start()
    button=ttk.Button(panel,text='Compare and export Excel',command=start);button.grid(row=6,column=0,columnspan=3,sticky='w')
    def poll():
        while not events.empty():
            kind,message=events.get_nowait();button.configure(state='normal')
            status.set(('Saved: ' if kind=='ok' else 'Failed: ')+message)
            if kind=='error' and not closing:messagebox.showerror('Comparison failed',message)
        if closing and (worker is None or not worker.is_alive()):root.destroy();return
        root.after(100,poll)
    def close():
        nonlocal closing
        closing=True;status.set('Finishing the current export before closing…')
    root.protocol('WM_DELETE_WINDOW',close);poll();root.mainloop()


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--demo',action='store_true')
    parser.add_argument('--csi',type=Path)
    parser.add_argument('--rtip',type=Path)
    parser.add_argument('--output',type=Path,default=Path('Variation_Seeker_output.xlsx'))
    parser.add_argument('--overwrite',action='store_true')
    parser.add_argument('--log',type=Path)
    args=parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s')
    try:
        if args.log:
            args.log.parent.mkdir(parents=True,exist_ok=True)
            LOG.addHandler(logging.FileHandler(args.log,encoding='utf-8'))
        if args.demo and (args.csi or args.rtip):raise DataError('Choose --demo or --csi and --rtip.')
        if args.demo:run(ROOT/'samples'/'csi.csv',ROOT/'samples'/'rtip.csv',args.output,args.overwrite)
        elif args.csi or args.rtip:
            if not args.csi or not args.rtip:raise DataError('Provide both --csi and --rtip.')
            run(args.csi,args.rtip,args.output,args.overwrite)
        else:gui()
        return 0
    except Exception as exc:
        LOG.error('%s',exc);return 1

if __name__=='__main__':sys.exit(main())
