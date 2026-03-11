# -*- coding: utf-8 -*-
"""
DWG文件语义预处理工具（批量去标注）
功能：批量处理DWG文件，移除标注、辅助线、中心线、孤立短线等非核心几何元素，输出清理后的DXF文件
依赖：ezdxf（解析DXF）、tkinter（GUI界面）、subprocess（调用ODA转换器）、odaFileConverter（DWG转DXF工具）
"""
import os, math, threading, subprocess, shutil, tempfile  # 新增 tempfile 导入
from pathlib import Path               
import tkinter as tk                   
from tkinter import filedialog, messagebox  
import ezdxf                           

# 配置常量 
ANN_TYPES = {"TEXT","MTEXT","DIMENSION","LEADER","MLEADER","TOLERANCE"}
EXTRA_TYPES = {"XLINE","RAY","POINT","SHAPE","ACAD_PROXY_ENTITY"}  
CEN_KEYS = ["CEN","CENTER","CENTRE","CL","CTR","AXIS","DATUM"]     
DIM_KEYS = ["DIM","DIMS","ANNOT","NOTE","TEXT","TAG"]
AUX_KEYS = ["AUX","HELP","GUIDE","CONSTRUCT","CONST","TEMP","PHANTOM","HIDDEN"]
KEYS = CEN_KEYS + DIM_KEYS + AUX_KEYS
GEOM = {"LINE","LWPOLYLINE","POLYLINE","ARC","CIRCLE","SPLINE","HATCH"}

# 工具函数
def kw_hit(s: str) -> bool:
    s = (s or "").upper()
    return any(k in s for k in KEYS)

def line_len(e) -> float:
    a,b = e.dxf.start, e.dxf.end
    return math.hypot(b.x-a.x, b.y-a.y)

def estimate_scale(doc) -> float:
    msp = doc.modelspace()
    xs, ys = [], []
    for e in msp:
        t = e.dxftype()
        try:
            if t == "LINE":
                a,b = e.dxf.start, e.dxf.end
                xs += [a.x,b.x]; ys += [a.y,b.y]
            elif t == "CIRCLE":
                c,r = e.dxf.center, float(e.dxf.radius)
                xs += [c.x-r,c.x+r]; ys += [c.y-r,c.y+r]
            elif t == "ARC":
                c,r = e.dxf.center, float(e.dxf.radius)
                xs += [c.x-r,c.x+r]; ys += [c.y-r,c.y+r]
            elif t == "LWPOLYLINE":
                for x,y,*_ in e.get_points():
                    xs.append(x); ys.append(y)
            elif t == "POLYLINE":
                for v in e.vertices:
                    p = v.dxf.location; xs.append(p.x); ys.append(p.y)
            elif t == "HATCH":
                for loop in e.loops:
                    for vert in loop.vertices:
                        xs.append(vert.x)
                        ys.append(vert.y)
        except:
            pass
    if not xs:
        return 100.0
    return max(max(xs)-min(xs), max(ys)-min(ys), 1.0)

def run_oda_single(oda_exe: Path, src_dwg: Path, out_dir: Path, ver="ACAD2018"):
    """
    【修复权限问题】调用ODA转换单个DWG文件，改用系统临时目录
    :param oda_exe: ODA转换器路径
    :param src_dwg: 单个源DWG文件路径
    :param out_dir: 输出目录
    :param ver: 目标CAD版本
    :return: (返回码, 错误信息)
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    
    # 核心修复：使用系统临时目录（无中文、权限充足）
    with tempfile.TemporaryDirectory() as tmp_src:  # 自动创建/删除临时目录
        tmp_src_path = Path(tmp_src)
        tmp_dwg = tmp_src_path / src_dwg.name
        
        # 复制文件到系统临时目录（增加异常处理）
        try:
            shutil.copy2(src_dwg, tmp_dwg)
        except Exception as e:
            return -1, f"复制文件失败：{str(e)}"
        
        # 调用ODA转换临时目录中的单个DWG
        cmd = [str(oda_exe), str(tmp_src_path), str(out_dir), ver, "DXF", "0", "1"]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)  # 增加超时
        except subprocess.TimeoutExpired:
            return -2, "ODA转换超时（60秒）"
        except Exception as e:
            return -3, f"ODA调用失败：{str(e)}"
    
    return r.returncode, r.stderr

def find_dxf(out_dir: Path, stem: str):
    for ext in (".dxf",".DXF"):
        p = out_dir / (stem + ext)
        if p.exists():
            return p
    return None

# 核心清理逻辑
def clean_space(space, stats: dict, scale: float):
    tol = max(0.05, min(1.0, scale*1e-4))
    short_th = max(1.0, min(12.0, scale*0.002))

    def qkey(x,y): return (round(x/tol), round(y/tol))
    end_cnt = {}

    for e in space.query("LINE"):
        a,b = e.dxf.start, e.dxf.end
        end_cnt[qkey(a.x,a.y)] = end_cnt.get(qkey(a.x,a.y), 0) + 1
        end_cnt[qkey(b.x,b.y)] = end_cnt.get(qkey(b.x,b.y), 0) + 1

    removed = 0
    for e in list(space):
        t = e.dxftype()

        if t in ANN_TYPES or t in EXTRA_TYPES:
            space.delete_entity(e)
            stats[t] = stats.get(t, 0) + 1
            removed += 1
            continue

        try:
            layer = getattr(e.dxf, "layer", "") or ""
            ltype = getattr(e.dxf, "linetype", "") or ""
        except:
            layer, ltype = "", ""

        if t in GEOM and (kw_hit(layer) or kw_hit(ltype)):
            if t == "HATCH":
                continue
            space.delete_entity(e)
            stats["AUX_CEN_RULE"] = stats.get("AUX_CEN_RULE", 0) + 1
            removed += 1
            continue

        if t == "LINE":
            L = line_len(e)
            if L <= short_th:
                a,b = e.dxf.start, e.dxf.end
                if end_cnt.get(qkey(a.x,a.y),0)==1 and end_cnt.get(qkey(b.x,b.y),0)==1:
                    space.delete_entity(e)
                    stats["ISOL_SHORT"] = stats.get("ISOL_SHORT", 0) + 1
                    removed += 1
                    continue

    stats["SHORT_TOL"] = float(tol)
    stats["SHORT_TH"] = float(short_th)
    return removed

def clean_doc(doc, log):
    stats = {}
    scale = estimate_scale(doc)

    def spaces():
        yield ("model", doc.modelspace())
        for lay in doc.layouts:
            if lay.name.lower() != "model":
                yield (f"layout:{lay.name}", lay)
        for blk in doc.blocks:
            if not blk.name.startswith("*"):
                yield (f"block:{blk.name}", blk)

    total = 0
    for name, sp in spaces():
        total += clean_space(sp, stats, scale)
    log(f"[CLEAN] removed={total} scale≈{scale:.2f} stats={stats}")
    return stats

# 单文件处理流程
def process_one(oda: Path, dwg: Path, out_dir: Path, log):
    if not dwg.exists():
        log(f"[FAIL] missing: {dwg}")
        return None

    log(f"[1] ODA: {dwg.name} -> {out_dir}")
    rc, err = run_oda_single(oda, dwg, out_dir)
    if rc != 0:
        log(f"[FAIL] ODA rc={rc}\n{err.strip()}")
        return None

    dxf = find_dxf(out_dir, dwg.stem)
    if not dxf:
        log(f"[FAIL] DXF not found for {dwg.stem} in {out_dir}")
        return None

    log(f"[2] Load DXF: {dxf.name}")
    doc = ezdxf.readfile(dxf)

    log("[3] Clean annotations/center/aux/leftovers...")
    clean_doc(doc, log)

    out = out_dir / f"{dwg.stem}_clean.dxf"
    doc.saveas(out)
    
    try:
        dxf.unlink()
        log(f"[CLEANUP] 删除原始转换文件: {dxf.name}")
    except Exception as e:
        log(f"[WARN] 无法删除原始转换文件: {e}")
    
    log(f"[DONE] {dwg.name} -> {out.name}")
    return out

# GUI界面类
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DWG 语义预处理工具（批量去标注）")
        self.geometry("980x640")

        # 修改ODA默认路径为指定的地址
        self.oda = tk.StringVar(value=r"C:\Users\17117\Desktop\GraduationProject\stage2\ODA\ODAFileConverter.exe")
        self.mode = tk.StringVar(value="file")
        self.inp  = tk.StringVar(value="")
        self.out  = tk.StringVar(value=r"C:\Users\17117\Desktop\GraduationProject\stage2")
        self.last = None

        self.ui()

    def norm(self, p):
        return str(Path(p).expanduser().resolve())

    def ui(self):
        f = tk.Frame(self); f.pack(fill="x", padx=12, pady=10)
        tk.Label(f, text="ODAFileConverter.exe：").grid(row=0, column=0, sticky="w")
        tk.Entry(f, textvariable=self.oda, width=90).grid(row=0, column=1, sticky="w")
        tk.Button(f, text="选择", command=self.pick_oda).grid(row=0, column=2, padx=6)

        tk.Label(f, text="输入模式：").grid(row=1, column=0, sticky="w", pady=(10,0))
        tk.Radiobutton(f, text="单个（DWG）", variable=self.mode, value="file").grid(row=1, column=1, sticky="w", pady=(10,0))
        tk.Radiobutton(f, text="批量（文件夹）", variable=self.mode, value="folder").grid(row=1, column=1, sticky="w", padx=200, pady=(10,0))

        tk.Label(f, text="输入路径：").grid(row=2, column=0, sticky="w")
        tk.Entry(f, textvariable=self.inp, width=90).grid(row=2, column=1, sticky="w")
        tk.Button(f, text="选择", command=self.pick_in).grid(row=2, column=2, padx=6)

        tk.Label(f, text="输出目录：").grid(row=3, column=0, sticky="w", pady=(10,0))
        tk.Entry(f, textvariable=self.out, width=90).grid(row=3, column=1, sticky="w", pady=(10,0))
        tk.Button(f, text="选择", command=self.pick_out).grid(row=3, column=2, padx=6, pady=(10,0))

        b = tk.Frame(self); b.pack(fill="x", padx=12, pady=8)
        tk.Button(b, text="开始处理", width=16, command=self.start).pack(side="left")
        tk.Button(b, text="打开输出目录", width=16, command=self.open_out).pack(side="left", padx=8)
        tk.Button(b, text="定位结果文件", width=16, command=self.reveal).pack(side="left", padx=8)
        tk.Button(b, text="清空日志", width=16, command=lambda: self.logbox.delete("1.0","end")).pack(side="left", padx=8)

        self.logbox = tk.Text(self, height=28, wrap="word")
        self.logbox.pack(fill="both", expand=True, padx=12, pady=10)
        self.log("准备就绪。")

    def log(self, s):
        self.logbox.insert("end", s + "\n")
        self.logbox.see("end")
        self.update_idletasks()

    def pick_oda(self):
        p = filedialog.askopenfilename(title="选择 ODAFileConverter.exe", filetypes=[("EXE","*.exe")])
        if p:
            self.oda.set(self.norm(p))

    def pick_in(self):
        if self.mode.get() == "folder":
            p = filedialog.askdirectory("选择DWG文件夹")
        else:
            p = filedialog.askopenfilename(title="选择DWG文件", filetypes=[("DWG","*.dwg")])
        if p:
            self.inp.set(self.norm(p))

    def pick_out(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.out.set(self.norm(p))

    def open_out(self):
        try:
            p = Path(self.out.get().strip()).resolve()
            p.mkdir(parents=True, exist_ok=True)
            os.startfile(str(p))
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def reveal(self):
        try:
            if not self.last:
                return messagebox.showinfo("提示","还没有生成结果文件。")
            f = Path(self.last).resolve()
            if not f.exists():
                return messagebox.showerror("错误", f"文件不存在：\n{f}")
            subprocess.run(["explorer","/select,", str(f)])
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def start(self):
        oda = Path(self.oda.get().strip())
        inp = Path(self.inp.get().strip())
        out = Path(self.out.get().strip())
        
        if not oda.exists():
            return messagebox.showerror("错误", f"找不到ODA：\n{oda}")
        if not inp.exists():
            return messagebox.showerror("错误", f"输入不存在：\n{inp}")
        
        out.mkdir(parents=True, exist_ok=True)

        dwgs = [inp] if inp.is_file() else sorted(inp.glob("*.dwg"))
        if not dwgs:
            return messagebox.showwarning("提示", "未找到 *.dwg")

        self.log(f"即将处理 {len(dwgs)} 个DWG，输出到：{out}")
        self.log("-"*60)
        threading.Thread(target=self.run_batch, args=(oda,dwgs,out), daemon=True).start()

    def run_batch(self, oda, dwgs, out):
        ok = 0
        for i, dwg in enumerate(dwgs, 1):
            self.log(f"\n[{i}/{len(dwgs)}] {dwg}")
            res = process_one(oda, dwg, out, self.log)
            if res:
                ok += 1
                self.last = str(res)
        self.log("\n完成")
        self.log(f"总计：{len(dwgs)} 成功：{ok} 失败：{len(dwgs)-ok}")

# 程序入口
if __name__ == "__main__":
    App().mainloop()