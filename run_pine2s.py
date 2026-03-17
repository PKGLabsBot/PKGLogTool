#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PINE2S#2 전용 로그 분석 런너
- H 폴더: System.txt (ShowErrorMSG), AjinError.txt (ErrorCode) 파싱
- V 폴더: Vision.txt / Total.txt (탭 구분, Error 레벨)
- 날짜는 폴더명(YYYYMMDD) 또는 파일명(YYYY-MM-DD_)에서 추출
"""
import sys
import os
import re
import traceback
import types

# ── tkinter mock (logtool.py의 App 클래스가 tk.Tk 상속하므로 정상 클래스 필요) ──
class _TkBase:
    def __init__(self, *a, **kw): pass
    def after(self, *a, **kw): pass
    def mainloop(self, *a, **kw): pass

class _StringVar:
    def __init__(self, *a, **kw): self._val = ""
    def get(self): return self._val
    def set(self, v): self._val = v

class _IntVar:
    def __init__(self, *a, **kw): self._val = 0
    def get(self): return self._val
    def set(self, v): self._val = v

class _BoolVar:
    def __init__(self, *a, **kw): self._val = False
    def get(self): return self._val
    def set(self, v): self._val = v

_tk = types.ModuleType("tkinter")
_tk.Tk = _TkBase
_tk.Frame = _TkBase
_tk.Label = _TkBase
_tk.Button = _TkBase
_tk.Entry = _TkBase
_tk.StringVar = _StringVar
_tk.IntVar = _IntVar
_tk.BooleanVar = _BoolVar
_tk.END = "end"
_tk.BOTH = "both"
_tk.LEFT = "left"
_tk.RIGHT = "right"
_tk.TOP = "top"
_tk.BOTTOM = "bottom"
_tk.N = _tk.S = _tk.E = _tk.W = _tk.NE = _tk.NW = _tk.SE = _tk.SW = ""
_tk.HORIZONTAL = "horizontal"

_ttk = types.ModuleType("tkinter.ttk")
class _Widget:
    def __init__(self, *a, **kw): pass
    def pack(self, *a, **kw): pass
    def grid(self, *a, **kw): pass
    def place(self, *a, **kw): pass
    def config(self, *a, **kw): pass
    def configure(self, *a, **kw): pass
    def __setitem__(self, k, v): pass
_ttk.Combobox = _Widget
_ttk.Progressbar = _Widget
_ttk.Frame = _Widget
_ttk.Label = _Widget
_ttk.Button = _Widget
_ttk.Entry = _Widget
_ttk.Notebook = _Widget

_fd = types.ModuleType("tkinter.filedialog")
_fd.askdirectory = lambda **kw: ""
_fd.asksaveasfilename = lambda **kw: ""

_mb = types.ModuleType("tkinter.messagebox")
_mb.showerror = lambda *a, **kw: None
_mb.showinfo  = lambda *a, **kw: None
_mb.askyesno  = lambda *a, **kw: False

sys.modules["tkinter"]             = _tk
sys.modules["tkinter.ttk"]         = _ttk
sys.modules["tkinter.filedialog"]  = _fd
sys.modules["tkinter.messagebox"]  = _mb

# ── sys.path 설정 ──────────────────────────────────────────────
LOGTOOL_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, LOGTOOL_DIR)
os.chdir(LOGTOOL_DIR)

from logtool import (
    build_dfs, write_excel, render_output_name,
    OUTPUT_DIR, write_debug, clear_debug_files, APP_VERSION
)

from collections import Counter
from datetime import datetime, date
from pathlib import Path
from typing import Optional

# ============================================================
# 설정
# ============================================================
EQUIP         = "PINE2S-Series"
HANDLER_ROOT  = r"C:\Users\tom\PKGTool\01_inbox\PINE2S log\PINE2S#2 log\H"
VISION_ROOT   = r"C:\Users\tom\PKGTool\01_inbox\PINE2S log\PINE2S#2 log\V"
TOP_N         = 30
SPIKE_FACTOR  = 2.0
SPIKE_WINDOW  = 7
HANDLER_TARGETS = ["System.txt", "AjinError.txt"]

# ============================================================
# 패턴
# ============================================================
RE_SYSTEM      = re.compile(r'ShowErrorMSG Error:\s*(?P<name>[^(\r\n\-]+?)\s*\((?P<code>\d+)\)', re.IGNORECASE)
RE_AJIN        = re.compile(r'-\s+(?P<desc>.+?),\s*ErrorCode:\s*(?P<code>\d+)', re.IGNORECASE)
RE_H_TIME      = re.compile(r'^(\d{2}):(\d{2}):(\d{2}):(\d{3})')
RE_V_LINE      = re.compile(r'^(\d{2})\.(\d{2})\.(\d{2})\.(\d{3})\t(\w+)\t(\w+)\t(.+)', re.DOTALL)
RE_V_ERR_CODE  = re.compile(r'(?:Error\s+Code\s*[=:]\s*|Error\s*\(\s*)(?P<code>\d+)', re.IGNORECASE)

# ============================================================
# 유틸
# ============================================================
def parse_yyyymmdd(s: str) -> Optional[date]:
    try:    return datetime.strptime(s.strip(), '%Y%m%d').date()
    except: return None

def parse_yyyymmdd_dash(s: str) -> Optional[date]:
    try:    return datetime.strptime(s.strip(), '%Y-%m-%d').date()
    except: return None

def make_ts(d: date, h: int, m: int, s: int, ms: int) -> datetime:
    return datetime(d.year, d.month, d.day, h, m, s, ms * 1000)

def _make_info(pc_type, files_count, files_ok, files_fail, total_lines, error_lines, ts_min, ts_max):
    return {"pc_type": pc_type, "files_count": files_count, "files_read_ok": files_ok,
            "files_read_fail": files_fail, "total_lines_read": total_lines,
            "error_lines_counted": error_lines, "start_time": ts_min,
            "end_time": ts_max, "encoding": "utf-8", "equipment": EQUIP}

def _make_counters(cnt_code_by_pc, cnt_day_code_pc, cnt_hour_code_pc, name_map):
    return {"cnt_code_by_pc": cnt_code_by_pc, "cnt_day_code_pc": cnt_day_code_pc,
            "cnt_hour_code_pc": cnt_hour_code_pc, "name_map": name_map}

# ============================================================
# Handler 분석
# ============================================================
def analyze_handler(handler_root: str):
    cnt_code_by_pc = Counter(); cnt_day_code_pc = Counter()
    cnt_hour_code_pc = Counter(); name_map = {}
    ts_min = ts_max = None
    total_lines = error_lines = files_ok = files_fail = 0

    root = Path(handler_root)
    if not root.exists():
        print(f"  [경고] Handler 루트 없음: {handler_root}")
        return _make_info("HANDLER",0,0,0,0,0,None,None), _make_counters(cnt_code_by_pc,cnt_day_code_pc,cnt_hour_code_pc,name_map)

    date_dirs = sorted(d for d in root.iterdir() if d.is_dir())
    print(f"  날짜 폴더: {len(date_dirs)}개")

    for date_dir in date_dirs:
        folder_date = parse_yyyymmdd(date_dir.name)
        if folder_date is None:
            continue

        for fname in HANDLER_TARGETS:
            fpath = date_dir / fname
            if not fpath.exists():
                continue

            try:
                with open(fpath, encoding='utf-8', errors='replace') as f:
                    files_ok += 1
                    for line in f:
                        total_lines += 1
                        if 'Error' not in line:
                            continue

                        code = name = None
                        ts = None

                        tm = RE_H_TIME.match(line.strip())
                        if tm:
                            try:
                                ts = make_ts(folder_date, int(tm.group(1)), int(tm.group(2)),
                                             int(tm.group(3)), int(tm.group(4)))
                            except Exception:
                                ts = datetime(folder_date.year, folder_date.month, folder_date.day)

                        m1 = RE_SYSTEM.search(line)
                        if m1:
                            code = m1.group('code')
                            name = m1.group('name').strip()

                        if not code:
                            m2 = RE_AJIN.search(line)
                            if m2:
                                code = m2.group('code')
                                desc = m2.group('desc') or ''
                                name = desc.split(',')[0].strip()[:60]

                        if not code:
                            continue

                        error_lines += 1
                        code_str = str(code)
                        name = name or ''
                        name_map[('HANDLER', code_str)] = name
                        cnt_code_by_pc[('HANDLER', code_str, name)] += 1

                        if ts:
                            if ts_min is None or ts < ts_min: ts_min = ts
                            if ts_max is None or ts > ts_max: ts_max = ts
                            cnt_day_code_pc[(ts.date(), 'HANDLER', code_str, name)] += 1
                            cnt_hour_code_pc[(ts.hour, 'HANDLER', code_str, name)] += 1

            except Exception as e:
                files_fail += 1
                write_debug(f"[HANDLER] read fail | {fpath} | {e}\n{traceback.format_exc()}")

    print(f"  파일: {files_ok}개 성공 / {files_fail}개 실패")
    print(f"  에러 라인: {error_lines:,}개 / 총 {total_lines:,}줄")

    return (_make_info("HANDLER", files_ok+files_fail, files_ok, files_fail, total_lines, error_lines, ts_min, ts_max),
            _make_counters(cnt_code_by_pc, cnt_day_code_pc, cnt_hour_code_pc, name_map))

# ============================================================
# Vision 분석
# ============================================================
def analyze_vision(vision_root: str):
    cnt_code_by_pc = Counter(); cnt_day_code_pc = Counter()
    cnt_hour_code_pc = Counter(); name_map = {}
    ts_min = ts_max = None
    total_lines = error_lines = files_ok = files_fail = 0
    total_files_found = 0

    root = Path(vision_root)
    if not root.exists():
        print(f"  [경고] Vision 루트 없음: {vision_root}")
        return _make_info("VISION",0,0,0,0,0,None,None), _make_counters(cnt_code_by_pc,cnt_day_code_pc,cnt_hour_code_pc,name_map)

    for cam_dir in sorted(root.iterdir()):
        if not cam_dir.is_dir():
            continue
        for date_dir in sorted(cam_dir.iterdir()):
            if not date_dir.is_dir():
                continue
            folder_date = parse_yyyymmdd_dash(date_dir.name)
            if folder_date is None:
                continue

            for fname in [f"{date_dir.name}_Total.txt", f"{date_dir.name}_Vision.txt"]:
                fpath = date_dir / fname
                if not fpath.exists():
                    continue
                total_files_found += 1

                try:
                    with open(fpath, encoding='utf-8', errors='replace') as f:
                        files_ok += 1
                        for line in f:
                            total_lines += 1
                            if 'Error' not in line:
                                continue

                            vm = RE_V_LINE.match(line)
                            if not vm:
                                continue

                            h, m, s, ms = int(vm.group(1)), int(vm.group(2)), int(vm.group(3)), int(vm.group(4))
                            level   = vm.group(5)
                            message = vm.group(7)

                            code = name = None
                            cm = RE_V_ERR_CODE.search(message)
                            if cm:
                                code = cm.group('code')
                                name = message[:80].strip()
                            elif level.lower() == 'error':
                                code = '999'
                                name = message[:80].strip()

                            if not code:
                                continue

                            error_lines += 1
                            code_str = str(code)
                            pc_key   = f"VISION-{cam_dir.name}"
                            name_map[(pc_key, code_str)] = name
                            cnt_code_by_pc[(pc_key, code_str, name)] += 1

                            try:
                                ts = make_ts(folder_date, h, m, s, ms)
                                if ts_min is None or ts < ts_min: ts_min = ts
                                if ts_max is None or ts > ts_max: ts_max = ts
                                cnt_day_code_pc[(ts.date(), pc_key, code_str, name)] += 1
                                cnt_hour_code_pc[(ts.hour,  pc_key, code_str, name)] += 1
                            except Exception:
                                pass

                except Exception as e:
                    files_fail += 1
                    write_debug(f"[VISION] read fail | {fpath} | {e}\n{traceback.format_exc()}")

    print(f"  Total/Vision 파일: {total_files_found}개 탐색")
    print(f"  파일: {files_ok}개 성공 / {files_fail}개 실패")
    print(f"  에러 라인: {error_lines:,}개 / 총 {total_lines:,}줄")

    return (_make_info("VISION", files_ok+files_fail, files_ok, files_fail, total_lines, error_lines, ts_min, ts_max),
            _make_counters(cnt_code_by_pc, cnt_day_code_pc, cnt_hour_code_pc, name_map))

# ============================================================
# 실행
# ============================================================
if __name__ == "__main__":
    print(f"[PINE2S#2 로그 분석]  설비군: {EQUIP}  (Tool v{APP_VERSION})")
    print(f"  Handler: {HANDLER_ROOT}")
    print(f"  Vision : {VISION_ROOT}")
    print()

    clear_debug_files()
    write_debug("===== PINE2S RUN START =====")

    all_infos = []; all_counters = []

    print("[ 1/2 ] Handler 분석 중...")
    info_h, cnts_h = analyze_handler(HANDLER_ROOT)
    all_infos.append(info_h); all_counters.append(cnts_h)

    print()
    print("[ 2/2 ] Vision 분석 중...")
    info_v, cnts_v = analyze_vision(VISION_ROOT)
    all_infos.append(info_v); all_counters.append(cnts_v)

    print()
    print("집계 중...")
    summary_df, top_all, top_by_pc, by_day, by_hour, anomalies = build_dfs(
        all_infos=all_infos, all_counters=all_counters,
        top_n=TOP_N, spike_factor=SPIKE_FACTOR, spike_window=SPIKE_WINDOW,
    )

    out_name = f"PINE2S-Series_v{APP_VERSION}_{datetime.now().strftime('%Y%m%d')}_log_report.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    print(f"Excel 작성 중... -> {out_path}")
    write_excel(
        out_path=out_path, summary_df=summary_df, top_all=top_all,
        top_by_pc=top_by_pc, by_day=by_day, by_hour=by_hour,
        anomalies=anomalies, chart_mode="A",
    )

    write_debug("===== PINE2S RUN END =====")
    print(f"\n[완료] 리포트: {out_path}")
