#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LogTool 헤드리스 실행 스크립트 - PINE2S-Series
GUI 없이 커맨드라인으로 분석 실행
"""
import sys
import os

# logtool 디렉토리를 sys.path에 추가
LOGTOOL_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, LOGTOOL_DIR)
os.chdir(LOGTOOL_DIR)

# tkinter 없이 import 할 수 없으므로 tkinter를 mock으로 교체
import types
tk_mock = types.ModuleType("tkinter")
tk_mock.Tk = object
tk_mock.StringVar = object
tk_mock.IntVar = object
tk_mock.BooleanVar = object
sys.modules.setdefault("tkinter", tk_mock)
ttk_mock = types.ModuleType("tkinter.ttk")
sys.modules.setdefault("tkinter.ttk", ttk_mock)
fd_mock = types.ModuleType("tkinter.filedialog")
sys.modules.setdefault("tkinter.filedialog", fd_mock)
mb_mock = types.ModuleType("tkinter.messagebox")
sys.modules.setdefault("tkinter.messagebox", mb_mock)

# 핵심 함수만 직접 임포트
from logtool import (
    load_profiles, iter_files, aggregate_logs, build_dfs, write_excel,
    RegexParser, render_output_name, write_debug, clear_debug_files, OUTPUT_DIR
)

# ============================================================
# 설정
# ============================================================
EQUIP        = "PINE2S-Series"
HANDLER_ROOT = r"C:\Users\tom\PKGTool\01_inbox\PINE2S log\PINE2S#2 log\H"
VISION_ROOT  = r"C:\Users\tom\PKGTool\01_inbox\PINE2S log\PINE2S#2 log\V"
TOP_N        = 30
SPIKE_FACTOR = 2.0
SPIKE_WINDOW = 7

# ============================================================
# 실행
# ============================================================
print(f"[LogTool 헤드리스] 설비군: {EQUIP}")
print(f"  Handler: {HANDLER_ROOT}")
print(f"  Vision : {VISION_ROOT}")

profiles = load_profiles()
prof = profiles.get(EQUIP)
if not prof:
    print(f"ERROR: 프로파일을 찾을 수 없습니다: {EQUIP}")
    sys.exit(1)

chart_mode       = (prof.get("chart_mode") or "A").strip().upper()
handler_prefix   = prof.get("handler_prefix", "")
vision_prefix    = prof.get("vision_prefix", "")
vision_only_err  = bool(prof.get("vision_only_error", False))
encoding         = "auto"

hp = prof.get("handler_parser", {})
vp = prof.get("vision_parser", {})

handler_parser = RegexParser(
    patterns=hp.get("patterns", []),
    dt_formats=hp.get("dt_formats", ["%Y-%m-%d %H:%M:%S"]),
    parser_name=f"{EQUIP}-HANDLER"
)
vision_parser = RegexParser(
    patterns=vp.get("patterns", []),
    dt_formats=vp.get("dt_formats", ["%Y-%m-%d %H:%M:%S"]),
    parser_name=f"{EQUIP}-VISION"
)

# 파일 수집
handler_files = iter_files(HANDLER_ROOT, prefix=handler_prefix)
if not handler_files:
    handler_files = iter_files(HANDLER_ROOT, prefix="")

vision_files = iter_files(VISION_ROOT, prefix=vision_prefix)
if not vision_files:
    vision_files = iter_files(VISION_ROOT, prefix="")

print(f"  Handler 파일: {len(handler_files)}개")
print(f"  Vision  파일: {len(vision_files)}개")

total_files = len(handler_files) + len(vision_files)
processed   = 0

def progress_cb(pc_type, idx, total_f, fp):
    current = processed + idx
    pct = (current / total_files * 100) if total_files > 0 else 0
    print(f"\r  [{pc_type}] {current}/{total_files} ({pct:.0f}%) - {os.path.basename(fp):<40}", end="", flush=True)

all_infos    = []
all_counters = []

clear_debug_files()
write_debug("===== HEADLESS ANALYSIS START =====")
write_debug(f"equip={EQUIP}, handler={HANDLER_ROOT}, vision={VISION_ROOT}")

if handler_files:
    print("\nHandler 로그 분석 중...")
    info_h, cnts_h = aggregate_logs(
        files=handler_files,
        pc_type="HANDLER",
        encoding=encoding,
        only_error_lines=True,
        only_codes_set=None,
        parser=handler_parser,
        progress_cb=progress_cb,
    )
    info_h["equipment"] = EQUIP
    all_infos.append(info_h)
    all_counters.append(cnts_h)
    processed += len(handler_files)
    print(f"\n  완료 (에러 라인: {info_h.get('error_lines_counted', 0):,})")

if vision_files:
    print("\nVision 로그 분석 중...")
    info_v, cnts_v = aggregate_logs(
        files=vision_files,
        pc_type="VISION",
        encoding=encoding,
        only_error_lines=vision_only_err,
        only_codes_set=None,
        parser=vision_parser,
        progress_cb=progress_cb,
    )
    info_v["equipment"] = EQUIP
    all_infos.append(info_v)
    all_counters.append(cnts_v)
    processed += len(vision_files)
    print(f"\n  완료 (에러 라인: {info_v.get('error_lines_counted', 0):,})")

print("\n집계 데이터 생성 중...")
summary_df, top_all, top_by_pc, by_day, by_hour, anomalies = build_dfs(
    all_infos=all_infos,
    all_counters=all_counters,
    top_n=TOP_N,
    spike_factor=SPIKE_FACTOR,
    spike_window=SPIKE_WINDOW,
)

out_name = render_output_name(
    prof.get("output_name", "{equip}_{date}_log_report.xlsx"), EQUIP
)
out_path = os.path.join(OUTPUT_DIR, out_name)

print(f"Excel 작성 중...")
print(f"  출력 경로: {out_path}")
write_excel(
    out_path=out_path,
    summary_df=summary_df,
    top_all=top_all,
    top_by_pc=top_by_pc,
    by_day=by_day,
    by_hour=by_hour,
    anomalies=anomalies,
    chart_mode=chart_mode,
)

write_debug("===== HEADLESS ANALYSIS END =====")
print(f"\n[완료] 리포트: {out_path}")
