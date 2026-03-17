#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import json
import copy
import threading
import traceback
from glob import glob
from datetime import datetime
from collections import Counter
from typing import Dict, Optional, List, Any, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


APP_NAME = "로그 분석 Tool"

# ============================================================
# 경로
# ============================================================
def app_dir() -> str:
    """실행 파일(frozen) 또는 스크립트 기준 디렉토리 반환."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_version() -> str:
    version_file = os.path.join(app_dir(), "version.txt")
    try:
        with open(version_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return "1.0"


APP_VERSION = get_version()

EQUIP_LIST = [
    "CYPRESS1",
    "CYPRESS2",
    "CYPRESS2+",
    "MAPLE",
    "PINE1",
    "PINE2",
    "PINE2S-Series",
    "JEDI",
    "AVIS",
    "LMS",
    "ASIS",
]


BASE_DIR = app_dir()
PROFILES_DIR = os.path.join(BASE_DIR, "profiles")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DEBUG_LOG = os.path.join(BASE_DIR, "debug_log.txt")
DEBUG_SAMPLES = os.path.join(BASE_DIR, "debug_samples.txt")

os.makedirs(PROFILES_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# 공통
# ============================================================
def safe_filename(s: str) -> str:
    bad = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for ch in bad:
        s = s.replace(ch, "_")
    return s


def render_output_name(template: str, equip: str) -> str:
    today = datetime.now().strftime("%Y%m%d")
    try:
        name = template.format(equip=safe_filename(equip), date=today, version=APP_VERSION)
    except Exception:
        name = f"{safe_filename(equip)}_{today}_log_report.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def write_debug(msg: str) -> None:
    try:
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{now}] {msg}\n")
    except Exception:
        pass


def write_debug_sample(msg: str) -> None:
    try:
        with open(DEBUG_SAMPLES, "a", encoding="utf-8") as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{now}] {msg}\n")
    except Exception:
        pass


def clear_debug_files() -> None:
    for p in [DEBUG_LOG, DEBUG_SAMPLES]:
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

def open_log_file(path: str, encoding: str):

    # 자동 인코딩 탐색
    encodings = ["utf-8", "utf-8-sig", "cp949", "euc-kr", "utf-16", "latin1"]

    for enc in encodings:
        try:
            return open(path, "r", encoding=enc, errors="ignore")
        except Exception:
            continue

    return open(path, "r", encoding="utf-8", errors="ignore")

def parse_dt(dt_str: str, fmts: List[str]) -> Optional[datetime]:
    dt_str = (dt_str or "").strip()
    if not dt_str:
        return None

    for fmt in fmts:
        try:
            return datetime.strptime(dt_str, fmt)
        except Exception:
            pass

    try:
        return datetime.fromisoformat(dt_str)
    except Exception:
        pass

    return None


# ============================================================
# 기본 프로파일 생성
# ============================================================
def default_profile_for(eq: str) -> Dict[str, Any]:
    return {
        "display": eq,
        "handler_prefix": "Error_",
        "vision_prefix": "Network_",
        "vision_only_error": True,
        "encoding": "utf-8",
        "chart_mode": "A",
        "top_n": 30,
        "output_name": "{equip}_v{version}_{date}_log_report.xlsx",

        # CYPRESS1 Handler 예시:
        # 2025-10-11 16:12:27:013    Error Code : 210, Error Name : PCB Reject Error
        "handler_parser": {
            "dt_formats": ["%Y-%m-%d %H:%M:%S:%f", "%Y-%m-%d %H:%M:%S"],
            "patterns": [
                r"(?P<date>\d{4}-\d{2}-\d{2}).*(?P<time>\d{2}:\d{2}:\d{2}):(?P<ms>\d{3}).*Error\s*Code\s*:\s*(?P<code>\d+)\s*,\s*Error\s*Name\s*:\s*(?P<name>.*)",
                r"Error\s*Code\s*:\s*(?P<code>\d+)\s*,\s*Error\s*Name\s*:\s*(?P<name>.*)$",
            ],
        },

        # CYPRESS1 Vision 예시:
        # [2025-11-04 09:48:32] SendMessageToHandler() - Error 10015 1202 W/F Align 실패
        "vision_parser": {
            "dt_formats": ["%Y-%m-%d %H:%M:%S"],
            "patterns": [
                r"\[(?P<dt>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\].*Error\s+(?P<code>\d+)\s+\d+\s+(?P<name>.*)",
                r".*Error\s+(?P<code>\d+)\s+\d+\s+(?P<name>.*)",
                r".*Error\s+(?P<code>\d+)\s+(?P<name>.*)"
            ],
        },
    }


def ensure_default_profiles() -> None:
    for eq in EQUIP_LIST:
        path = os.path.join(PROFILES_DIR, f"{eq}.json")
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                json.dump(default_profile_for(eq), f, ensure_ascii=False, indent=2)
            write_debug(f"default profile created: {path}")


def load_profiles() -> Dict[str, Any]:
    ensure_default_profiles()
    profiles = {}

    for file in sorted(os.listdir(PROFILES_DIR)):
        if not file.lower().endswith(".json"):
            continue

        path = os.path.join(PROFILES_DIR, file)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            name = data.get("display") or os.path.splitext(file)[0]
            profiles[name] = data
        except Exception as e:
            write_debug(f"profile load fail: file={file}, error={e}")

    for eq in EQUIP_LIST:
        if eq not in profiles:
            profiles[eq] = copy.deepcopy(default_profile_for(eq))

    return profiles


# ============================================================
# 프로파일 기반 정규식 파서
# ============================================================
class RegexParser:
    def __init__(self, patterns: List[str], dt_formats: List[str], parser_name: str):
        self.parser_name = parser_name
        self.dt_formats = dt_formats or ["%Y-%m-%d %H:%M:%S"]
        self.regexes = []
        self.parse_fail_count = 0
        self.max_fail_samples = 50

        for p in (patterns or []):
            try:
                self.regexes.append(re.compile(p, re.IGNORECASE))
            except Exception as e:
                write_debug(f"regex compile fail | parser={self.parser_name} | pattern={p} | error={e}")

    def parse(self, line: str) -> Dict[str, Any]:
        raw = line.rstrip("\n")

        for rx in self.regexes:
            m = rx.search(raw)
            if not m:
                continue

            gd = m.groupdict()
            code = (gd.get("code") or "").strip()
            name = (gd.get("name") or "").strip()
            ts = None

            if gd.get("dt"):
                ts = parse_dt(gd.get("dt"), self.dt_formats)
            elif gd.get("date") and gd.get("time"):
                date = (gd.get("date") or "").strip()
                time = (gd.get("time") or "").strip()
                ms = (gd.get("ms") or "").strip()

                if ms:
                    micro = ms + "000" if len(ms) == 3 else ms
                    dt_value = f"{date} {time}:{micro}"
                else:
                    dt_value = f"{date} {time}"

                ts = parse_dt(dt_value, self.dt_formats)

            if not code:
                return {"timestamp": ts, "error_code": None, "error_name": None}

            if ts is None and ("dt" in gd or ("date" in gd and "time" in gd)):
                if self.parse_fail_count < self.max_fail_samples:
                    write_debug_sample(f"[{self.parser_name}] datetime parse failed | line={raw[:300]}")
                    self.parse_fail_count += 1

            return {"timestamp": ts, "error_code": str(code), "error_name": name}

        if self.parse_fail_count < self.max_fail_samples:
            write_debug_sample(f"[{self.parser_name}] regex parse fail | line={raw[:300]}")
            self.parse_fail_count += 1

        return {"timestamp": None, "error_code": None, "error_name": None}


# ============================================================
# 파일 수집
# ============================================================
def iter_files(root_or_pattern: str, prefix: str) -> List[str]:
    if not root_or_pattern:
        return []

    if any(ch in root_or_pattern for ch in ["*", "?", "["]):
        return sorted(set([f for f in glob(root_or_pattern) if os.path.isfile(f)]))

    if os.path.isfile(root_or_pattern):
        return [root_or_pattern]

    if not os.path.isdir(root_or_pattern):
        raise FileNotFoundError(f"Input not found: {root_or_pattern}")

    out = []
    prefix_l = (prefix or "").lower().strip()

    for root, _, files in os.walk(root_or_pattern):
        for fn in files:
            low = fn.lower()

            # 확장자 검사
            if not (low.endswith(".txt") or low.endswith(".log")):
                continue

            # prefix가 있으면 "시작"이 아니라 "포함"으로 검사
            if prefix_l and prefix_l not in low:
                continue

            out.append(os.path.join(root, fn))

    return sorted(set(out))


# ============================================================
# 집계
# ============================================================
def aggregate_logs(files: List[str],
                   pc_type: str,
                   encoding: str,
                   only_error_lines: bool,
                   only_codes_set: Optional[set],
                   parser: RegexParser,
                   progress_cb=None) -> Tuple[dict, dict]:

    cnt_code_by_pc = Counter()
    cnt_day_code_pc = Counter()
    cnt_hour_code_pc = Counter()
    name_map = {}

    total_lines_read = 0
    error_lines_counted = 0
    ts_min = None
    ts_max = None
    files_read_ok = 0
    files_read_fail = 0

    total_files = len(files)

    for idx, fp in enumerate(files, start=1):
        if progress_cb:
            progress_cb(pc_type, idx, total_files, fp)

        try:
            with open_log_file(fp, encoding) as f:
                files_read_ok += 1

                for line_no, line in enumerate(f, start=1):

                    # if pc_type == "VISION":
                    #    print("VISION LINE:", line.strip())

                    total_lines_read += 1

                    if "Error" not in line:
                        continue

                    
                    d = parser.parse(line)

                    code = d.get("error_code")
                    name = (d.get("error_name") or "").strip()
                    ts = d.get("timestamp")

                    if only_error_lines and not code:
                        continue
                    if not code:
                        continue

                    code = str(code).strip()
                    if only_codes_set is not None and code not in only_codes_set:
                        continue

                    error_lines_counted += 1
                    name_map[(pc_type, code)] = name

                    cnt_code_by_pc[(pc_type, code, name)] += 1

                    if ts is not None:
                        if ts_min is None or ts < ts_min:
                            ts_min = ts
                        if ts_max is None or ts > ts_max:
                            ts_max = ts

                        cnt_day_code_pc[(ts.date(), pc_type, code, name)] += 1
                        cnt_hour_code_pc[(ts.hour, pc_type, code, name)] += 1

        except Exception as e:
            files_read_fail += 1
            write_debug(
                f"aggregate error | pc_type={pc_type} | file={fp} | error={e}\n{traceback.format_exc()}"
            )
            continue

    info = {
        "pc_type": pc_type,
        "files_count": len(files),
        "files_read_ok": files_read_ok,
        "files_read_fail": files_read_fail,
        "total_lines_read": total_lines_read,
        "error_lines_counted": error_lines_counted,
        "start_time": ts_min,
        "end_time": ts_max,
        "encoding": encoding,
    }

    counters = {
        "cnt_code_by_pc": cnt_code_by_pc,
        "cnt_day_code_pc": cnt_day_code_pc,
        "cnt_hour_code_pc": cnt_hour_code_pc,
        "name_map": name_map,
    }

    return info, counters

# ============================================================
# Error 장비 분류
# ============================================================
DEVICE_RULES = {
    "RFID": ["RFID"],
    "OHT": ["OHT"],
    "SERVER": ["LOT", "RECIPE", "RESULT", "PMS"]
}

def classify_device(error_name: str, pc_type: str) -> str:

    name = (error_name or "").upper()

    # 키워드 기반 분류
    for device, keywords in DEVICE_RULES.items():
        for k in keywords:
            if k in name:
                return device

    # Vision 에러
    if pc_type == "VISION":
        return "VISION"

    # 나머지는 Handler
    return "HANDLER"

# ============================================================
# 급증 탐지
# ============================================================
def compute_spikes_by_code(day_code_df: pd.DataFrame, factor: float = 2.0, window: int = 7) -> pd.DataFrame:
    if day_code_df.empty:
        return pd.DataFrame(columns=["date", "error_code", "count", "prev_day", "rolling_mean", "reason"])

    out = day_code_df.sort_values(["error_code", "date"]).reset_index(drop=True)
    spikes = []

    for code, g in out.groupby("error_code"):
        g = g.sort_values("date").copy()
        g["prev"] = g["count"].shift(1)
        g["rolling_mean"] = g["count"].shift(1).rolling(window=window, min_periods=max(3, window // 2)).mean()

        for _, r in g.iterrows():
            c = float(r["count"])
            prev = r["prev"]
            rm = r["rolling_mean"]
            reasons = []

            if pd.notna(prev) and prev > 0 and c >= float(prev) * factor and (c - float(prev)) >= 5:
                reasons.append(f"전일({int(prev)}) 대비 {factor}배↑")

            if pd.notna(rm) and rm > 0 and c >= float(rm) * factor and (c - float(rm)) >= 5:
                reasons.append(f"{window}일 평균({rm:.1f}) 대비 {factor}배↑")

            if reasons:
                spikes.append({
                    "date": r["date"],
                    "error_code": str(code),
                    "count": int(r["count"]),
                    "prev_day": int(prev) if pd.notna(prev) else None,
                    "rolling_mean": float(rm) if pd.notna(rm) else None,
                    "reason": " / ".join(reasons),
                })

    return pd.DataFrame(spikes)


# ============================================================
# DF 생성
# ============================================================
def build_dfs(all_infos: List[dict], all_counters: List[dict],
              top_n: int, spike_factor: float, spike_window: int):

    summary_df = pd.DataFrame(all_infos)

    cnt_code_all = Counter()
    cnt_code_by_pc = Counter()
    cnt_day_code_pc = Counter()
    cnt_hour_code_pc = Counter()
    name_map = {}

    for c in all_counters:
        cnt_code_by_pc.update(c["cnt_code_by_pc"])
        cnt_day_code_pc.update(c["cnt_day_code_pc"])
        cnt_hour_code_pc.update(c["cnt_hour_code_pc"])
        name_map.update(c["name_map"])

    for (pc, code, name), v in cnt_code_by_pc.items():
        cnt_code_all[(code, name)] += v

    top_all = pd.DataFrame(
        [(code, name, cnt) for (code, name), cnt in cnt_code_all.most_common(top_n)],
        columns=["error_code", "error_name", "count"]
    )

    rows = [(pc, code, name, cnt) for (pc, code, name), cnt in cnt_code_by_pc.items()]
    top_by_pc = pd.DataFrame(rows, columns=["pc_type", "error_code", "error_name", "count"])
    if not top_by_pc.empty:
        top_by_pc = top_by_pc.sort_values(["pc_type", "count"], ascending=[True, False])
        top_by_pc = top_by_pc.groupby("pc_type", as_index=False, group_keys=False).head(top_n)

    rows = [(day, pc, code, name, cnt) for (day, pc, code, name), cnt in cnt_day_code_pc.items()]
    by_day = pd.DataFrame(rows, columns=["date", "pc_type", "error_code", "error_name", "count"])
    if not by_day.empty:
        by_day = by_day.sort_values(["date", "pc_type", "count"], ascending=[True, True, False])

    rows = [(hour, pc, code, name, cnt) for (hour, pc, code, name), cnt in cnt_hour_code_pc.items()]
    by_hour = pd.DataFrame(rows, columns=["hour", "pc_type", "error_code", "error_name", "count"])
    if not by_hour.empty:
        by_hour = by_hour.sort_values(["hour", "pc_type", "count"], ascending=[True, True, False])

    anomalies_list = []
    if not by_day.empty:
        for pc, g in by_day.groupby("pc_type"):
            g2 = g.groupby(["date", "error_code"], as_index=False)["count"].sum()
            spikes = compute_spikes_by_code(g2, factor=spike_factor, window=spike_window)
            if not spikes.empty:
                spikes.insert(0, "pc_type", pc)
                spikes["error_name"] = spikes["error_code"].map(lambda code: name_map.get((pc, str(code)), ""))
                anomalies_list.append(spikes)

    anomalies = pd.concat(anomalies_list, ignore_index=True) if anomalies_list else pd.DataFrame(
        columns=["pc_type", "date", "error_code", "error_name", "count", "prev_day", "rolling_mean", "reason"]
    )

    return summary_df, top_all, top_by_pc, by_day, by_hour, anomalies


# ============================================================
# Excel 보조
# ============================================================
def autosize_worksheet(ws, df: pd.DataFrame, max_width: int = 40):
    if df is None or df.empty:
        return
    for col_idx, col in enumerate(df.columns):
        width = len(str(col))
        try:
            sample = df[col].astype(str).head(1000)
            if not sample.empty:
                width = max(width, sample.map(len).max())
        except Exception:
            pass
        ws.set_column(col_idx, col_idx, min(width + 2, max_width))


# ============================================================
# Excel 출력
# ============================================================
def write_excel(out_path: str,
                summary_df: pd.DataFrame,
                top_all: pd.DataFrame,
                top_by_pc: pd.DataFrame,
                by_day: pd.DataFrame,
                by_hour: pd.DataFrame,
                anomalies: pd.DataFrame,
                chart_mode: str = "A") -> None:

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    chart_mode = (chart_mode or "A").strip().upper()
    if chart_mode not in ("A", "B"):
        chart_mode = "A"

    if summary_df.empty and top_all.empty and top_by_pc.empty and by_day.empty and by_hour.empty and anomalies.empty:
        with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
            pd.DataFrame({"message": ["분석 결과 데이터가 없습니다."]}).to_excel(writer, sheet_name="Summary", index=False)
        return

    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        workbook = writer.book

        if summary_df.empty:
            summary_df = pd.DataFrame(columns=["pc_type", "files_count", "total_lines_read", "error_lines_counted"])
        if top_all.empty:
            top_all = pd.DataFrame(columns=["error_code", "error_name", "count"])
        if top_by_pc.empty:
            top_by_pc = pd.DataFrame(columns=["pc_type", "error_code", "error_name", "count"])
        if by_day.empty:
            by_day = pd.DataFrame(columns=["date", "pc_type", "error_code", "error_name", "count"])
        if by_hour.empty:
            by_hour = pd.DataFrame(columns=["hour", "pc_type", "error_code", "error_name", "count"])
        if anomalies.empty:
            anomalies = pd.DataFrame(columns=["pc_type", "date", "error_code", "error_name", "count", "prev_day", "rolling_mean", "reason"])

        summary_df.to_excel(writer, sheet_name="요약", index=False)
        top_all.to_excel(writer, sheet_name="전체 에러 TOP", index=False)
        top_by_pc.to_excel(writer, sheet_name="장치별 에러 TOP", index=False)
        by_day.to_excel(writer, sheet_name="일자별 에러 통계", index=False)
        by_hour.to_excel(writer, sheet_name="시간대별 에러 통계", index=False)
        anomalies.to_excel(writer, sheet_name="SPIKE 경고", index=False)

        autosize_worksheet(writer.sheets["요약"], summary_df)
        autosize_worksheet(writer.sheets["전체 에러 TOP"], top_all)
        autosize_worksheet(writer.sheets["장치별 에러 TOP"], top_by_pc)
        autosize_worksheet(writer.sheets["일자별 에러 통계"], by_day)
        autosize_worksheet(writer.sheets["시간대별 에러 통계"], by_hour)

        # Heatmap 추가
        if not by_hour.empty:
            ws = writer.sheets["시간대별 에러 통계"]
            rows = len(by_hour)
            ws.conditional_format(
                1, 4, rows, 4,
                {
                    "type": "3_color_scale",
                    "min_color": "#FFFFFF",
                    "mid_color": "#FFF2CC",
                    "max_color": "#FF6666",
                }
            )

        autosize_worksheet(writer.sheets["SPIKE 경고"], anomalies, max_width=60)

        ws_chart = workbook.add_worksheet("Charts")

        # ============================================================
        # Device Error List 생성 (장비별 실제 에러 목록)
        # ============================================================
        device_rows = []

        for _, row in top_by_pc.iterrows():

            device = classify_device(row["error_name"], row["pc_type"])

            device_rows.append({
                "device": device,
                "error_name": row["error_name"],
                "count": row["count"]
            })

        device_df = pd.DataFrame(device_rows)

        # 장비별 실제 에러 목록 시트
        device_df.sort_values(
            ["device", "count"],
            ascending=[True, False]
        ).to_excel(
            writer,
            sheet_name="Device_Error_List",
            index=False
        )

        # ============================================================
        # Device Error Summary (차트용)
        # ============================================================
        device_summary = (
            device_df.groupby("device")["count"]
            .sum()
            .reset_index()
        )

        device_summary.to_excel(
            writer,
            sheet_name="Device_Error_Summary",
            index=False
        )

        # ============================================================
        # Device Pie Chart
        # ============================================================
        pie_chart = workbook.add_chart({"type": "pie"})

        rows = len(device_summary)

        pie_chart.add_series({
            "name": "Device Error Ratio",
            "categories": ["Device_Error_Summary", 1, 0, rows, 0],
            "values": ["Device_Error_Summary", 1, 1, rows, 1],
            "data_labels": {"percentage": True},
        })

        pie_chart.set_title({"name": "Device별 Error 비율"})

        ws_chart.insert_chart(1, 14, pie_chart, {"x_scale": 1.2, "y_scale": 1.2})

        top5_pairs = list(
            zip(
                top_all["error_code"].astype(str).head(5),
                top_all["error_name"].head(5)
            )
        )
        if by_day.empty or not top5_pairs:
            ws_chart.write(0, 0, "차트를 생성할 데이터가 부족합니다. (timestamp 있는 로그/Top 코드 부족)")
            return

        chart_src = by_day[
            by_day.apply(
                lambda r: (str(r["error_code"]), r["error_name"]) in top5_pairs,
                axis=1
            )
        ].copy()
        if chart_src.empty:
            ws_chart.write(0, 0, "Top5 코드에 대한 날짜별 데이터가 없습니다.")
            return

        pivot = (
            chart_src.pivot_table(
                index=["date", "pc_type"],
                columns=["error_code","error_name"],
                values="count",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
            .sort_values(["date", "pc_type"])
        )

        new_cols = []
        for c in pivot.columns:
            if isinstance(c, tuple):
                if c[0] in ["date", "pc_type"]:
                    new_cols.append(c[0])
                else:
                    new_cols.append(f"{c[0]} {c[1]}")
            else:
                new_cols.append(c)

        pivot.columns = new_cols

        # 에러코드 + 에러명 표시
        code_name_map = dict(zip(top_all["error_code"].astype(str), top_all["error_name"]))

        pivot.rename(
            columns=lambda c: f"{c} {code_name_map.get(str(c), '')}".strip()
            if c not in ["date", "pc_type"] else c,
            inplace=True
        )

        if pivot.empty:
            ws_chart.write(0, 0, "차트를 생성할 Pivot 데이터가 없습니다.")
            return

        pivot.to_excel(writer, sheet_name="Pivot_Day_Top5_PC", index=False)
        autosize_worksheet(writer.sheets["Pivot_Day_Top5_PC"], pivot)

        # ============================================================
        # PC별 Top5 Pivot 생성
        # ============================================================

        pcs = sorted(by_day["pc_type"].dropna().astype(str).unique().tolist())

        pc_pivot_sheets = {}

        for pc in pcs:

            pc_data = by_day[by_day["pc_type"] == pc]

            if pc_data.empty:
                continue

            pc_top_codes = (
                pc_data.groupby("error_code")["count"]
                .sum()
                .sort_values(ascending=False)
                .head(5)
                .index.astype(str)
                .tolist()
            )

            pc_chart_src = pc_data[pc_data["error_code"].astype(str).isin(pc_top_codes)]

            pc_pivot = (
               pc_chart_src.pivot_table(
                  index="date",
                  columns=["error_code","error_name"],
                  values="count",
                  aggfunc="sum",
                  fill_value=0
               )
               .reset_index()
               .sort_values("date")
            )
            pc_pivot.columns = [
                f"{c[0]} {c[1]}" if isinstance(c, tuple) else c
                for c in pc_pivot.columns
            ]

            # 에러코드 + 에러명 표시
            pc_pivot.rename(
                columns=lambda c: f"{c} {code_name_map.get(str(c), '')}".strip()
                if c != "date" else c,
                inplace=True
            )

            sheet_name = f"Pivot_{pc}_Top5"[:31]

            pc_pivot.to_excel(writer, sheet_name=sheet_name, index=False)

            autosize_worksheet(writer.sheets[sheet_name], pc_pivot)

            pc_pivot_sheets[pc] = pc_pivot

        # ============================================================
        # Charts 생성
        # ============================================================


        # Error Bar Chart
        if not top_all.empty:

            chart_data = top_all.head(20)

            chart_data.to_excel(writer, sheet_name="그래프데이터", index=False)

            rows = len(chart_data)

            bar_chart = workbook.add_chart({"type": "column"})

            bar_chart.add_series({
                "name": "Error Count",
                "categories": ["그래프데이터", 1, 1, rows, 1],
                "values": ["그래프데이터", 1, 2, rows, 2],
                "data_labels": {"value": True},
            })

            bar_chart.set_title({"name": "Error 발생 횟수"})
            bar_chart.set_x_axis({"name": "Error Name"})
            bar_chart.set_y_axis({"name": "Count"})
            bar_chart.set_legend({"none": True})

            ws_chart.insert_chart(1, 0, bar_chart, {"x_scale": 1.3, "y_scale": 1.2})


        # 전체 Top5 Trend Chart 생성 (Pivot_Day_Top5_PC 사용)

        start_row = 24
        start_col = 0

        # 전체 Top5 순서 유지
        top5_order = top_all["error_code"].astype(str).head(5).tolist()

        for code in top5_order:

            # Pivot 컬럼 찾기 (코드 + 이름 포함)
            col_name = next((c for c in pivot.columns if c.startswith(code)), None)

            if not col_name:
                continue

            j = pivot.columns.get_loc(col_name)

            chart = workbook.add_chart({"type": "line"})

            rank = top5_order.index(code) + 1
            error_name = code_name_map.get(code, "")

            chart.set_title({
                "name": f"TOP{rank} {error_name} ({code})"
            })

            # HANDLER 라인
            handler_rows = pivot[pivot["pc_type"] == "HANDLER"]

            if not handler_rows.empty:
                n = len(handler_rows)

                chart.add_series({
                    "name": "HANDLER",
                    "categories": ["Pivot_Day_Top5_PC", 1, 0, n, 0],
                    "values": ["Pivot_Day_Top5_PC", 1, j, n, j],
                })

            # VISION 라인
            vision_rows = pivot[pivot["pc_type"] == "VISION"]

            if not vision_rows.empty:
                start = len(handler_rows) + 1
                end = start + len(vision_rows) - 1

                chart.add_series({
                    "name": "VISION",
                    "categories": ["Pivot_Day_Top5_PC", start, 0, end, 0],
                    "values": ["Pivot_Day_Top5_PC", start, j, end, j],
                })

            chart.set_x_axis({"name": "Date"})
            chart.set_y_axis({"name": "Count"})
            chart.set_legend({"position": "bottom"})

            ws_chart.insert_chart(start_row, start_col, chart, {"x_scale": 1.3, "y_scale": 0.9})

            start_col += 18
            if start_col > 90:
                start_col = 0
                start_row += 18

        # Charts 시트를 맨 앞으로 이동 (루프 완료 후 1회만 실행)
        worksheets = workbook.worksheets_objs
        charts_sheet = next((ws for ws in worksheets if ws.name == "Charts"), None)
        if charts_sheet:
            worksheets.remove(charts_sheet)
            worksheets.insert(0, charts_sheet)


# ============================================================
# GUI
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("980x630")
        self.resizable(False, False)

        self.profiles = load_profiles()
        self.is_running = False

        self.var_equip = tk.StringVar(value=EQUIP_LIST[0])

        self.var_handler = tk.StringVar(value=LOGS_DIR)
        self.var_vision = tk.StringVar(value=LOGS_DIR)
        self.var_output = tk.StringVar(value=os.path.join(OUTPUT_DIR, "log_report.xlsx"))

        self.var_top = tk.StringVar(value="30")
        self.var_spike_factor = tk.StringVar(value="2.0")
        self.var_spike_window = tk.StringVar(value="7")

        self.var_handler_prefix = tk.StringVar(value="Error_")
        self.var_vision_prefix = tk.StringVar(value="Network_")
        self.var_vision_only_error = tk.BooleanVar(value=True)
        self.var_error_codes = tk.StringVar(value="")  # 쉼표 구분 에러코드 필터

        pad = 8
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=pad, pady=pad)

        tab_main = tk.Frame(nb)
        tab_help = tk.Frame(nb)
        nb.add(tab_main, text="분석")
        nb.add(tab_help, text="설명")

        frm = tab_main

        tk.Label(frm, text="설비군 선택:").grid(row=0, column=0, sticky="w")
        self.cmb = ttk.Combobox(frm, textvariable=self.var_equip, values=EQUIP_LIST, width=22, state="readonly")
        self.cmb.grid(row=0, column=1, sticky="w")
        self.cmb.bind("<<ComboboxSelected>>", lambda _e: self.apply_profile())

        tk.Button(frm, text="profiles 위치", command=self.show_profiles_path).grid(row=0, column=2, padx=5, sticky="w")

        tk.Label(frm, text="Handler 루트 폴더:").grid(row=1, column=0, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_handler, width=84).grid(row=1, column=1, sticky="w", pady=(pad, 0))
        tk.Button(frm, text="폴더 선택", command=self.pick_handler).grid(row=1, column=2, padx=5, pady=(pad, 0))

        tk.Label(frm, text="Vision 루트 폴더:").grid(row=2, column=0, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_vision, width=84).grid(row=2, column=1, sticky="w", pady=(pad, 0))
        tk.Button(frm, text="폴더 선택", command=self.pick_vision).grid(row=2, column=2, padx=5, pady=(pad, 0))

        tk.Label(frm, text="출력 Excel(.xlsx):").grid(row=3, column=0, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_output, width=84).grid(row=3, column=1, sticky="w", pady=(pad, 0))
        tk.Button(frm, text="저장 위치", command=self.pick_output).grid(row=3, column=2, padx=5, pady=(pad, 0))

        tk.Label(frm, text="Top N:").grid(row=4, column=1, sticky="w", padx=(220, 0), pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_top, width=6).grid(row=4, column=1, sticky="w", padx=(270, 0), pady=(pad, 0))

        tk.Label(frm, text="Error Code 필터:").grid(row=5, column=0, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_error_codes, width=50).grid(row=5, column=1, sticky="w", pady=(pad, 0))
        tk.Label(frm, text="(쉼표 구분, 비우면 전체)").grid(row=5, column=2, sticky="w", padx=5, pady=(pad, 0))

        tk.Label(frm, text="급증 배수:").grid(row=6, column=0, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_spike_factor, width=8).grid(row=6, column=1, sticky="w", pady=(pad, 0))

        tk.Label(frm, text="급증 기준(이동평균 일수):").grid(row=6, column=1, sticky="w", padx=(220, 0), pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_spike_window, width=6).grid(row=6, column=1, sticky="w", padx=(390, 0), pady=(pad, 0))

        tk.Label(frm, text="파일 Prefix(설비군 적용):").grid(row=8, column=0, sticky="w", pady=(pad, 0))
        tk.Label(frm, text="Handler Prefix").grid(row=8, column=1, sticky="w", pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_handler_prefix, width=14).grid(row=8, column=1, sticky="w", padx=(130, 0), pady=(pad, 0))
        tk.Label(frm, text="Vision Prefix").grid(row=8, column=1, sticky="w", padx=(320, 0), pady=(pad, 0))
        tk.Entry(frm, textvariable=self.var_vision_prefix, width=14).grid(row=8, column=1, sticky="w", padx=(410, 0), pady=(pad, 0))

        tk.Checkbutton(frm, text="Vision은 Error 라인만 집계", variable=self.var_vision_only_error).grid(
            row=9, column=1, sticky="w", pady=(pad, 0)
        )

        tk.Button(frm, text="리포트 생성", command=self.run).grid(row=10, column=1, sticky="w", pady=(18, 0))
        tk.Button(frm, text="종료", command=self.destroy).grid(row=10, column=1, sticky="w", padx=(120, 0), pady=(18, 0))

        self.progress = ttk.Progressbar(frm, orient="horizontal", length=420, mode="determinate")
        self.progress.grid(row=11, column=1, sticky="w", pady=(10, 0))

        self.lbl_status = tk.Label(frm, text="대기")
        self.lbl_status.grid(row=12, column=1, sticky="w", pady=(6, 0))

        help_text = (
            f"{APP_NAME} v{APP_VERSION}\n\n"
            "1) 설비군 선택\n"
            " - profiles 폴더의 해당 JSON이 자동 적용됩니다.\n\n"
            "2) Handler/Vision 루트 폴더\n"
            " - 하위 폴더 포함 재귀 탐색합니다.\n"
            " - Prefix로 시작하는 .txt/.log 파일만 읽습니다.\n\n"
            "3) 출력 Excel\n"
            " - 결과 리포트 저장 위치입니다.\n\n"
            "4) 인코딩\n"
            " - 읽기 실패 시 utf-8-sig / cp949 / euc-kr / utf-16 / latin1 순으로 자동 시도합니다.\n\n"
            "5) Error Code 필터\n"
            " - 특정 코드만 분석할 때 사용합니다.\n\n"
            "6) 디버그 파일\n"
            f" - debug_log.txt: {DEBUG_LOG}\n"
            f" - debug_samples.txt: {DEBUG_SAMPLES}\n\n"
            "7) 무단배포시 두들겨 맞을 수 있습니다."
        )

        txt = tk.Text(tab_help, wrap="word")
        txt.insert("1.0", help_text)
        txt.configure(state="disabled")

        scroll = tk.Scrollbar(tab_help, command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)

        scroll.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        self.apply_profile()

    def set_status(self, text: str):
        self.after(0, lambda: self.lbl_status.config(text=text))

    def set_progress(self, value: float):
        def _update():
            self.progress["value"] = max(0, min(100, value))
        self.after(0, _update)

    def show_profiles_path(self):
        messagebox.showinfo(
            "profiles 위치",
            f"프로파일 폴더 위치:\n{PROFILES_DIR}\n\n설비별 JSON 파일을 수정하면 됩니다."
        )

    def apply_profile(self):
        eq = (self.var_equip.get() or "").strip()
        prof = self.profiles.get(eq)
        if not prof:
            return

        self.var_handler_prefix.set(prof.get("handler_prefix", "Error_"))
        self.var_vision_prefix.set(prof.get("vision_prefix", "Network_"))
        self.var_vision_only_error.set(bool(prof.get("vision_only_error", True)))
        self.var_top.set(str(prof.get("top_n", 30)))

        tpl = prof.get("output_name", "{equip}_{date}_log_report.xlsx")
        fname = render_output_name(tpl, eq)

        cur = self.var_output.get().strip()
        if (not cur) or cur.lower().endswith("log_report.xlsx") or os.path.basename(cur).lower().endswith("_log_report.xlsx"):
            self.var_output.set(os.path.join(OUTPUT_DIR, fname))

    def pick_handler(self):
        p = filedialog.askdirectory(title="Handler 루트 폴더 선택")
        if p:
            self.var_handler.set(p)

    def pick_vision(self):
        p = filedialog.askdirectory(title="Vision 루트 폴더 선택")
        if p:
            self.var_vision.set(p)

    def pick_output(self):
        p = filedialog.asksaveasfilename(
            title="저장할 Excel 파일 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=OUTPUT_DIR,
        )
        if p:
            self.var_output.set(p)

    def run(self):
        if self.is_running:
            messagebox.showwarning("안내", "현재 분석이 진행 중입니다.")
            return

        t = threading.Thread(target=self._run_worker, daemon=True)
        t.start()

    def _run_worker(self):
        self.is_running = True
        self.set_progress(0)
        self.set_status("분석 준비 중")

        try:
            clear_debug_files()

            eq = (self.var_equip.get() or "").strip()
            prof = self.profiles.get(eq)
            if not prof:
                self.after(0, lambda: messagebox.showerror("오류", f"설비군 프로파일을 찾지 못했습니다: {eq}"))
                return

            handler_root = self.var_handler.get().strip()
            vision_root = self.var_vision.get().strip()
            if not handler_root and not vision_root:
                self.after(0, lambda: messagebox.showerror("오류", "Handler 또는 Vision 루트 폴더를 최소 1개 이상 선택해 주세요."))
                return

            encoding = "auto"
            out = self.var_output.get().strip()
            if not out:
                out = os.path.join(OUTPUT_DIR, render_output_name(prof.get("output_name", "{equip}_{date}_log_report.xlsx"), eq))
                self.var_output.set(out)

            try:
                top_n = int(self.var_top.get().strip() or "30")
                if top_n <= 0:
                    raise ValueError
            except Exception:
                self.after(0, lambda: messagebox.showerror("오류", "Top N은 1 이상의 정수여야 합니다."))
                return

            # Error Code 필터 파싱 (쉼표 구분, 빈 값이면 None = 전체 포함)
            raw_codes = self.var_error_codes.get().strip()
            if raw_codes:
                only_codes_set = {c.strip() for c in raw_codes.split(",") if c.strip()}
            else:
                only_codes_set = None

            try:
                spike_factor = float(self.var_spike_factor.get().strip() or "2.0")
                if spike_factor <= 0:
                    raise ValueError
            except Exception:
                self.after(0, lambda: messagebox.showerror("오류", "급증 배수는 0보다 큰 숫자여야 합니다."))
                return

            try:
                spike_window = int(self.var_spike_window.get().strip() or "7")
                if spike_window <= 0:
                    raise ValueError
            except Exception:
                self.after(0, lambda: messagebox.showerror("오류", "급증 기준(이동평균 일수)는 1 이상의 정수여야 합니다."))
                return

            # chart_mode: 프로파일에서 읽기 (기본값 "A")
            chart_mode = (prof.get("chart_mode") or "A").strip().upper()
            if chart_mode not in ("A", "B"):
                chart_mode = "A"

            handler_prefix = (self.var_handler_prefix.get() or "").strip()
            vision_prefix = (self.var_vision_prefix.get() or "").strip()
            vision_only_error = bool(self.var_vision_only_error.get())

            write_debug("===== ANALYSIS START =====")
            write_debug(f"app={APP_NAME} v{APP_VERSION}")
            write_debug(f"equipment={eq}")
            write_debug(f"handler_root={handler_root}")
            write_debug(f"vision_root={vision_root}")
            write_debug(f"encoding={encoding}")
            write_debug(f"output={out}")

            hp = (prof.get("handler_parser") or {})
            vp = (prof.get("vision_parser") or {})

            handler_parser = RegexParser(
                patterns=hp.get("patterns", []),
                dt_formats=hp.get("dt_formats", ["%Y-%m-%d %H:%M:%S"]),
                parser_name=f"{eq}-HANDLER"
            )
            vision_parser = RegexParser(
                patterns=vp.get("patterns", []),
                dt_formats=vp.get("dt_formats", ["%Y-%m-%d %H:%M:%S"]),
                parser_name=f"{eq}-VISION"
            )

            all_infos = []
            all_counters = []

            total_file_count = 0
            handler_files = []
            vision_files = []

            if handler_root:
                handler_files = iter_files(handler_root, prefix=handler_prefix)

                # Prefix로 못 찾으면 전체 로그 다시 검색
                if not handler_files:
                    handler_files = iter_files(handler_root, prefix="")

                total_file_count += len(handler_files)

            if vision_root:
                vision_files = iter_files(vision_root, prefix=vision_prefix)

                # Prefix로 못 찾으면 전체 로그 재검색 (Handler와 동일 fallback)
                if not vision_files:
                    vision_files = iter_files(vision_root, prefix="")

                total_file_count += len(vision_files)

            processed_files = 0

            def progress_cb(pc_type, idx, total, fp):
                nonlocal processed_files, total_file_count
                current = processed_files + idx
                percent = (current / total_file_count) * 100 if total_file_count > 0 else 0
                self.set_progress(percent)
                self.set_status(f"{pc_type} 분석 중... ({current}/{total_file_count}) {os.path.basename(fp)}")

            if handler_files:
                info_h, counters_h = aggregate_logs(
                    files=handler_files,
                    pc_type="HANDLER",
                    encoding=encoding,
                    only_error_lines=True,
                    only_codes_set=only_codes_set,
                    parser=handler_parser,
                    progress_cb=progress_cb,
                )
                info_h["equipment"] = eq
                all_infos.append(info_h)
                all_counters.append(counters_h)
                processed_files += len(handler_files)

            if vision_files:
                info_v, counters_v = aggregate_logs(
                    files=vision_files,
                    pc_type="VISION",
                    encoding=encoding,
                    only_error_lines=vision_only_error,
                    only_codes_set=only_codes_set,
                    parser=vision_parser,
                    progress_cb=progress_cb,
                )
                info_v["equipment"] = eq
                all_infos.append(info_v)
                all_counters.append(counters_v)
                processed_files += len(vision_files)

            self.set_status("집계 데이터 생성 중")
            self.set_progress(92)

            summary_df, top_all, top_by_pc, by_day, by_hour, anomalies = build_dfs(
                all_infos=all_infos,
                all_counters=all_counters,
                top_n=top_n,
                spike_factor=spike_factor,
                spike_window=spike_window,
            )

            self.set_status("Excel 작성 중")
            self.set_progress(96)

            write_excel(
                out_path=out,
                summary_df=summary_df,
                top_all=top_all,
                top_by_pc=top_by_pc,
                by_day=by_day,
                by_hour=by_hour,
                anomalies=anomalies,
                chart_mode=chart_mode,
            )

            self.set_progress(100)
            self.set_status("완료")

            write_debug("===== ANALYSIS END =====")
            self.after(
                0,
                lambda: messagebox.showinfo(
                    "완료",
                    f"리포트 생성 완료:\n설비군: {eq}\n파일: {out}\n\n문제 추적 파일:\n{DEBUG_LOG}"
                )
            )

        except Exception as e:
            err_msg = str(e)
            write_debug(f"FATAL ERROR: {err_msg}\n{traceback.format_exc()}")
            self.set_status("실패")
            self.after(
                0,
                lambda: messagebox.showerror(
                    "실패",
                    f"에러:\n{err_msg}\n\n자세한 내용은 아래 파일을 확인하세요.\n{DEBUG_LOG}"
                )
            )
        finally:
            self.is_running = False


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
