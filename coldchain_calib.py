import csv
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
import argparse
from collections import defaultdict

import pandas as pd


DATA_DIR = Path(os.environ.get("COLDCHAIN_DATA", Path.cwd() / "data"))
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
DB_DIR = DATA_DIR / "db"


CATEGORY_DEFAULTS = {
    "冷冻": {"tolerance": 1.0, "critical": 2.0},
    "冷藏": {"tolerance": 0.8, "critical": 1.5},
    "恒温": {"tolerance": 0.5, "critical": 1.0},
    "果蔬": {"tolerance": 1.2, "critical": 2.0},
    "通用": {"tolerance": 1.0, "critical": 1.5},
}

STATUS_PASS = "合格"
STATUS_WARN = "临近超差"
STATUS_FAIL = "严重超差"

GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
CYAN = "\033[96m"
BOLD = "\033[1m"
RESET = "\033[0m"


def ensure_dirs():
    for d in (INPUT_DIR, OUTPUT_DIR, DB_DIR):
        d.mkdir(parents=True, exist_ok=True)


def cprint(msg, color="", bold=False):
    prefix = f"{BOLD if bold else ''}{color}"
    print(f"{prefix}{msg}{RESET}")


def find_input_file(patterns, desc):
    for pat in patterns:
        for f in sorted(INPUT_DIR.glob(pat)):
            return f
    cprint(f"[错误] 未找到{desc}文件。请将文件放入: {INPUT_DIR}", RED, bold=True)
    cprint(f"       支持的文件名模式: {', '.join(patterns)}", RED)
    return None


def read_table(filepath):
    if filepath.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(filepath, dtype=object)
    return pd.read_csv(filepath, dtype=object, encoding="utf-8-sig")


def norm_col(name):
    return str(name).strip().lower().replace(" ", "").replace("_", "")


def find_col(df, candidates):
    cols = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = norm_col(cand)
        if key in cols:
            return cols[key]
    return None


def parse_date(val):
    if pd.isna(val) or val == "":
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
        "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y.%m.%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(s).to_pydatetime()
    except Exception:
        return None


def parse_float(val):
    if pd.isna(val) or val == "":
        return None
    try:
        return float(str(val).strip().replace(",", ""))
    except (ValueError, TypeError):
        return None


def cmd_import(args):
    ensure_dirs()
    cprint("=" * 60, CYAN, bold=True)
    cprint("  冷链校准核对工具 - 数据导入", CYAN, bold=True)
    cprint("=" * 60, CYAN)
    print(f"数据目录: {DATA_DIR}")
    print()

    veh_file = find_input_file(
        ["vehicle*.*", "车辆*.*", "car*.*", "清单*.*"], "车辆清单"
    )
    std_file = find_input_file(
        ["standard*.*", "标准*.*", "温度计*.*", "ref*.*"], "标准温度计读数"
    )
    probe_file = find_input_file(
        ["probe*.*", "探头*.*", "车载*.*", "reading*.*"], "车载探头读数"
    )
    if not (veh_file and std_file and probe_file):
        return 1

    cprint(f"[读取] 车辆清单: {veh_file.name}", CYAN)
    df_veh = read_table(veh_file)
    cprint(f"[读取] 标准读数: {std_file.name}", CYAN)
    df_std = read_table(std_file)
    cprint(f"[读取] 探头读数: {probe_file.name}", CYAN)
    df_probe = read_table(probe_file)
    print()

    veh_plate = find_col(df_veh, ["车牌", "车牌号", "plate", "license"])
    veh_pos = find_col(df_veh, ["探头位置", "位置", "车厢位置", "position", "location", "pos"])
    veh_last = find_col(df_veh, ["上次校准时间", "校准时间", "最后校准", "last_calibration", "last_calib"])
    veh_cycle = find_col(df_veh, ["校准周期", "周期(天)", "校准周期(天)", "cycle", "interval"])
    veh_car = find_col(df_veh, ["车厢号", "车厢", "厢号", "compartment", "carriage"])

    issues_veh = []
    if not veh_plate:
        issues_veh.append("缺少【车牌】列")
    if not veh_pos:
        issues_veh.append("缺少【探头位置】列")
    if not veh_last:
        issues_veh.append("缺少【上次校准时间】列")
    if issues_veh:
        cprint(f"[错误] 车辆清单字段问题: {'; '.join(issues_veh)}", RED, bold=True)
        return 1

    std_plate = find_col(df_std, ["车牌", "车牌号", "plate", "license"])
    std_pos = find_col(df_std, ["探头位置", "位置", "position", "location", "pos"])
    std_time = find_col(df_std, ["校准时间", "时间", "校准日期", "calibration_time", "time", "date"])
    std_temp = find_col(df_std, ["标准温度", "温度", "读数", "standard_temp", "temperature", "temp"])
    std_car = find_col(df_std, ["车厢号", "车厢", "厢号", "compartment", "carriage"])

    issues_std = []
    if not std_plate:
        issues_std.append("缺少【车牌】列")
    if not std_pos:
        issues_std.append("缺少【探头位置】列")
    if not std_time:
        issues_std.append("缺少【校准时间】列")
    if not std_temp:
        issues_std.append("缺少【标准温度】列")
    if issues_std:
        cprint(f"[错误] 标准读数字段问题: {'; '.join(issues_std)}", RED, bold=True)
        return 1

    probe_plate = find_col(df_probe, ["车牌", "车牌号", "plate", "license"])
    probe_pos = find_col(df_probe, ["探头位置", "位置", "position", "location", "pos"])
    probe_time = find_col(df_probe, ["校准时间", "时间", "校准日期", "calibration_time", "time", "date"])
    probe_temp = find_col(df_probe, ["探头读数", "读数", "车载温度", "温度", "probe_temp", "reading", "temperature", "temp"])
    probe_car = find_col(df_probe, ["车厢号", "车厢", "厢号", "compartment", "carriage"])

    issues_probe = []
    if not probe_plate:
        issues_probe.append("缺少【车牌】列")
    if not probe_pos:
        issues_probe.append("缺少【探头位置】列")
    if not probe_time:
        issues_probe.append("缺少【校准时间】列")
    if not probe_temp:
        issues_probe.append("缺少【探头读数】列")
    if issues_probe:
        cprint(f"[错误] 探头读数字段问题: {'; '.join(issues_probe)}", RED, bold=True)
        return 1

    cprint("--- 车辆清单检查 ---", BOLD)
    vehicles = []
    veh_missing = []
    for idx, row in df_veh.iterrows():
        plate = str(row[veh_plate]).strip() if not pd.isna(row[veh_plate]) else ""
        pos = str(row[veh_pos]).strip() if not pd.isna(row[veh_pos]) else ""
        last = parse_date(row[veh_last]) if veh_last else None
        cycle_days = parse_float(row[veh_cycle]) if veh_cycle else None
        car = str(row[veh_car]).strip() if (veh_car and not pd.isna(row[veh_car])) else "主厢"

        missing = []
        if not plate:
            missing.append("车牌")
        if not pos:
            missing.append("探头位置")
        if veh_last and not last:
            missing.append("校准时间")
        if missing:
            veh_missing.append((idx + 2, plate or "空", "、".join(missing)))
            continue

        vehicles.append({
            "plate": plate,
            "position": pos,
            "carriage": car,
            "last_calibration": last,
            "cycle_days": cycle_days if cycle_days else 30,
        })

    if veh_missing:
        cprint(f"  ⚠  发现 {len(veh_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<12}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in veh_missing[:20]:
            cprint(f"  {lineno:<8}{plate:<12}{miss}", YELLOW)
        if len(veh_missing) > 20:
            cprint(f"  ... 其余 {len(veh_missing) - 20} 行省略", YELLOW)
        print()

    cprint(f"  ✓ 车辆清单有效记录: {len(vehicles)} 条", GREEN)
    print()

    cprint("--- 标准温度计读数检查 ---", BOLD)
    standards = []
    std_missing = []
    for idx, row in df_std.iterrows():
        plate = str(row[std_plate]).strip() if not pd.isna(row[std_plate]) else ""
        pos = str(row[std_pos]).strip() if not pd.isna(row[std_pos]) else ""
        ts = parse_date(row[std_time])
        temp = parse_float(row[std_temp])
        car = str(row[std_car]).strip() if (std_car and not pd.isna(row[std_car])) else "主厢"

        missing = []
        if not plate:
            missing.append("车牌")
        if not pos:
            missing.append("探头位置")
        if not ts:
            missing.append("校准时间")
        if temp is None:
            missing.append("标准温度")
        if missing:
            std_missing.append((idx + 2, plate or "空", "、".join(missing)))
            continue

        standards.append({
            "plate": plate,
            "position": pos,
            "carriage": car,
            "calibration_time": ts,
            "standard_temp": temp,
        })

    if std_missing:
        cprint(f"  ⚠  发现 {len(std_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<12}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in std_missing[:20]:
            cprint(f"  {lineno:<8}{plate:<12}{miss}", YELLOW)
        if len(std_missing) > 20:
            cprint(f"  ... 其余 {len(std_missing) - 20} 行省略", YELLOW)
        print()

    cprint(f"  ✓ 标准读数有效记录: {len(standards)} 条", GREEN)
    print()

    cprint("--- 车载探头读数检查 ---", BOLD)
    probes = []
    probe_missing = []
    for idx, row in df_probe.iterrows():
        plate = str(row[probe_plate]).strip() if not pd.isna(row[probe_plate]) else ""
        pos = str(row[probe_pos]).strip() if not pd.isna(row[probe_pos]) else ""
        ts = parse_date(row[probe_time])
        temp = parse_float(row[probe_temp])
        car = str(row[probe_car]).strip() if (probe_car and not pd.isna(row[probe_car])) else "主厢"

        missing = []
        if not plate:
            missing.append("车牌")
        if not pos:
            missing.append("探头位置")
        if not ts:
            missing.append("校准时间")
        if temp is None:
            missing.append("探头读数")
        if missing:
            probe_missing.append((idx + 2, plate or "空", "、".join(missing)))
            continue

        probes.append({
            "plate": plate,
            "position": pos,
            "carriage": car,
            "calibration_time": ts,
            "probe_temp": temp,
        })

    if probe_missing:
        cprint(f"  ⚠  发现 {len(probe_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<12}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in probe_missing[:20]:
            cprint(f"  {lineno:<8}{plate:<12}{miss}", YELLOW)
        if len(probe_missing) > 20:
            cprint(f"  ... 其余 {len(probe_missing) - 20} 行省略", YELLOW)
        print()

    cprint(f"  ✓ 探头读数有效记录: {len(probes)} 条", GREEN)
    print()

    df_vehicles = pd.DataFrame(vehicles)
    df_standards = pd.DataFrame(standards)
    df_probes = pd.DataFrame(probes)

    df_vehicles.to_pickle(DB_DIR / "vehicles.pkl")
    df_standards.to_pickle(DB_DIR / "standards.pkl")
    df_probes.to_pickle(DB_DIR / "probes.pkl")

    csv_dir = OUTPUT_DIR / "imported"
    csv_dir.mkdir(parents=True, exist_ok=True)
    df_vehicles.to_csv(csv_dir / "vehicles.csv", index=False, encoding="utf-8-sig")
    df_standards.to_csv(csv_dir / "standards.csv", index=False, encoding="utf-8-sig")
    df_probes.to_csv(csv_dir / "probes.csv", index=False, encoding="utf-8-sig")

    cprint("=" * 60, GREEN)
    cprint(f"  ✓ 导入完成！数据已保存至: {DB_DIR}", GREEN, bold=True)
    cprint(f"  共 {len(vehicles)} 辆车, {len(standards)} 条标准, {len(probes)} 条探头记录", GREEN)
    cprint("=" * 60, GREEN)

    total_issues = len(veh_missing) + len(std_missing) + len(probe_missing)
    if total_issues > 0:
        cprint(f"  ⚠  共 {total_issues} 行数据存在问题，请检查后重新导入。", YELLOW, bold=True)
    return 0


def _load_db():
    if not (DB_DIR / "vehicles.pkl").exists():
        cprint("[错误] 数据未导入，请先执行 import 命令", RED, bold=True)
        return None, None, None
    df_v = pd.read_pickle(DB_DIR / "vehicles.pkl")
    df_s = pd.read_pickle(DB_DIR / "standards.pkl")
    df_p = pd.read_pickle(DB_DIR / "probes.pkl")
    return df_v, df_s, df_p


def _build_pairs(df_standards, df_probes):
    pairs = []
    std_grouped = defaultdict(list)
    for _, s in df_standards.iterrows():
        key = (s["plate"], s["carriage"], s["position"])
        std_grouped[key].append(s)

    probe_grouped = defaultdict(list)
    for _, p in df_probes.iterrows():
        key = (p["plate"], p["carriage"], p["position"])
        probe_grouped[key].append(p)

    used_probe_idx = set()

    for key, std_list in std_grouped.items():
        probe_list = probe_grouped.get(key, [])
        for s in std_list:
            best_p = None
            best_delta = None
            best_p_idx = None
            for pi, p in enumerate(probe_list):
                if id(p) in used_probe_idx:
                    continue
                delta = abs((s["calibration_time"] - p["calibration_time"]).total_seconds())
                if best_delta is None or delta < best_delta:
                    if delta <= 3600:
                        best_delta = delta
                        best_p = p
                        best_p_idx = pi
            if best_p is not None:
                used_probe_idx.add(id(probe_list[best_p_idx]))
                diff = best_p["probe_temp"] - s["standard_temp"]
                pairs.append({
                    "plate": s["plate"],
                    "carriage": s["carriage"],
                    "position": s["position"],
                    "calibration_time": s["calibration_time"],
                    "standard_temp": s["standard_temp"],
                    "probe_temp": best_p["probe_temp"],
                    "deviation": round(diff, 3),
                    "abs_deviation": round(abs(diff), 3),
                })

    return pd.DataFrame(pairs)


def cmd_check(args):
    ensure_dirs()
    df_v, df_s, df_p = _load_db()
    if df_v is None:
        return 1

    tolerance = args.tolerance
    critical = args.critical
    category = args.category
    if category and category in CATEGORY_DEFAULTS and tolerance is None:
        tolerance = CATEGORY_DEFAULTS[category]["tolerance"]
        critical = CATEGORY_DEFAULTS[category]["critical"]

    if tolerance is None:
        tolerance = 1.0
    if critical is None:
        critical = tolerance * 1.5

    cprint("=" * 70, CYAN, bold=True)
    cprint("  冷链校准核对工具 - 校准核对", CYAN, bold=True)
    cprint("=" * 70, CYAN)
    info = f"  允许偏差: ±{tolerance}°C   严重超差阈值: ±{critical}°C"
    if category:
        info += f"   运输品类: {category}"
    cprint(info, CYAN)
    print()

    df_pairs = _build_pairs(df_s, df_p)
    if df_pairs.empty:
        cprint("[错误] 未找到可匹配的标准-探头记录对（需同一车牌、车厢、位置、时间差≤1小时）", RED, bold=True)
        return 1

    def classify(row):
        d = row["abs_deviation"]
        if d > critical:
            return STATUS_FAIL
        elif d > tolerance:
            return STATUS_WARN
        return STATUS_PASS

    df_pairs["status"] = df_pairs.apply(classify, axis=1)

    df_pairs = df_pairs.sort_values(
        ["plate", "carriage", "position", "calibration_time"]
    ).reset_index(drop=True)

    df_pairs["consecutive_high"] = False
    grouped = df_pairs.groupby(["plate", "carriage", "position"])
    for key, grp in grouped:
        if len(grp) >= 2:
            grp_sorted = grp.sort_values("calibration_time")
            for i in range(1, len(grp_sorted)):
                cur = grp_sorted.iloc[i]
                prev = grp_sorted.iloc[i - 1]
                if cur["deviation"] > 0 and prev["deviation"] > 0 and cur["abs_deviation"] > tolerance * 0.5:
                    df_pairs.loc[grp_sorted.index[i], "consecutive_high"] = True

    status_groups = {
        STATUS_PASS: df_pairs[df_pairs["status"] == STATUS_PASS],
        STATUS_WARN: df_pairs[df_pairs["status"] == STATUS_WARN],
        STATUS_FAIL: df_pairs[df_pairs["status"] == STATUS_FAIL],
    }

    consec_high = df_pairs[df_pairs["consecutive_high"] == True]

    df_merged = df_pairs.merge(
        df_v[["plate", "carriage", "position", "cycle_days", "last_calibration"]],
        on=["plate", "carriage", "position"],
        how="left",
    )

    df_merged.to_pickle(DB_DIR / "check_results.pkl")
    df_pairs.to_csv(OUTPUT_DIR / "check_results.csv", index=False, encoding="utf-8-sig")

    def print_group(title, color, df, show_all=False):
        cprint("-" * 70, color, bold=True)
        cprint(f"  {title}（共 {len(df)} 条）", color, bold=True)
        cprint("-" * 70, color)
        if df.empty:
            cprint("  (无记录)", color)
            print()
            return
        header = f"  {'车牌':<10}{'车厢':<8}{'位置':<10}{'校准时间':<20}{'偏差(°C)':>10}"
        cprint(header, color, bold=True)
        print_limit = len(df) if show_all else min(len(df), 15)
        for _, r in df.head(print_limit).iterrows():
            dev = f"{r['deviation']:+.3f}"
            flag = " ↗↗" if r.get("consecutive_high") else ""
            line = f"  {r['plate']:<10}{r['carriage']:<8}{r['position']:<10}{r['calibration_time'].strftime('%Y-%m-%d %H:%M'):<20}{dev:>10}{flag}"
            cprint(line, color)
        if not show_all and len(df) > print_limit:
            cprint(f"  ... 其余 {len(df) - print_limit} 条略，详见 check_results.csv", color)
        print()

    print_group(STATUS_PASS, GREEN, status_groups[STATUS_PASS], show_all=args.verbose)
    print_group(STATUS_WARN, YELLOW, status_groups[STATUS_WARN], show_all=True)
    print_group(STATUS_FAIL, RED, status_groups[STATUS_FAIL], show_all=True)

    if not consec_high.empty:
        cprint("=" * 70, YELLOW, bold=True)
        cprint("  ⚠  同一车厢/位置连续两次偏高警告（建议重点检查）", YELLOW, bold=True)
        cprint("=" * 70, YELLOW)
        header = f"  {'车牌':<10}{'车厢':<8}{'位置':<10}{'最近校准时间':<20}{'本次偏差':>10}"
        cprint(header, YELLOW, bold=True)
        for _, r in consec_high.iterrows():
            dev = f"{r['deviation']:+.3f}°C"
            line = f"  {r['plate']:<10}{r['carriage']:<8}{r['position']:<10}{r['calibration_time'].strftime('%Y-%m-%d %H:%M'):<20}{dev:>10}"
            cprint(line, YELLOW)
        print()

    total = len(df_pairs)
    p_pass = len(status_groups[STATUS_PASS])
    p_warn = len(status_groups[STATUS_WARN])
    p_fail = len(status_groups[STATUS_FAIL])
    cprint("=" * 70, BOLD)
    cprint(f"  核对汇总: 共 {total} 条 | {GREEN}合格 {p_pass}{RESET} | {YELLOW}临近超差 {p_warn}{RESET} | {RED}严重超差 {p_fail}{RESET}", BOLD)
    if p_pass < total:
        rate = p_pass / total * 100
        cprint(f"  合格率: {rate:.1f}%", BOLD)
    cprint(f"  详细结果: {OUTPUT_DIR / 'check_results.csv'}", CYAN)
    cprint("=" * 70, BOLD)
    return 0


def cmd_summary(args):
    ensure_dirs()
    df_v, df_s, df_p = _load_db()
    if df_v is None:
        return 1

    check_path = DB_DIR / "check_results.pkl"
    if not check_path.exists():
        cprint("[错误] 未找到核对结果，请先执行 check 命令", RED, bold=True)
        return 1
    df_check = pd.read_pickle(check_path)

    tolerance = args.tolerance if args.tolerance else 1.0
    critical = args.critical if args.critical else tolerance * 1.5
    cycle_days = args.cycle_days if args.cycle_days else 30
    today = datetime.now()

    cprint("=" * 70, BOLD)
    cprint("          冷链探头校准月度摘要报告（供车队经理审阅）", BOLD)
    cprint("=" * 70, BOLD)
    cprint(f"  生成时间: {today.strftime('%Y-%m-%d %H:%M')}      周期阈值: {cycle_days}天", CYAN)
    print()

    total_vehicles = df_v["plate"].nunique()
    total_probes = len(df_v)
    total_checked = df_check["plate"].count()
    p_pass = (df_check["status"] == STATUS_PASS).sum()
    p_warn = (df_check["status"] == STATUS_WARN).sum()
    p_fail = (df_check["status"] == STATUS_FAIL).sum()
    pass_rate = (p_pass / total_checked * 100) if total_checked else 0

    cprint("【一、总体情况】", BOLD)
    cprint(f"  • 登记车辆总数: {total_vehicles} 台", "")
    cprint(f"  • 登记探头总数: {total_probes} 个", "")
    cprint(f"  • 本月校准核对记录: {total_checked} 条", "")
    cprint(f"  • 合格率: {GREEN}{pass_rate:.1f}%{RESET}  ({GREEN}{p_pass}合格{RESET} / {YELLOW}{p_warn}临近{RESET} / {RED}{p_fail}超差{RESET})", "")
    print()

    cprint("【二、需立即复检车辆】（存在临近超差或严重超差）", BOLD)
    problem_probes = df_check[df_check["status"] != STATUS_PASS].copy()
    problem_vehicles = problem_probes.groupby("plate").agg(
        问题数=("status", "count"),
        严重超差=("status", lambda s: (s == STATUS_FAIL).sum()),
        最严重偏差=("abs_deviation", "max"),
    ).sort_values(["严重超差", "问题数"], ascending=False)

    if problem_vehicles.empty:
        cprint("  ✓ 本月所有校准记录均合格，无需复检。", GREEN)
    else:
        header = f"  {'车牌':<12}{'问题探头数':>10}{'严重超差':>10}{'最大偏差(°C)':>14}"
        cprint(header, BOLD)
        for plate, row in problem_vehicles.iterrows():
            color = RED if row["严重超差"] > 0 else YELLOW
            line = f"  {plate:<12}{int(row['问题数']):>10}{int(row['严重超差']):>10}{row['最严重偏差']:>14.3f}"
            cprint(line, color)
    print()

    cprint("【三、建议停用探头】（严重超差 + 连续两次偏高）", BOLD)
    stop_candidates = df_check[
        (df_check["status"] == STATUS_FAIL) | (df_check["consecutive_high"] == True)
    ].copy()

    if stop_candidates.empty:
        cprint("  ✓ 本月无需要建议停用的探头。", GREEN)
    else:
        stop_candidates["reason"] = stop_candidates.apply(
            lambda r: "严重超差" if r["status"] == STATUS_FAIL else "", axis=1
        )
        stop_candidates["reason"] = stop_candidates.apply(
            lambda r: r["reason"] + ("+连续偏高" if r["consecutive_high"] else "").lstrip("+"),
            axis=1,
        )
        header = f"  {'车牌':<12}{'车厢':<8}{'位置':<10}{'偏差(°C)':>10}  {'问题原因':<20}"
        cprint(header, BOLD)
        for _, r in stop_candidates.iterrows():
            dev = f"{r['deviation']:+.3f}"
            line = f"  {r['plate']:<12}{r['carriage']:<8}{r['position']:<10}{dev:>10}  {r['reason']:<20}"
            cprint(line, RED)
    print()

    cprint(f"【四、即将到期校准任务】（{cycle_days}天周期，未来15天内需校准）", BOLD)
    upcoming = []
    for _, v in df_v.iterrows():
        last = v["last_calibration"]
        cyc = v["cycle_days"] if v["cycle_days"] else cycle_days
        if isinstance(last, datetime):
            next_calib = last + timedelta(days=cyc)
            days_left = (next_calib - today).days
            if days_left <= 15:
                upcoming.append({
                    "plate": v["plate"],
                    "carriage": v["carriage"],
                    "position": v["position"],
                    "last": last,
                    "next": next_calib,
                    "days_left": days_left,
                })

    if not upcoming:
        cprint("  ✓ 未来15天内无即将到期的校准任务。", GREEN)
    else:
        upcoming.sort(key=lambda x: x["days_left"])
        header = f"  {'车牌':<12}{'车厢':<8}{'位置':<10}{'上次校准':<16}{'下次校准':<16}{'剩余天数':>8}"
        cprint(header, BOLD)
        for u in upcoming:
            color = RED if u["days_left"] < 0 else (YELLOW if u["days_left"] <= 7 else CYAN)
            status = "已逾期" if u["days_left"] < 0 else f"{u['days_left']}天"
            line = (
                f"  {u['plate']:<12}{u['carriage']:<8}{u['position']:<10}"
                f"{u['last'].strftime('%Y-%m-%d'):<16}{u['next'].strftime('%Y-%m-%d'):<16}{status:>8}"
            )
            cprint(line, color)
    print()

    cprint("【五、运营建议】", BOLD)
    suggestions = []
    if p_fail > 0:
        suggestions.append(f"• {RED}{p_fail}个探头严重超差{RESET}，建议立即停用并安排专业维修或更换，避免货损风险。")
    if p_warn > 0:
        suggestions.append(f"• {YELLOW}{p_warn}个探头临近超差{RESET}，建议增加抽查频率并提前安排校准。")
    if not stop_candidates.empty:
        suggestions.append(f"• {RED}{len(stop_candidates)}个探头建议停用{RESET}，调度时避免分配给高价值冷链货物。")
    overdue = [u for u in upcoming if u["days_left"] < 0]
    soon = [u for u in upcoming if 0 <= u["days_left"] <= 7]
    if overdue:
        suggestions.append(f"• {RED}{len(overdue)}个探头已逾期校准{RESET}，请立即安排校准并评估期间数据可靠性。")
    elif soon:
        suggestions.append(f"• {YELLOW}{len(soon)}个探头一周内需校准{RESET}，请提前预约校准资源。")
    if pass_rate < 80:
        suggestions.append(f"• 整体合格率仅 {pass_rate:.1f}%，建议排查探头老化或安装位置问题，组织统一专项校准。")
    if not suggestions:
        suggestions.append("• 本月整体运营状况良好，请继续保持现有校准管理节奏。")

    for s in suggestions:
        cprint(f"  {s}", "")
    print()

    report_path = OUTPUT_DIR / f"summary_{today.strftime('%Y%m')}.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("冷链探头校准月度摘要报告\n")
        f.write(f"生成时间: {today.strftime('%Y-%m-%d %H:%M')}\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"一、总体情况\n")
        f.write(f"  登记车辆: {total_vehicles} 台, 探头: {total_probes} 个\n")
        f.write(f"  核对记录: {total_checked} 条, 合格率: {pass_rate:.1f}%\n")
        f.write(f"  合格: {p_pass}, 临近: {p_warn}, 超差: {p_fail}\n\n")

        f.write("二、需复检车辆\n")
        if problem_vehicles.empty:
            f.write("  无\n")
        else:
            for plate, row in problem_vehicles.iterrows():
                f.write(f"  {plate}: 问题{int(row['问题数'])}个, 严重{int(row['严重超差'])}个, 最大偏差{row['最严重偏差']:.3f}°C\n")
        f.write("\n")

        f.write("三、建议停用探头\n")
        if stop_candidates.empty:
            f.write("  无\n")
        else:
            for _, r in stop_candidates.iterrows():
                reason = "严重超差" if r["status"] == STATUS_FAIL else ""
                if r["consecutive_high"]:
                    reason = reason + "+连续偏高" if reason else "连续偏高"
                f.write(f"  {r['plate']} {r['carriage']} {r['position']}: 偏差{r['deviation']:+.3f}°C, 原因: {reason}\n")
        f.write("\n")

        f.write(f"四、即将到期校准任务（{cycle_days}天周期）\n")
        if not upcoming:
            f.write("  无\n")
        else:
            for u in upcoming:
                status = f"逾期{-u['days_left']}天" if u["days_left"] < 0 else f"剩{u['days_left']}天"
                f.write(f"  {u['plate']} {u['carriage']} {u['position']}: 下次{u['next'].strftime('%Y-%m-%d')} ({status})\n")
        f.write("\n")

        f.write("五、运营建议\n")
        for s in suggestions:
            plain = s.replace(RED, "").replace(YELLOW, "").replace(GREEN, "").replace(CYAN, "").replace(RESET, "").replace(BOLD, "")
            f.write(f"  {plain}\n")

    cprint(f"  报告已保存至: {report_path}", CYAN, bold=True)
    cprint("=" * 70, BOLD)
    return 0


def main():
    parser = argparse.ArgumentParser(
        prog="coldchain-calib",
        description="冷链探头校准核对工具 - 适用于冷藏车队月底批量校准数据核对",
        epilog="""
示例:
  coldchain-calib import              # 从 data/input 导入三张表
  coldchain-calib check -c 冷冻       # 按冷冻品类标准核对
  coldchain-calib check -t 1.0 -k 2.0 # 自定义偏差±1°C / 严重超差±2°C
  coldchain-calib summary -d 30       # 按30天周期输出摘要报告
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", help="可用命令")

    p_import = sub.add_parser("import", help="导入车辆清单、标准读数、探头读数数据")

    p_check = sub.add_parser("check", help="执行校准偏差核对")
    p_check.add_argument("-t", "--tolerance", type=float, default=None,
                         help="允许偏差阈值(°C)，默认根据品类或1.0°C")
    p_check.add_argument("-k", "--critical", type=float, default=None,
                         help="严重超差阈值(°C)，默认根据品类或偏差的1.5倍")
    p_check.add_argument("-c", "--category", default=None,
                         choices=list(CATEGORY_DEFAULTS.keys()),
                         help="运输品类(自动匹配阈值)")
    p_check.add_argument("-v", "--verbose", action="store_true",
                         help="完整显示合格记录(默认仅显示前15条)")

    p_summary = sub.add_parser("summary", help="生成车队经理摘要报告")
    p_summary.add_argument("-t", "--tolerance", type=float, default=None,
                           help="允许偏差阈值(用于摘要判断)")
    p_summary.add_argument("-k", "--critical", type=float, default=None,
                           help="严重超差阈值")
    p_summary.add_argument("-d", "--cycle-days", type=int, default=None,
                           help="校准周期(天)，默认30天或车辆清单设定")

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 0

    ensure_dirs()

    if args.command == "import":
        return cmd_import(args)
    elif args.command == "check":
        return cmd_check(args)
    elif args.command == "summary":
        return cmd_summary(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main() or 0)
