import csv
import io
import json
import os
import sys
import zipfile
import shutil
from datetime import datetime, timedelta
from pathlib import Path
import argparse
from collections import defaultdict, OrderedDict

import pandas as pd

if sys.platform.startswith("win"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass


DATA_DIR = Path(os.environ.get("COLDCHAIN_DATA", Path.cwd() / "data"))
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
DB_DIR = DATA_DIR / "db"
BATCH_DIR = DB_DIR / "batches"
ARCHIVE_DIR = DATA_DIR / "archive"
BATCH_INDEX = DB_DIR / "batch_index.json"


CATEGORY_DEFAULTS = {
    "冷冻": {"tolerance": 1.0, "critical": 2.0},
    "冷藏": {"tolerance": 0.8, "critical": 1.5},
    "恒温": {"tolerance": 0.5, "critical": 1.0},
    "果蔬": {"tolerance": 1.2, "critical": 2.0},
    "通用": {"tolerance": 1.0, "critical": 1.5},
}

FIELD_SCHEMA = {
    "vehicles": OrderedDict([
        ("plate",      {"label": "车牌",         "required": True,
                        "candidates": ["车牌", "车牌号", "plate", "license", "车辆牌照"]}),
        ("position",   {"label": "探头位置",      "required": True,
                        "candidates": ["探头位置", "位置", "车厢位置", "安装位置",
                                       "position", "location", "pos", "probe_location"]}),
        ("last_calib", {"label": "上次校准时间",   "required": True,
                        "candidates": ["上次校准时间", "校准时间", "最后校准", "上次校准日期",
                                       "last_calibration", "last_calib", "calibration_date"]}),
        ("cycle_days", {"label": "校准周期",      "required": False,
                        "candidates": ["校准周期", "周期(天)", "校准周期(天)", "校准间隔",
                                       "cycle", "interval", "calibration_cycle"]}),
        ("carriage",   {"label": "车厢号",        "required": False,
                        "candidates": ["车厢号", "车厢", "厢号", "车厢编号",
                                       "compartment", "carriage", "car_no"]}),
    ]),
    "standards": OrderedDict([
        ("plate",      {"label": "车牌",         "required": True,
                        "candidates": ["车牌", "车牌号", "plate", "license", "车辆牌照"]}),
        ("position",   {"label": "探头位置",      "required": True,
                        "candidates": ["探头位置", "位置", "车厢位置", "安装位置",
                                       "position", "location", "pos", "probe_location"]}),
        ("calib_time", {"label": "校准时间",      "required": True,
                        "candidates": ["校准时间", "时间", "校准日期", "记录时间",
                                       "calibration_time", "time", "date", "datetime"]}),
        ("std_temp",   {"label": "标准温度",      "required": True,
                        "candidates": ["标准温度", "温度", "读数", "标准读数", "温度计读数",
                                       "standard_temp", "temperature", "temp", "ref_temp", "reading"]}),
        ("carriage",   {"label": "车厢号",        "required": False,
                        "candidates": ["车厢号", "车厢", "厢号", "车厢编号",
                                       "compartment", "carriage", "car_no"]}),
    ]),
    "probes": OrderedDict([
        ("plate",      {"label": "车牌",         "required": True,
                        "candidates": ["车牌", "车牌号", "plate", "license", "车辆牌照"]}),
        ("position",   {"label": "探头位置",      "required": True,
                        "candidates": ["探头位置", "位置", "车厢位置", "安装位置",
                                       "position", "location", "pos", "probe_location"]}),
        ("calib_time", {"label": "校准时间",      "required": True,
                        "candidates": ["校准时间", "时间", "校准日期", "记录时间",
                                       "calibration_time", "time", "date", "datetime"]}),
        ("probe_temp", {"label": "探头读数",      "required": True,
                        "candidates": ["探头读数", "读数", "车载温度", "温度", "车载探头",
                                       "probe_temp", "reading", "temperature", "temp", "probe_reading"]}),
        ("carriage",   {"label": "车厢号",        "required": False,
                        "candidates": ["车厢号", "车厢", "厢号", "车厢编号",
                                       "compartment", "carriage", "car_no"]}),
    ]),
}

FILE_PATTERNS = {
    "vehicles": (["vehicle*.*", "车辆*.*", "car*.*", "清单*.*", "fleet*.*"], "车辆清单"),
    "standards": (["standard*.*", "标准*.*", "温度计*.*", "ref*.*", "reference*.*"], "标准温度计读数"),
    "probes":    (["probe*.*", "探头*.*", "车载*.*", "reading*.*", "sensor*.*"], "车载探头读数"),
}

STATUS_PASS = "合格"
STATUS_WARN = "临近超差"
STATUS_FAIL = "严重超差"

GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
CYAN = "\033[96m"
MAGENTA = "\033[95m"
BOLD = "\033[1m"
DIM = "\033[2m"
RESET = "\033[0m"


def ensure_dirs():
    for d in (INPUT_DIR, OUTPUT_DIR, DB_DIR, BATCH_DIR, ARCHIVE_DIR):
        d.mkdir(parents=True, exist_ok=True)


def cprint(msg, color="", bold=False):
    prefix = f"{BOLD if bold else ''}{color}"
    print(f"{prefix}{msg}{RESET}")


def find_input_file(patterns, desc):
    for pat in patterns:
        for f in sorted(INPUT_DIR.glob(pat)):
            if f.name.startswith("~$"):
                continue
            return f
    return None


def read_table(filepath):
    if filepath.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(filepath, dtype=object)
    return pd.read_csv(filepath, dtype=object, encoding="utf-8-sig")


def norm_col(name):
    return str(name).strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def find_col(df, candidates):
    cols = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = norm_col(cand)
        if key in cols:
            return cols[key]
    for cand in candidates:
        key = norm_col(cand)
        for col_key, col_name in cols.items():
            if key in col_key or col_key in key:
                return col_name
    return None


def parse_date(val):
    if pd.isna(val) or val == "" or val is None:
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
        "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y.%m.%d",
        "%Y%m%d", "%Y%m%d %H%M%S", "%Y%m%d%H%M%S",
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
    if pd.isna(val) or val == "" or val is None:
        return None
    try:
        return float(str(val).strip().replace(",", "").replace("°C", "").replace("℃", ""))
    except (ValueError, TypeError):
        return None


def parse_int(val):
    if pd.isna(val) or val == "" or val is None:
        return None
    try:
        x = float(str(val).strip())
        return int(round(x))
    except (ValueError, TypeError):
        return None


def gen_batch_id():
    return "BAT-" + datetime.now().strftime("%Y%m%d-%H%M%S")


def load_batch_index():
    if BATCH_INDEX.exists():
        try:
            with open(BATCH_INDEX, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_batch_index(idx):
    with open(BATCH_INDEX, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=2)


def register_batch(batch_id, meta):
    idx = load_batch_index()
    entry = {"batch_id": batch_id, **meta, "created_at": datetime.now().isoformat()}
    idx.insert(0, entry)
    save_batch_index(idx)


def list_batches(limit=10):
    idx = load_batch_index()
    return idx[:limit]


def resolve_batch(requested=None):
    idx = load_batch_index()
    if not idx:
        return None, None
    if requested:
        for e in idx:
            if e["batch_id"] == requested or e["batch_id"].endswith(requested):
                return e["batch_id"], e
        cprint(f"[错误] 未找到批次: {requested}", RED, bold=True)
        cprint(f"  可用批次: {', '.join(e['batch_id'] for e in idx[:5])}", RED)
        return None, None
    return idx[0]["batch_id"], idx[0]


def batch_path(batch_id):
    return BATCH_DIR / batch_id


def load_batch(batch_id):
    bp = batch_path(batch_id)
    if not bp.exists():
        return None, None, None
    v = pd.read_pickle(bp / "vehicles.pkl") if (bp / "vehicles.pkl").exists() else None
    s = pd.read_pickle(bp / "standards.pkl") if (bp / "standards.pkl").exists() else None
    p = pd.read_pickle(bp / "probes.pkl") if (bp / "probes.pkl").exists() else None
    return v, s, p


def has_checked(batch_id):
    bp = batch_path(batch_id)
    return bp.exists() and (bp / "check_results.pkl").exists()


def load_check_results(batch_id):
    bp = batch_path(batch_id)
    if not bp.exists() or not (bp / "check_results.pkl").exists():
        return None
    return pd.read_pickle(bp / "check_results.pkl")


def cal_status(last, cyc, today):
    if not isinstance(last, datetime):
        return "未知", "-", ""
    next_c = last + timedelta(days=cyc)
    days_left = (next_c - today).days
    if days_left < 0:
        label = f"逾期{-days_left}天"
        tag = "[逾期]"
        color = RED
    elif days_left <= 7:
        label = f"剩{days_left}天"
        tag = "[临期]"
        color = YELLOW
    else:
        label = f"剩{days_left}天"
        tag = f"·{label}"
        color = ""
    return tag, label, color


def cmd_import(args):
    ensure_dirs()
    batch_id = gen_batch_id()
    bp = batch_path(batch_id)
    bp.mkdir(parents=True, exist_ok=True)

    cprint("=" * 72, CYAN, bold=True)
    cprint(f"  冷链校准核对工具 - 数据导入   [批次 {batch_id}]", CYAN, bold=True)
    cprint("=" * 72, CYAN)
    print(f"数据目录: {DATA_DIR}")
    print()

    files = {}
    missing_files = []
    for key, (patterns, desc) in FILE_PATTERNS.items():
        f = find_input_file(patterns, desc)
        if f:
            files[key] = f
            cprint(f"[读取] {desc}: {f.name}", CYAN)
        else:
            missing_files.append((desc, patterns))

    if missing_files:
        for desc, patterns in missing_files:
            cprint(f"[错误] 未找到{desc}文件", RED, bold=True)
            cprint(f"       请放入 {INPUT_DIR}，文件名需匹配: {', '.join(patterns)}", RED)
        return 1

    dfs = {k: read_table(v) for k, v in files.items()}
    print()

    field_report_lines = []
    field_report_lines.append("=" * 70)
    field_report_lines.append(f"冷链校准 - 字段识别报告    批次: {batch_id}")
    field_report_lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    field_report_lines.append("=" * 70)
    field_report_lines.append("")

    all_matched_cols = defaultdict(set)
    mapping = {}
    fatal_errors = []

    for fkey in ("vehicles", "standards", "probes"):
        fdesc = FILE_PATTERNS[fkey][1]
        fname = files[fkey].name
        df = dfs[fkey]
        schema = FIELD_SCHEMA[fkey]

        cprint("─" * 72, MAGENTA, bold=True)
        cprint(f"  [字段识别] {fdesc}", MAGENTA, bold=True)
        cprint("─" * 72, MAGENTA)
        cprint(f"  源文件: {fname}  |  共 {len(df)} 行 x {len(df.columns)} 列", DIM)
        print()

        field_report_lines.append(f"【{fdesc}】 源文件: {fname}")
        field_report_lines.append(f"  原始列名清单: {', '.join(str(c) for c in df.columns)}")
        field_report_lines.append("")

        matched_map = {}
        matched_cols_in_df = set()
        missing_required = []

        cprint(f"  {'逻辑字段':<16}{'匹配到的列名':<24}{'状态':<10}{'说明'}", BOLD)
        field_report_lines.append(f"  {'逻辑字段':<16}{'匹配到的列名':<24}{'状态':<10}{'候选别名'}")
        field_report_lines.append("  " + "-" * 78)

        for field_key, spec in schema.items():
            col = find_col(df, spec["candidates"])
            matched_map[field_key] = col

            if col:
                matched_cols_in_df.add(col)
                all_matched_cols[fkey].add(col)
                status = "V"
                status_c = GREEN
                desc_txt = f"识别为「{col}」"
            else:
                status = "X" if spec["required"] else "—"
                status_c = RED if spec["required"] else DIM
                desc_txt = "未匹配" + ("（必填，请修正表头）" if spec["required"] else "（可选，将用默认值）")
                if spec["required"]:
                    missing_required.append(spec["label"])

            cprint(f"  {spec['label']:<16}{(col or ''):<24}", status_c, bold=False)
            cprint(f"{status:<10}{desc_txt}", status_c)

            field_report_lines.append(
                f"  {spec['label']:<16}{(col or '(未匹配)'):<24}{status:<10}"
                f"{'|'.join(spec['candidates'][:5])}"
            )

        mapping[fkey] = matched_map

        unused_cols = [str(c) for c in df.columns if c not in matched_cols_in_df]
        print()
        if unused_cols:
            cprint(f"  [提示] 未识别列（请核查是否为额外字段）: {', '.join(unused_cols)}", YELLOW)
            field_report_lines.append(f"  [提示] 未识别列: {', '.join(unused_cols)}")
        else:
            cprint(f"  V 所有 {len(df.columns)} 列均已识别", GREEN)
            field_report_lines.append("  V 所有列均已识别")

        if missing_required:
            fatal_errors.append(f"{fdesc}缺少必填字段: {', '.join(missing_required)}")
            cprint(f"  X 缺少必填字段: {', '.join(missing_required)}", RED, bold=True)

        print()
        field_report_lines.append("")

    report_path = OUTPUT_DIR / f"field_report_{batch_id}.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(field_report_lines))
    cprint(f"  [报告] 字段识别报告已保存: {report_path}", CYAN)
    print()

    if fatal_errors:
        cprint("=" * 72, RED, bold=True)
        cprint("  X 字段识别存在致命错误，导入中止，请修改表头后重试：", RED, bold=True)
        for e in fatal_errors:
            cprint(f"    · {e}", RED)
        cprint()
        cprint(f"  [报告] 字段识别报告已保存，可查看：", YELLOW)
        cprint(f"  {report_path}", YELLOW)
        cprint("=" * 72, RED)
        return 1

    cprint("=" * 72, BOLD)
    cprint("  [统计] 行级数据质量检查", BOLD)
    cprint("=" * 72, BOLD)
    print()

    veh_map = mapping["vehicles"]
    std_map = mapping["standards"]
    probe_map = mapping["probes"]

    df_v = dfs["vehicles"]
    df_s = dfs["standards"]
    df_p = dfs["probes"]

    cprint("--- 车辆清单 ---", BOLD)
    vehicles = []
    veh_missing = []
    for idx, row in df_v.iterrows():
        plate = str(row[veh_map["plate"]]).strip() if not pd.isna(row[veh_map["plate"]]) else ""
        pos = str(row[veh_map["position"]]).strip() if not pd.isna(row[veh_map["position"]]) else ""
        last = parse_date(row[veh_map["last_calib"]]) if veh_map["last_calib"] else None
        cycle_days = parse_float(row[veh_map["cycle_days"]]) if veh_map["cycle_days"] else None
        car = str(row[veh_map["carriage"]]).strip() if (veh_map["carriage"] and not pd.isna(row[veh_map["carriage"]])) else "主厢"

        missing = []
        if not plate:
            missing.append("车牌")
        if not pos:
            missing.append("探头位置")
        if veh_map["last_calib"] and not last:
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
        cprint(f"  [!] 发现 {len(veh_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<14}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in veh_missing[:25]:
            cprint(f"  {lineno:<8}{plate:<14}{miss}", YELLOW)
        if len(veh_missing) > 25:
            cprint(f"  ... 其余 {len(veh_missing) - 25} 行省略", YELLOW)
        print()
    cprint(f"  V 有效记录: {len(vehicles)} / {len(df_v)} 行", GREEN)
    print()

    cprint("--- 标准温度计读数 ---", BOLD)
    standards = []
    std_missing = []
    for idx, row in df_s.iterrows():
        plate = str(row[std_map["plate"]]).strip() if not pd.isna(row[std_map["plate"]]) else ""
        pos = str(row[std_map["position"]]).strip() if not pd.isna(row[std_map["position"]]) else ""
        ts = parse_date(row[std_map["calib_time"]])
        temp = parse_float(row[std_map["std_temp"]])
        car = str(row[std_map["carriage"]]).strip() if (std_map["carriage"] and not pd.isna(row[std_map["carriage"]])) else "主厢"

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
        cprint(f"  [!] 发现 {len(std_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<14}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in std_missing[:25]:
            cprint(f"  {lineno:<8}{plate:<14}{miss}", YELLOW)
        if len(std_missing) > 25:
            cprint(f"  ... 其余 {len(std_missing) - 25} 行省略", YELLOW)
        print()
    cprint(f"  V 有效记录: {len(standards)} / {len(df_s)} 行", GREEN)
    print()

    cprint("--- 车载探头读数 ---", BOLD)
    probes = []
    probe_missing = []
    for idx, row in df_p.iterrows():
        plate = str(row[probe_map["plate"]]).strip() if not pd.isna(row[probe_map["plate"]]) else ""
        pos = str(row[probe_map["position"]]).strip() if not pd.isna(row[probe_map["position"]]) else ""
        ts = parse_date(row[probe_map["calib_time"]])
        temp = parse_float(row[probe_map["probe_temp"]])
        car = str(row[probe_map["carriage"]]).strip() if (probe_map["carriage"] and not pd.isna(row[probe_map["carriage"]])) else "主厢"

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
        cprint(f"  [!] 发现 {len(probe_missing)} 行缺少关键字段:", YELLOW, bold=True)
        cprint(f"  {'行号':<8}{'车牌':<14}{'缺少字段'}", YELLOW)
        for lineno, plate, miss in probe_missing[:25]:
            cprint(f"  {lineno:<8}{plate:<14}{miss}", YELLOW)
        if len(probe_missing) > 25:
            cprint(f"  ... 其余 {len(probe_missing) - 25} 行省略", YELLOW)
        print()
    cprint(f"  V 有效记录: {len(probes)} / {len(df_p)} 行", GREEN)
    print()

    df_vehicles = pd.DataFrame(vehicles)
    df_standards = pd.DataFrame(standards)
    df_probes = pd.DataFrame(probes)

    df_vehicles.to_pickle(bp / "vehicles.pkl")
    df_standards.to_pickle(bp / "standards.pkl")
    df_probes.to_pickle(bp / "probes.pkl")

    csv_dir = OUTPUT_DIR / "imported" / batch_id
    csv_dir.mkdir(parents=True, exist_ok=True)
    df_vehicles.to_csv(csv_dir / "vehicles.csv", index=False, encoding="utf-8-sig")
    df_standards.to_csv(csv_dir / "standards.csv", index=False, encoding="utf-8-sig")
    df_probes.to_csv(csv_dir / "probes.csv", index=False, encoding="utf-8-sig")

    df_vehicles.to_pickle(DB_DIR / "vehicles.pkl")
    df_standards.to_pickle(DB_DIR / "standards.pkl")
    df_probes.to_pickle(DB_DIR / "probes.pkl")

    meta = {
        "source_files": {k: v.name for k, v in files.items()},
        "counts": {"vehicles": len(vehicles), "standards": len(standards), "probes": len(probes)},
        "row_issues": {"vehicles": len(veh_missing), "standards": len(std_missing), "probes": len(probe_missing)},
    }
    register_batch(batch_id, meta)

    total_issues = len(veh_missing) + len(std_missing) + len(probe_missing)

    cprint("=" * 72, GREEN, bold=True)
    cprint(f"  V 导入成功！批次: {batch_id}", GREEN, bold=True)
    cprint(f"  数据保存: {bp}", GREEN)
    cprint(f"  {len(vehicles)} 辆车 · {len(standards)} 条标准 · {len(probes)} 条探头", GREEN)
    if total_issues:
        cprint(f"  [!] {total_issues} 行存在问题，已跳过，详见上方报告", YELLOW, bold=True)
    cprint("=" * 72, GREEN)
    print()

    cmd_list_batches(None, show=3)
    return 0


def cmd_list_batches(args, show=None):
    idx = load_batch_index()
    cprint("=" * 72, BOLD)
    cprint("  [列表] 历史批次列表", BOLD)
    cprint("=" * 72, BOLD)
    if not idx:
        cprint("  (暂无批次记录，先执行 import 命令)", DIM)
        return 0
    print_limit = show if show else (getattr(args, "show", None) or 20)
    cprint(f"  {'批次号':<24}{'创建时间':<22}{'车辆/标准/探头':<22}{'问题行':<10}", BOLD)
    for e in idx[:print_limit]:
        c = e["counts"]
        i = e["row_issues"]
        total_issue = sum(i.values())
        issue_color = RED if total_issue > 0 else GREEN
        line = (
            f"  {e['batch_id']:<24}"
            f"{datetime.fromisoformat(e['created_at']).strftime('%Y-%m-%d %H:%M:%S'):<22}"
            f"{c['vehicles']} / {c['standards']} / {c['probes']:<14}"
            f"{'共' + str(total_issue) if total_issue else '0':<10}"
        )
        cprint(line, issue_color)
    if len(idx) > print_limit:
        cprint(f"  ... 共 {len(idx)} 个批次，其余省略", DIM)
    cprint("=" * 72, BOLD)
    return 0


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

    used_probe = set()

    for key, std_list in std_grouped.items():
        probe_list = probe_grouped.get(key, [])
        for s in std_list:
            best_p = None
            best_delta = None
            best_id = None
            for pi, p in enumerate(probe_list):
                pid = (key, id(p), pi)
                if pid in used_probe:
                    continue
                delta = abs((s["calibration_time"] - p["calibration_time"]).total_seconds())
                if best_delta is None or delta < best_delta:
                    if delta <= 3600 * 4:
                        best_delta = delta
                        best_p = p
                        best_id = pid
            if best_p is not None:
                used_probe.add(best_id)
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

    batch_id, batch_meta = resolve_batch(getattr(args, "batch", None))
    if batch_id is None:
        cprint("[错误] 未找到可用批次，请先执行 import 命令", RED, bold=True)
        return 1
    df_v, df_s, df_p = load_batch(batch_id)
    if df_v is None:
        cprint(f"[错误] 批次 {batch_id} 数据损坏，请重新导入", RED, bold=True)
        return 1

    tolerance = args.tolerance
    critical = args.critical
    category = args.category
    cycle_days_arg = args.cycle_days

    if category and category in CATEGORY_DEFAULTS:
        if tolerance is None:
            tolerance = CATEGORY_DEFAULTS[category]["tolerance"]
        if critical is None:
            critical = CATEGORY_DEFAULTS[category]["critical"]
    if tolerance is None:
        tolerance = 1.0
    if critical is None:
        critical = round(tolerance * 1.5, 2)

    today = datetime.now()

    cprint("=" * 88, CYAN, bold=True)
    cprint(f"  冷链校准核对工具 - 校准核对    [批次 {batch_id}]", CYAN, bold=True)
    cprint("=" * 88, CYAN)
    info = f"  允许偏差: ±{tolerance}°C   严重超差阈值: ±{critical}°C"
    if category:
        info += f"   品类: {category}"
    cyc_txt = f"全局{cycle_days_arg}天" if cycle_days_arg else "按车辆清单设定"
    info += f"   校准周期: {cyc_txt}"
    cprint(info, CYAN)
    print()

    df_pairs = _build_pairs(df_s, df_p)
    if df_pairs.empty:
        cprint("[错误] 未找到可匹配的标准-探头记录对（需同车牌+车厢+位置，时间差≤4小时）", RED, bold=True)
        return 1

    def classify(row):
        d = row["abs_deviation"]
        if d > critical:
            return STATUS_FAIL
        elif d > tolerance:
            return STATUS_WARN
        return STATUS_PASS

    df_pairs["status"] = df_pairs.apply(classify, axis=1)
    df_pairs = df_pairs.sort_values(["plate", "carriage", "position", "calibration_time"]).reset_index(drop=True)

    df_pairs["consecutive_high"] = False
    grouped = df_pairs.groupby(["plate", "carriage", "position"])
    for key, grp in grouped:
        if len(grp) >= 2:
            grp_sorted = grp.sort_values("calibration_time")
            for i in range(1, len(grp_sorted)):
                cur = grp_sorted.iloc[i]
                prev = grp_sorted.iloc[i - 1]
                if (cur["deviation"] > 0 and prev["deviation"] > 0
                        and cur["abs_deviation"] > tolerance * 0.5):
                    df_pairs.loc[grp_sorted.index[i], "consecutive_high"] = True

    df_merged = df_pairs.merge(
        df_v[["plate", "carriage", "position", "cycle_days", "last_calibration"]],
        on=["plate", "carriage", "position"],
        how="left",
    )
    if cycle_days_arg:
        df_merged["cycle_days"] = cycle_days_arg

    cal_tags = []
    cal_labels = []
    cal_colors = []
    for _, r in df_merged.iterrows():
        cyc = r["cycle_days"] if (r["cycle_days"] and not pd.isna(r["cycle_days"])) else 30
        t, l, c = cal_status(r["last_calibration"], int(cyc), today)
        cal_tags.append(t)
        cal_labels.append(l)
        cal_colors.append(c)
    df_merged["cal_tag"] = cal_tags
    df_merged["cal_label"] = cal_labels
    df_merged["_cal_color"] = cal_colors

    def sort_priority(row):
        p_cal = 0
        if row["cal_tag"] == "[逾期]":
            p_cal = 0
        elif row["cal_tag"] == "[临期]":
            p_cal = 1
        else:
            p_cal = 2
        p_status = 0
        if row["status"] == STATUS_FAIL:
            p_status = 0
        elif row["status"] == STATUS_WARN:
            p_status = 1
        else:
            p_status = 2
        p_consec = 0 if row["consecutive_high"] else 1
        p_dev = -row["abs_deviation"]
        return (p_cal, p_status, p_consec, p_dev)

    df_merged["_sort_key"] = df_merged.apply(sort_priority, axis=1)
    df_merged = df_merged.sort_values("_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)

    check_path = batch_path(batch_id) / "check_results.pkl"
    df_merged.to_pickle(check_path)
    df_merged.to_pickle(DB_DIR / "check_results.pkl")
    export_csv = OUTPUT_DIR / f"check_results_{batch_id}.csv"
    df_out = df_merged.drop(columns=["_cal_color"], errors="ignore")
    df_out.to_csv(export_csv, index=False, encoding="utf-8-sig")

    status_groups = {
        STATUS_PASS: df_merged[df_merged["status"] == STATUS_PASS].copy(),
        STATUS_WARN: df_merged[df_merged["status"] == STATUS_WARN].copy(),
        STATUS_FAIL: df_merged[df_merged["status"] == STATUS_FAIL].copy(),
    }
    consec_high = df_merged[df_merged["consecutive_high"] == True].copy()

    cprint("=" * 110, BOLD)
    cprint(f"  核对结果表格（按 逾期优先 → 严重超差优先 → 临近超差 → 合格）", BOLD)
    cprint("=" * 110, BOLD)
    print()

    hdr = (f"  {'#':>3} {'车牌':<10}{'车厢':<8}{'位置':<10}{'校准时间':<18}"
           f"{'标准':>7}{'探头':>7}{'偏差(°C)':>10}{'等级':<8}{'校准':<10}{'标记':<6}")
    cprint(hdr, BOLD)
    cprint("  " + "-" * 104, DIM)

    print_limit = len(df_merged) if args.verbose else min(len(df_merged), 50)
    shown = 0
    for idx, r in df_merged.head(print_limit).iterrows():
        shown += 1
        dev = f"{r['deviation']:+.3f}"
        status = r["status"]
        status_c = GREEN if status == STATUS_PASS else (YELLOW if status == STATUS_WARN else RED)
        cal_c = r["_cal_color"] or ""
        flag = "↗↗" if r.get("consecutive_high") else ""
        flag_c = MAGENTA if flag else ""

        std_t = f"{r['standard_temp']:.1f}"
        probe_t = f"{r['probe_temp']:.1f}"

        line = (
            f"  {shown:>3} "
            f"{r['plate']:<10}"
            f"{r['carriage']:<8}"
            f"{r['position']:<10}"
            f"{r['calibration_time'].strftime('%Y-%m-%d %H:%M'):<18}"
            f"{std_t:>7}"
            f"{probe_t:>7}"
            f"{dev:>10}  "
            f"{status:<8}"
        )
        line_with_cal = line + f"{r['cal_label']:<10}"
        line_full = line_with_cal + f" {flag:<6}"
        
        status_part = line[:len(line)]
        cal_part = f"{r['cal_label']:<10}"
        flag_part = f" {flag:<6}"
        full_line = status_part + cal_part + flag_part
        print(full_line)

    if not args.verbose and len(df_merged) > print_limit:
        cprint(f"  ... 其余 {len(df_merged) - print_limit} 条略，加 -v 显示全部或查看 {export_csv.name}", DIM)
    print()

    cprint("=" * 110, BOLD)
    cprint(f"  分类汇总  |  共 {len(df_merged)} 条  |  {GREEN}合格 {len(status_groups[STATUS_PASS])}  |  {YELLOW}临近超差 {len(status_groups[STATUS_WARN])}  |  {RED}严重超差 {len(status_groups[STATUS_FAIL])}  |", BOLD)
    rate = len(status_groups[STATUS_PASS]) / len(df_merged) * 100 if len(df_merged) else 0
    cprint(f"             |  合格率 {rate:.1f}%", BOLD)

    overdue_cnt = (df_merged["cal_tag"] == "[逾期]").sum()
    soon_cnt = (df_merged["cal_tag"] == "[临期]").sum()
    consec_cnt = len(consec_high)
    if overdue_cnt or soon_cnt or consec_cnt:
        parts = []
        if overdue_cnt:
            parts.append(f"{RED}逾期 {overdue_cnt} 个{RESET}")
        if soon_cnt:
            parts.append(f"{YELLOW}临期 {soon_cnt} 个{RESET}")
        if consec_cnt:
            parts.append(f"{MAGENTA}连续偏高 {consec_cnt} 个{RESET}")
        cprint(f"             |  {'  |  '.join(parts)}", BOLD)

    cprint(f"  CSV导出: {export_csv}", CYAN)
    cprint("=" * 110, BOLD)
    return 0


def build_comparison(cur_batch, prev_batch, tolerance=1.0):
    _, _ = resolve_batch(cur_batch)
    b1, _ = resolve_batch(prev_batch)
    b2 = cur_batch
    if not b1 or not b2:
        return None

    _, s1, p1 = load_batch(b1)
    _, s2, p2 = load_batch(b2)
    if s1 is None or s2 is None:
        return None

    pairs1 = _build_pairs(s1, p1)
    pairs2 = _build_pairs(s2, p2)

    def latest(df):
        if df.empty:
            return df
        return df.sort_values("calibration_time").groupby(["plate", "carriage", "position"]).tail(1).reset_index(drop=True)

    l1 = latest(pairs1)[["plate", "carriage", "position", "deviation", "abs_deviation", "calibration_time"]]
    l1.columns = ["plate", "carriage", "position", "dev_prev", "abs_prev", "time_prev"]
    l2 = latest(pairs2)[["plate", "carriage", "position", "deviation", "abs_deviation", "calibration_time"]]
    l2.columns = ["plate", "carriage", "position", "dev_cur", "abs_cur", "time_cur"]

    merged = l1.merge(l2, on=["plate", "carriage", "position"], how="outer")
    merged["delta_dev"] = (merged["dev_cur"].fillna(0) - merged["dev_prev"].fillna(0)).round(3)
    merged["abs_change"] = merged["delta_dev"].abs()

    def judge(r):
        prev_bad = pd.notna(r.get("abs_prev")) and r["abs_prev"] is not None and r["abs_prev"] > tolerance
        cur_bad = pd.notna(r.get("abs_cur")) and r["abs_cur"] is not None and r["abs_cur"] > tolerance
        worsen = pd.notna(r["delta_dev"]) and r["abs_change"] > 0.3 and r["delta_dev"] > 0
        if cur_bad and worsen:
            return "明显恶化"
        if cur_bad and not prev_bad:
            return "新出现超差"
        if worsen and r["abs_change"] >= 0.5:
            return "偏差扩大"
        if prev_bad and not cur_bad:
            return "已恢复"
        if pd.notna(r["abs_change"]) and r["abs_change"] <= 0.2:
            return "基本稳定"
        return "略有波动"

    merged["trend"] = merged.apply(judge, axis=1)
    return merged.sort_values("abs_change", ascending=False).reset_index(drop=True)


def _strip_ansi(s):
    return (s.replace(RED, "").replace(YELLOW, "").replace(GREEN, "")
             .replace(CYAN, "").replace(MAGENTA, "").replace(BOLD, "")
             .replace(DIM, "").replace(RESET, ""))


def cmd_summary(args):
    ensure_dirs()

    batch_id, batch_meta = resolve_batch(getattr(args, "batch", None))
    if batch_id is None:
        cprint("[错误] 未找到可用批次，请先执行 import + check", RED, bold=True)
        return 1
    df_v, df_s, df_p = load_batch(batch_id)
    if df_v is None:
        cprint(f"[错误] 批次 {batch_id} 数据损坏", RED, bold=True)
        return 1

    if not has_checked(batch_id):
        cprint(f"[错误] 批次 {batch_id} 尚未执行核对，请先运行：", RED, bold=True)
        cprint(f"  python coldchain_calib.py check -c 冷冻 -d 30 -b {batch_id}", RED)
        return 1
    check_path = batch_path(batch_id) / "check_results.pkl"
    df_check = pd.read_pickle(check_path)

    tolerance = args.tolerance if args.tolerance else 1.0
    critical = args.critical if args.critical else round(tolerance * 1.5, 2)
    cycle_days_global = args.cycle_days
    today = datetime.now()

    idx = load_batch_index()
    prev_batch_id = getattr(args, "compare_with", None)
    prev_meta = None

    if prev_batch_id:
        prev_batch_id, prev_meta = resolve_batch(prev_batch_id)
        if prev_batch_id is None:
            cprint(f"[错误] 指定的对比批次不存在", RED, bold=True)
            return 1
        if not has_checked(prev_batch_id):
            cprint(f"[错误] 对比批次 {prev_batch_id} 尚未执行核对，请先运行：", RED, bold=True)
            cprint(f"  python coldchain_calib.py check -c 冷冻 -d 30 -b {prev_batch_id}", RED)
            return 1
    else:
        for i, e in enumerate(idx):
            if e["batch_id"] == batch_id:
                for j in range(i + 1, len(idx)):
                    candidate = idx[j]["batch_id"]
                    if has_checked(candidate):
                        prev_batch_id = candidate
                        prev_meta = idx[j]
                        break
                break

    unchecked_batches = []
    for e in idx:
        if not has_checked(e["batch_id"]):
            unchecked_batches.append(e["batch_id"])
    if unchecked_batches:
        cprint(f"  [提示] 以下批次已导入但未核对：{', '.join(unchecked_batches)}", YELLOW)
        cprint(f"  请先执行核对命令，避免混用其他批次结果", YELLOW)
        print()

    cprint("=" * 80, BOLD)
    cprint("          冷链探头校准月度摘要报告（车队经理版）", BOLD)
    cprint("=" * 80, BOLD)
    cprint(f"  生成时间: {today.strftime('%Y-%m-%d %H:%M')}    批次: {batch_id}"
           f"{'    对比: ' + prev_batch_id if prev_batch_id else ''}", CYAN)
    print()

    total_vehicles = df_v["plate"].nunique()
    total_probes = len(df_v)
    total_checked = len(df_check)
    p_pass = (df_check["status"] == STATUS_PASS).sum()
    p_warn = (df_check["status"] == STATUS_WARN).sum()
    p_fail = (df_check["status"] == STATUS_FAIL).sum()
    pass_rate = (p_pass / total_checked * 100) if total_checked else 0

    cprint("【一、总体情况】", BOLD)
    cprint(f"  • 登记车辆总数: {total_vehicles} 台    登记探头总数: {total_probes} 个", "")
    cprint(f"  • 本月校准核对: {total_checked} 条记录", "")
    pr_color = GREEN if pass_rate >= 85 else (YELLOW if pass_rate >= 70 else RED)
    cprint(f"  • 合格率: {pr_color}{pass_rate:.1f}%{RESET}"
           f"   ({GREEN}{p_pass}合格{RESET} / {YELLOW}{p_warn}临近{RESET} / {RED}{p_fail}超差{RESET})", "")
    print()

    cprint("【二、需立即复检车辆】（存在临近超差或严重超差）", BOLD)
    problem_probes = df_check[df_check["status"] != STATUS_PASS].copy()
    problem_vehicles = problem_probes.groupby("plate").agg(
        问题数=("status", "count"),
        严重超差=("status", lambda s: (s == STATUS_FAIL).sum()),
        临近超差=("status", lambda s: (s == STATUS_WARN).sum()),
        最严重偏差=("abs_deviation", "max"),
        逾期探头=("cal_tag", lambda s: sum(1 for x in s if "逾期" in str(x))),
    ).sort_values(["严重超差", "问题数"], ascending=False)

    df_problem_veh_export = problem_vehicles.reset_index().copy()
    df_problem_veh_export.columns = [
        "车牌", "问题探头数", "严重超差数", "临近超差数", "逾期探头数", "最大偏差(°C)"
    ]
    if problem_vehicles.empty:
        cprint("  V 本月所有校准记录均合格，无需复检。", GREEN)
    else:
        hdr = f"  {'车牌':<12}{'问题探头':>8}{'严重超差':>10}{'临近超差':>10}{'逾期':>6}  {'最大偏差(°C)':>12}"
        cprint(hdr, BOLD)
        for plate, row in problem_vehicles.iterrows():
            color = RED if row["严重超差"] > 0 else YELLOW
            line = (f"  {plate:<12}{int(row['问题数']):>8}{int(row['严重超差']):>10}"
                    f"{int(row['临近超差']):>10}{int(row['逾期探头']):>6}  {row['最严重偏差']:>12.3f}")
            cprint(line, color)
    print()

    cprint("【三、建议停用探头】（严重超差 + 连续两次偏高）", BOLD)
    stop_candidates = df_check[
        (df_check["status"] == STATUS_FAIL) | (df_check["consecutive_high"] == True)
    ].copy()

    def make_reason(r):
        parts = []
        if r["status"] == STATUS_FAIL:
            parts.append("严重超差")
        if r["consecutive_high"]:
            parts.append("连续偏高")
        if "逾期" in str(r.get("cal_tag", "")):
            parts.append("校准逾期")
        return "+".join(parts) if parts else "-"

    stop_candidates["停用原因"] = stop_candidates.apply(make_reason, axis=1)
    df_stop_export = stop_candidates.copy()
    df_stop_export = df_stop_export.rename(columns={
        "plate": "车牌", "carriage": "车厢号", "position": "探头位置",
        "calibration_time": "校准时间", "standard_temp": "标准温度(°C)",
        "probe_temp": "探头读数(°C)", "deviation": "偏差(°C)",
        "abs_deviation": "绝对偏差(°C)", "status": "偏差等级",
        "cal_label": "校准到期状态", "last_calibration": "上次校准时间",
        "cycle_days": "校准周期(天)",
    })
    export_cols = [
        "车牌", "车厢号", "探头位置", "校准时间", "标准温度(°C)", "探头读数(°C)",
        "偏差(°C)", "偏差等级", "上次校准时间", "校准周期(天)", "校准到期状态", "停用原因",
    ]
    export_cols = [c for c in export_cols if c in df_stop_export.columns]
    df_stop_export = df_stop_export[export_cols]

    if stop_candidates.empty:
        cprint("  V 本月无需要建议停用的探头。", GREEN)
    else:
        hdr = (f"  {'车牌':<12}{'车厢':<8}{'位置':<10}{'偏差(°C)':>10}  "
               f"{'等级':<8}{'校准状态':<12}{'停用原因':<24}")
        cprint(hdr, BOLD)
        for _, r in stop_candidates.iterrows():
            dev = f"{r['deviation']:+.3f}"
            lvl = r["status"]
            lvl_c = RED if lvl == STATUS_FAIL else (YELLOW if lvl == STATUS_WARN else GREEN)
            cal_c = RED if "逾期" in str(r["cal_tag"]) else (YELLOW if "临期" in str(r["cal_tag"]) else "")
            reason = r["停用原因"]
            cprint(
                f"  {r['plate']:<12}{r['carriage']:<8}{r['position']:<10}"
                f"{dev:>10}  ", lvl_c, bold=False)
            cprint(f"{lvl:<8}", lvl_c)
            cprint(f"{r['cal_label']:<12}", cal_c)
            cprint(f" {reason:<24}", RED if "严重" in reason else YELLOW)
            print()
    print()

    cprint(f"【四、即将到期校准任务】（未来15天内，含逾期）", BOLD)
    upcoming = []
    for _, v in df_v.iterrows():
        last = v["last_calibration"]
        cyc = v["cycle_days"] if (v["cycle_days"] and not pd.isna(v["cycle_days"])) else (cycle_days_global or 30)
        if isinstance(last, datetime):
            next_calib = last + timedelta(days=int(cyc))
            days_left = (next_calib - today).days
            if days_left <= 15:
                upcoming.append({
                    "车牌": v["plate"],
                    "车厢号": v["carriage"],
                    "探头位置": v["position"],
                    "上次校准": last,
                    "下次校准": next_calib,
                    "剩余天数": days_left,
                    "状态": ("已逾期" + str(-days_left) + "天") if days_left < 0 else (str(days_left) + "天"),
                    "紧急程度": 0 if days_left < 0 else (1 if days_left <= 7 else 2),
                })

    df_upcoming_export = pd.DataFrame()
    if not upcoming:
        cprint("  V 未来15天内无即将到期的校准任务。", GREEN)
    else:
        upcoming.sort(key=lambda x: (x["紧急程度"], x["剩余天数"]))
        df_upcoming_export = pd.DataFrame(upcoming).drop(columns=["紧急程度"])
        hdr = (f"  {'车牌':<12}{'车厢':<8}{'位置':<10}{'上次校准':<14}"
               f"{'下次校准':<14}{'剩余':>8}")
        cprint(hdr, BOLD)
        for u in upcoming:
            color = RED if u["剩余天数"] < 0 else (YELLOW if u["剩余天数"] <= 7 else CYAN)
            line = (f"  {u['车牌']:<12}{u['车厢号']:<8}{u['探头位置']:<10}"
                    f"{u['上次校准'].strftime('%Y-%m-%d'):<14}"
                    f"{u['下次校准'].strftime('%Y-%m-%d'):<14}{u['状态']:>8}")
            cprint(line, color)
    print()

    df_compare_export = pd.DataFrame()
    df_cmp = None
    if prev_batch_id:
        df_cmp = build_comparison(batch_id, prev_batch_id, tolerance=tolerance)
        if df_cmp is not None and not df_cmp.empty:
            cprint(f"【五、与上一批次（{prev_batch_id}）偏差变化对比】", BOLD)
            df_cmp_export = df_cmp.copy()
            df_cmp_export.columns = [
                "车牌", "车厢号", "探头位置",
                "上次偏差(°C)", "上次|偏差|", "上次校准时间",
                "本次偏差(°C)", "本次|偏差|", "本次校准时间",
                "偏差变化量(°C)", "变化绝对值", "趋势判断",
            ]
            df_compare_export = df_cmp_export
            high_change = df_cmp[df_cmp["abs_change"] >= 0.3]
            if high_change.empty:
                cprint("  V 两次核对偏差基本一致，未发现明显漂移。", GREEN)
            else:
                hdr = (f"  {'车牌':<12}{'车厢':<8}{'位置':<10}"
                       f"{'上次偏差':>10}{'本次偏差':>10}{'变化':>10}  {'趋势':<10}")
                cprint(hdr, BOLD)
                for _, r in high_change.head(15).iterrows():
                    d_prev = f"{r['dev_prev']:+.2f}" if pd.notna(r.get("dev_prev")) else "-"
                    d_cur = f"{r['dev_cur']:+.2f}" if pd.notna(r.get("dev_cur")) else "-"
                    d_delta = f"{r['delta_dev']:+.2f}" if pd.notna(r.get("delta_dev")) else "-"
                    trend = r["trend"]
                    tc = RED if trend in ("明显恶化", "新出现超差") else (YELLOW if trend == "偏差扩大" else GREEN)
                    line = (f"  {r['plate']:<12}{r['carriage']:<8}{r['position']:<10}"
                            f"{d_prev:>10}{d_cur:>10}{d_delta:>10}  {trend:<10}")
                    cprint(line, tc)
                if len(high_change) > 15:
                    cprint(f"  ... 其余 {len(high_change) - 15} 条变化记录详见 Excel 报告", DIM)
            print()

    cprint("【运营建议】", BOLD)
    suggestions = []
    if p_fail > 0:
        suggestions.append(f"• {RED}{p_fail}个探头严重超差{RESET}，建议立即停用并安排专业维修或更换，避免货损风险。")
    if p_warn > 0:
        suggestions.append(f"• {YELLOW}{p_warn}个探头临近超差{RESET}，建议增加抽查频率并提前安排校准。")
    if not stop_candidates.empty:
        suggestions.append(f"• {RED}{len(stop_candidates)}个探头建议停用{RESET}，调度时避免分配给高价值冷链货物。")
    overdue = [u for u in upcoming if u["剩余天数"] < 0]
    soon = [u for u in upcoming if 0 <= u["剩余天数"] <= 7]
    if overdue:
        suggestions.append(f"• {RED}{len(overdue)}个探头已逾期校准{RESET}，请立即安排校准并评估期间数据可靠性。")
    elif soon:
        suggestions.append(f"• {YELLOW}{len(soon)}个探头一周内需校准{RESET}，请提前预约校准资源。")
    if pass_rate < 80:
        suggestions.append(f"• 整体合格率仅 {pass_rate:.1f}%，建议排查探头老化或安装位置问题，组织统一专项校准。")
    if prev_batch_id and df_cmp is not None and not df_cmp.empty:
        worsen = len(df_cmp[df_cmp["trend"].isin(["明显恶化", "新出现超差", "偏差扩大"])])
        if worsen:
            suggestions.append(f"• 较上次核对有{YELLOW}{worsen}个探头偏差恶化{RESET}，建议关注探头老化趋势并缩短校准周期。")
    if not suggestions:
        suggestions.append("• 本月整体运营状况良好，请继续保持现有校准管理节奏。")

    for s in suggestions:
        cprint(f"  {s}", "")
    print()

    report_base = OUTPUT_DIR / f"summary_{batch_id}_{today.strftime('%Y%m%d')}"
    txt_path = report_base.with_suffix(".txt")
    xlsx_path = report_base.with_suffix(".xlsx")

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("冷链探头校准月度摘要报告（车队经理版）\n")
        f.write(f"生成时间: {today.strftime('%Y-%m-%d %H:%M')}    批次: {batch_id}\n")
        if prev_batch_id:
            f.write(f"对比批次: {prev_batch_id}\n")
        f.write("=" * 70 + "\n\n")
        f.write("一、总体情况\n")
        f.write(f"  登记车辆: {total_vehicles} 台, 探头: {total_probes} 个\n")
        f.write(f"  核对记录: {total_checked} 条, 合格率: {pass_rate:.1f}%\n")
        f.write(f"  合格: {p_pass}, 临近超差: {p_warn}, 严重超差: {p_fail}\n\n")

        f.write("二、需立即复检车辆\n")
        if problem_vehicles.empty:
            f.write("  无\n")
        else:
            for plate, row in problem_vehicles.iterrows():
                f.write(f"  {plate}: 问题{int(row['问题数'])}个, 严重{int(row['严重超差'])}个, "
                        f"临近{int(row['临近超差'])}个, 最大偏差{row['最严重偏差']:.3f}°C\n")
        f.write("\n")

        f.write("三、建议停用探头\n")
        if stop_candidates.empty:
            f.write("  无\n")
        else:
            for _, r in stop_candidates.iterrows():
                f.write(f"  {r['plate']} {r['carriage']} {r['position']}: "
                        f"偏差{r['deviation']:+.3f}°C ({r['status']}), "
                        f"校准状态: {r['cal_label']}, 原因: {r['停用原因']}\n")
        f.write("\n")

        f.write("四、即将到期校准任务\n")
        if not upcoming:
            f.write("  无\n")
        else:
            for u in upcoming:
                f.write(f"  {u['车牌']} {u['车厢号']} {u['探头位置']}: "
                        f"上次{u['上次校准'].strftime('%Y-%m-%d')} → "
                        f"下次{u['下次校准'].strftime('%Y-%m-%d')} ({u['状态']})\n")
        f.write("\n")

        if not df_compare_export.empty and prev_batch_id:
            f.write(f"五、批次对比（{prev_batch_id} → {batch_id}）\n")
            for _, r in df_cmp.iterrows():
                dp = f"{r['dev_prev']:+.2f}" if pd.notna(r.get("dev_prev")) else "-"
                dc = f"{r['dev_cur']:+.2f}" if pd.notna(r.get("dev_cur")) else "-"
                dd = f"{r['delta_dev']:+.2f}" if pd.notna(r.get("delta_dev")) else "-"
                f.write(f"  {r['plate']} {r['carriage']} {r['position']}: "
                        f"{dp} → {dc} (Δ{dd}) [{r['trend']}]\n")
            f.write("\n")

        f.write("六、运营建议\n")
        for s in suggestions:
            f.write(f"  {_strip_ansi(s)}\n")

    try:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            overview = pd.DataFrame([
                ["批次号", batch_id],
                ["生成时间", today.strftime("%Y-%m-%d %H:%M:%S")],
                ["登记车辆总数（台）", total_vehicles],
                ["登记探头总数（个）", total_probes],
                ["本月核对记录（条）", total_checked],
                ["合格（条）", int(p_pass)],
                ["临近超差（条）", int(p_warn)],
                ["严重超差（条）", int(p_fail)],
                ["合格率（%）", round(pass_rate, 2)],
                ["允许偏差阈值（°C）", tolerance],
                ["严重超差阈值（°C）", critical],
                ["建议停用探头数", len(stop_candidates)],
                ["逾期校准探头数", len(overdue)],
                ["7天内到期探头数", len(soon)],
                ["对比批次", prev_batch_id if prev_batch_id else "（无）"],
            ], columns=["项目", "数值"])
            overview.to_excel(writer, sheet_name="总体概览", index=False)

            if not df_problem_veh_export.empty:
                df_problem_veh_export.to_excel(writer, sheet_name="需复检车辆", index=False)
            else:
                pd.DataFrame([{"说明": "本月无需复检车辆，全部合格 V"}]).to_excel(
                    writer, sheet_name="需复检车辆", index=False)

            if not df_stop_export.empty:
                for c in df_stop_export.columns:
                    if pd.api.types.is_datetime64_any_dtype(df_stop_export[c]):
                        df_stop_export[c] = df_stop_export[c].dt.strftime("%Y-%m-%d %H:%M")
                df_stop_export.to_excel(writer, sheet_name="建议停用探头", index=False)
            else:
                pd.DataFrame([{"说明": "本月无建议停用探头 V"}]).to_excel(
                    writer, sheet_name="建议停用探头", index=False)

            if not df_upcoming_export.empty:
                for c in df_upcoming_export.columns:
                    if pd.api.types.is_datetime64_any_dtype(df_upcoming_export[c]):
                        df_upcoming_export[c] = df_upcoming_export[c].dt.strftime("%Y-%m-%d")
                df_upcoming_export.to_excel(writer, sheet_name="即将到期任务", index=False)
            else:
                pd.DataFrame([{"说明": "未来15天内无到期校准任务 V"}]).to_excel(
                    writer, sheet_name="即将到期任务", index=False)

            if not df_compare_export.empty:
                for c in df_compare_export.columns:
                    if pd.api.types.is_datetime64_any_dtype(df_compare_export[c]):
                        df_compare_export[c] = df_compare_export[c].apply(
                            lambda x: x.strftime("%Y-%m-%d %H:%M") if pd.notna(x) and isinstance(x, datetime) else x
                        )
                df_compare_export.to_excel(writer, sheet_name="偏差变化对比", index=False)
            else:
                pd.DataFrame([{"说明": "未找到上一批次对比数据或无明显变化"}]).to_excel(
                    writer, sheet_name="偏差变化对比", index=False)

            sug_df = pd.DataFrame({"运营建议": [_strip_ansi(s)[2:] for s in suggestions]})
            sug_df.to_excel(writer, sheet_name="运营建议", index=False)

            wb = writer.book
            for sn in wb.sheetnames:
                ws = wb[sn]
                for col in ws.columns:
                    max_len = 10
                    letter = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value is not None:
                                ln = len(str(cell.value))
                                if ln > max_len:
                                    max_len = ln
                        except Exception:
                            pass
                    ws.column_dimensions[letter].width = min(max_len + 4, 50)

        cprint(f"  [文本] 文本报告: {txt_path}", CYAN, bold=True)
        cprint(f"  [Excel] Excel报告: {xlsx_path}", CYAN, bold=True)
    except Exception as ex:
        cprint(f"  [!] Excel导出失败: {ex}（文本报告仍已生成）", YELLOW)
        cprint(f"  [文本] 文本报告: {txt_path}", CYAN, bold=True)

    cprint("=" * 80, BOLD)
    return 0


def cmd_clean(args):
    ensure_dirs()
    keep = args.keep
    dry_run = args.dry_run
    archive = args.archive
    delete = args.delete
    skip_confirm = args.yes

    if keep < 1:
        cprint("[错误] 保留批次数量至少为1", RED, bold=True)
        return 1

    if not archive and not delete:
        archive = True
        cprint("[提示] 未指定操作模式，默认归档（--archive）。加 --delete 可直接删除。", YELLOW)
        print()

    idx = load_batch_index()
    if len(idx) <= keep:
        cprint(f"[信息] 当前共 {len(idx)} 个批次，保留 {keep} 个，无需清理。", GREEN)
        return 0

    keep_entries = idx[:keep]
    clean_entries = idx[keep:]

    cprint("=" * 72, BOLD)
    cprint("  批次清理工具", BOLD)
    cprint("=" * 72, BOLD)
    cprint(f"  当前批次总数: {len(idx)}    保留最近: {keep}个    需清理: {len(clean_entries)}个", "")
    mode = "归档为zip" if archive else "直接删除"
    cprint(f"  操作模式: {mode}", "")
    if dry_run:
        cprint(f"  运行模式: 仅预览（--dry-run）", YELLOW)
    print()

    cprint(f"  【保留的批次（最近{keep}个）】", GREEN, bold=True)
    for i, e in enumerate(keep_entries):
        c = e["counts"]
        cprint(f"  {i+1:>2}. {e['batch_id']:<24} {e['created_at'][:19]}  {c['vehicles']}/{c['standards']}/{c['probes']} 条", GREEN)
    print()

    cprint(f"  【将清理的批次（{len(clean_entries)}个）】", RED, bold=True)
    for i, e in enumerate(clean_entries):
        c = e["counts"]
        cprint(f"  {i+1:>2}. {e['batch_id']:<24} {e['created_at'][:19]}  {c['vehicles']}/{c['standards']}/{c['probes']} 条", RED)
    print()

    if not skip_confirm and not dry_run:
        confirm = input(f"  确认{mode}以上 {len(clean_entries)} 个批次？(yes/NO): ").strip().lower()
        if confirm != "yes":
            cprint("  已取消操作。", YELLOW)
            return 0
        print()

    archive_path = None
    if archive:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        archive_name = f"archive_{ts}_batch_{len(clean_entries)}.zip"
        archive_path = ARCHIVE_DIR / archive_name

    errors = []
    archived_count = 0
    deleted_count = 0

    for e in clean_entries:
        bid = e["batch_id"]
        bp = batch_path(bid)

        try:
            if not bp.exists():
                cprint(f"  [!] 批次 {bid} 目录不存在，跳过", YELLOW)
                continue

            if archive and not dry_run:
                if not zipfile.is_zipfile(archive_path):
                    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        for root, _, files in os.walk(bp):
                            for f in files:
                                fpath = Path(root) / f
                                arcname = f"{bid}/{fpath.relative_to(bp)}"
                                zf.write(fpath, arcname)
                else:
                    with zipfile.ZipFile(archive_path, "a", zipfile.ZIP_DEFLATED) as zf:
                        for root, _, files in os.walk(bp):
                            for f in files:
                                fpath = Path(root) / f
                                arcname = f"{bid}/{fpath.relative_to(bp)}"
                                zf.write(fpath, arcname)

                archived_count += 1
                cprint(f"  V 已归档: {bid}", GREEN)
                shutil.rmtree(bp)
                cprint(f"  V 已删除源目录: {bid}", DIM)
            elif delete and not dry_run:
                shutil.rmtree(bp)
                deleted_count += 1
                cprint(f"  V 已删除: {bid}", GREEN)
            else:
                cprint(f"  [预览] 将处理: {bid}", DIM)

        except Exception as ex:
            errors.append(f"{bid}: {str(ex)}")
            cprint(f"  X 处理失败: {bid} - {ex}", RED)

    if not dry_run:
        new_idx = keep_entries
        save_batch_index(new_idx)

        cprint()
        cprint("=" * 72, BOLD)
        if archive:
            cprint(f"  V 清理完成！已归档 {archived_count} 个批次", GREEN, bold=True)
            cprint(f"  压缩包: {archive_path}", GREEN)
        else:
            cprint(f"  V 清理完成！已删除 {deleted_count} 个批次", GREEN, bold=True)

        if errors:
            cprint(f"  X 失败 {len(errors)} 个:", RED)
            for e in errors:
                cprint(f"    · {e}", RED)
        cprint(f"  剩余批次: {len(new_idx)} 个", "")
        cprint("=" * 72, BOLD)
    else:
        cprint("=" * 72, YELLOW)
        cprint(f"  [预览] 以上是将要执行的操作，去掉 --dry-run 即可实际执行", YELLOW, bold=True)
        cprint("=" * 72, YELLOW)

    return 0


def main():
    parser = argparse.ArgumentParser(
        prog="coldchain-calib",
        description="冷链探头校准核对工具 - 适用于冷藏车队月底批量校准数据核对",
        epilog="""
示例:
  coldchain-calib import                       # 从 data/input 导入三张表，生成批次
  coldchain-calib batches                      # 查看历史批次列表
  coldchain-calib check -c 冷冻 -d 30          # 按冷冻品类+30天周期核对(最新批次)
  coldchain-calib check -b BAT-202605 -d 45    # 指定5月批次+45天周期核对
  coldchain-calib check -t 1.0 -k 2.0 -d 30    # 自定义偏差±1°C/超差±2°C/周期30天
  coldchain-calib summary -d 30                # 生成30天周期的摘要报告(含Excel)
  coldchain-calib summary -b BAT-202605        # 生成指定批次的摘要
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", help="可用命令")

    p_import = sub.add_parser("import", help="导入车辆清单、标准读数、探头读数数据(生成批次)")

    p_batches = sub.add_parser("batches", help="查看历史批次列表")
    p_batches.add_argument("-n", "--show", type=int, default=20, help="显示最近N个批次，默认20")

    p_check = sub.add_parser("check", help="执行校准偏差核对")
    p_check.add_argument("-t", "--tolerance", type=float, default=None,
                         help="允许偏差阈值(°C)，默认根据品类或1.0°C")
    p_check.add_argument("-k", "--critical", type=float, default=None,
                         help="严重超差阈值(°C)，默认根据品类或偏差的1.5倍")
    p_check.add_argument("-c", "--category", default=None,
                         choices=list(CATEGORY_DEFAULTS.keys()),
                         help="运输品类(自动匹配阈值): 冷冻/冷藏/恒温/果蔬/通用")
    p_check.add_argument("-d", "--cycle-days", type=int, default=None,
                         help="校准周期(天)，如 30 / 45。覆盖车辆清单中的设定")
    p_check.add_argument("-b", "--batch", default=None,
                         help="指定批次号(支持完整或后缀)，默认使用最新批次")
    p_check.add_argument("-v", "--verbose", action="store_true",
                         help="完整显示合格记录(默认仅显示前20条)")

    p_summary = sub.add_parser("summary", help="生成车队经理摘要报告(文本+Excel)")
    p_summary.add_argument("-t", "--tolerance", type=float, default=None,
                           help="允许偏差阈值(°C)")
    p_summary.add_argument("-k", "--critical", type=float, default=None,
                           help="严重超差阈值(°C)")
    p_summary.add_argument("-d", "--cycle-days", type=int, default=None,
                           help="校准周期(天)，默认30天或车辆清单设定")
    p_summary.add_argument("-b", "--batch", default=None,
                           help="指定批次号，默认使用最新批次")
    p_summary.add_argument("--compare-with", default=None,
                           help="指定对比批次号(不指定则自动找上一批已核对的)")

    p_clean = sub.add_parser("clean", help="批次清理：保留最近N个，旧批次归档或删除")
    p_clean.add_argument("--keep", type=int, default=10,
                         help="保留最近N个批次，默认10")
    p_clean.add_argument("--archive", action="store_true",
                         help="将清理的批次打包为zip压缩包，保存到archive目录")
    p_clean.add_argument("--delete", action="store_true",
                         help="不归档，直接删除旧批次（请谨慎使用）")
    p_clean.add_argument("--dry-run", action="store_true",
                         help="仅预览操作，不实际删除或归档")
    p_clean.add_argument("-y", "--yes", action="store_true",
                         help="跳过确认提示，直接执行")

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 0

    ensure_dirs()

    if args.command == "import":
        return cmd_import(args)
    elif args.command == "batches":
        return cmd_list_batches(args)
    elif args.command == "check":
        return cmd_check(args)
    elif args.command == "summary":
        return cmd_summary(args)
    elif args.command == "clean":
        return cmd_clean(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main() or 0)