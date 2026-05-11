import json
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from nurse_scheduler_v2 import RuleConfig, build_employees, date_range, export_to_excel, parse_date, solve_schedule

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.sample.json"
SHIFT_OPTIONS = ["D", "E", "N"]


def _optional_shift_cell(val) -> str:
    """空白欄位在 data_editor 常變成 float NaN；str(NaN) 會變成 'nan' 而誤觸驗證。"""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except TypeError:
        pass
    s = str(val).strip()
    if s.lower() == "nan":
        return ""
    return s


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_config(cfg: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def employees_to_df(employees: list[dict]) -> pd.DataFrame:
    rows = []
    for emp in employees:
        pref = emp.get("preferred_shifts")
        if pref is None and emp.get("preferred_shift"):
            pref = [emp.get("preferred_shift")]
        pref = pref or []
        rows.append(
            {
                "name": emp.get("name", ""),
                "main_shift": pref[0] if len(pref) >= 1 else "D",
                "second_shift": pref[1] if len(pref) >= 2 else "",
                "min_rest_days": int(emp.get("min_rest_days", 8)),
                "previous_last_shift": str(emp.get("previous_last_shift", "1off")),
                "previous_night_count": "" if emp.get("previous_night_count") in (None, "") else str(emp.get("previous_night_count")),
                "requested_r_dates": list(emp.get("requested_r_dates", [])),
            }
        )
    return pd.DataFrame(rows)


def to_label(d: date) -> str:
    weekday = "一二三四五六日"[d.weekday()]
    return f"{d.strftime('%Y-%m-%d')}（{weekday}）"


def row_to_employee(row: pd.Series, r_dates: list[str]) -> dict:
    name = str(row["name"]).strip()
    main_shift = str(row["main_shift"]).strip()
    second_shift = _optional_shift_cell(row.get("second_shift"))
    if main_shift not in SHIFT_OPTIONS:
        raise ValueError(f"{name} 主偏好班別必須是 D/E/N")
    preferred_shifts = [main_shift]
    if second_shift:
        if second_shift not in SHIFT_OPTIONS:
            raise ValueError(f"{name} 次偏好班別必須是 D/E/N")
        if second_shift == main_shift:
            raise ValueError(f"{name} 主偏好與次偏好不可相同")
        preferred_shifts.append(second_shift)
    return {
        "name": name,
        "preferred_shifts": preferred_shifts,
        "min_rest_days": int(row["min_rest_days"]),
        "previous_last_shift": str(row["previous_last_shift"]).strip(),
        "previous_night_count": (
            int(str(row["previous_night_count"]).strip()) if str(row["previous_night_count"]).strip() else None
        ),
        "requested_r_dates": r_dates,
    }


st.set_page_config(page_title="護理排班系統", page_icon="🗓️", layout="wide")
st.title("護理排班系統")
st.caption(
    "範例：同時綁「六人上段夜班＋月底 E/N 剛好 20」常無解，請分批填。"
    "「雅芳2N、佩萱1E、雅萍2E、尹汝2E」與四起 20 併用也易無解；預設檔該四欄用 1off 以保證可排。"
)

try:
    cfg = load_config()
except FileNotFoundError:
    st.error(f"找不到設定檔：{CONFIG_PATH}（請確認已與 streamlit_app.py 一併上傳到 GitHub）")
    st.stop()
if "base_seed" not in st.session_state:
    st.session_state.base_seed = int(cfg.get("random_seed", 42))
if "version_offset" not in st.session_state:
    st.session_state.version_offset = 0

default_start = parse_date(cfg["date_range"]["start"])
default_end = parse_date(cfg["date_range"]["end"])

col1, col2, col3 = st.columns(3)
with col1:
    start_date = st.date_input("排班開始日期", value=default_start)
with col2:
    end_date = st.date_input("排班結束日期", value=default_end)
with col3:
    output_file = st.text_input("輸出檔名", value=cfg.get("output_file", "nurse_schedule.xlsx"))

if start_date > end_date:
    st.error("開始日期不可晚於結束日期")
    st.stop()

all_days = date_range(start_date, end_date)
day_options = [d.strftime("%Y-%m-%d") for d in all_days]
label_map = {d.strftime("%Y-%m-%d"): to_label(d) for d in all_days}

st.subheader("每日需求設定")
d_col, e_col, n_col = st.columns(3)
count_options = list(range(1, 11))
d_cfg = cfg["daily_requirements"]["D"]
e_cfg = cfg["daily_requirements"]["E"]
n_cfg = cfg["daily_requirements"]["N"]
d_default = list(range(int(d_cfg["min"]), int(d_cfg["max"]) + 1))
e_min = int(e_cfg["min"] if "min" in e_cfg else e_cfg["required"])
e_max = int(e_cfg["max"] if "max" in e_cfg else e_cfg["required"])
n_min = int(n_cfg["min"] if "min" in n_cfg else n_cfg["required"])
n_max = int(n_cfg["max"] if "max" in n_cfg else n_cfg["required"])
e_default = list(range(e_min, e_max + 1))
n_default = list(range(n_min, n_max + 1))

with d_col:
    d_counts = st.multiselect("白班 D 人數", options=count_options, default=d_default)
with e_col:
    e_counts = st.multiselect("小夜 E 人數", options=count_options, default=e_default)
with n_col:
    n_counts = st.multiselect("大夜 N 人數", options=count_options, default=n_default)

default_holidays = [d for d in cfg.get("national_holidays", []) if d in day_options]
national_holidays = st.multiselect(
    "國定假日",
    options=day_options,
    default=default_holidays,
    format_func=lambda x: label_map.get(x, x),
)

st.subheader("員工基本設定")

emp_df = employees_to_df(cfg["employees"])
basic_df = emp_df.drop(columns=["requested_r_dates"])
edited_df = st.data_editor(
    basic_df,
    use_container_width=True,
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "main_shift": st.column_config.SelectboxColumn("主偏好班別", options=SHIFT_OPTIONS, required=True),
        "second_shift": st.column_config.SelectboxColumn("次偏好班別(可空白)", options=[""] + SHIFT_OPTIONS),
        "previous_last_shift": st.column_config.TextColumn("上段最後一班(例:3D,2OF)"),
        "previous_night_count": st.column_config.TextColumn(
            "上段夜班數(留空=不限制；表內如12+8=本曆月至月底尚須8班E/N，上段+本月剛好共20班)"
        ),
    },
)

st.subheader("預假")
r_date_map: dict[str, list[str]] = {}
for _, row in edited_df.iterrows():
    name = str(row["name"]).strip()
    if not name:
        continue
    old = next((e.get("requested_r_dates", []) for e in cfg["employees"] if e.get("name") == name), [])
    default_r = [d for d in old if d in day_options]
    selected = st.multiselect(
        f"{name}",
        options=day_options,
        default=default_r,
        format_func=lambda x: label_map.get(x, x),
        key=f"r_{name}",
    )
    r_date_map[name] = selected

st.subheader("產生班表")
st.caption(f"目前版次：第 {st.session_state.version_offset + 1} 版")

btn_col1, btn_col2 = st.columns(2)
generate_clicked = btn_col1.button("產生班表（第1版）", type="primary", use_container_width=True)
next_clicked = btn_col2.button("下一版", use_container_width=True)

if generate_clicked or next_clicked:
    try:
        if generate_clicked:
            st.session_state.version_offset = 0
        elif next_clicked:
            st.session_state.version_offset += 1

        if not d_counts or not e_counts or not n_counts:
            raise ValueError("每日需求的 D/E/N 人數都至少要選一個數字")

        seed_used = st.session_state.base_seed + st.session_state.version_offset
        employees = []
        for _, row in edited_df.iterrows():
            name = str(row["name"]).strip()
            if not name:
                continue
            employees.append(row_to_employee(row, r_date_map.get(name, [])))

        new_cfg = {
            "date_range": {"start": start_date.strftime("%Y-%m-%d"), "end": end_date.strftime("%Y-%m-%d")},
            "daily_requirements": {
                "D": {"min": min(d_counts), "max": max(d_counts)},
                "E": {"min": min(e_counts), "max": max(e_counts)},
                "N": {"min": min(n_counts), "max": max(n_counts)},
            },
            "national_holidays": national_holidays,
            "random_seed": int(seed_used),
            "employees": employees,
            "output_file": output_file.strip() or "nurse_schedule.xlsx",
        }
        save_config(new_cfg)

        employee_objs = build_employees(new_cfg["employees"])
        days = date_range(parse_date(new_cfg["date_range"]["start"]), parse_date(new_cfg["date_range"]["end"]))
        rule = RuleConfig(
            min_d=new_cfg["daily_requirements"]["D"]["min"],
            max_d=new_cfg["daily_requirements"]["D"]["max"],
            min_e=new_cfg["daily_requirements"]["E"]["min"],
            max_e=new_cfg["daily_requirements"]["E"]["max"],
            min_n=new_cfg["daily_requirements"]["N"]["min"],
            max_n=new_cfg["daily_requirements"]["N"]["max"],
            national_holidays={parse_date(d) for d in new_cfg["national_holidays"]},
            random_seed=seed_used,
        )
        schedule = solve_schedule(employee_objs, days, rule)
        out_name = new_cfg["output_file"].strip() or "nurse_schedule.xlsx"
        out_path = BASE_DIR / out_name
        export_to_excel(str(out_path), employee_objs, schedule, days, rule.national_holidays)
        st.success(f"完成（第 {st.session_state.version_offset + 1} 版）")
        if out_path.is_file():
            st.download_button(
                label="下載 Excel",
                data=out_path.read_bytes(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{seed_used}_{st.session_state.version_offset}",
            )
    except Exception as e:
        st.error(f"產生失敗：{e}")


if __name__ == "__main__":
    import subprocess
    import sys

    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
    except Exception:  # noqa: BLE001
        get_script_run_ctx = lambda: None  # type: ignore[misc]

    if get_script_run_ctx() is None:
        subprocess.run(
            [sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()), *sys.argv[1:]],
            check=False,
        )
        sys.exit(0)
