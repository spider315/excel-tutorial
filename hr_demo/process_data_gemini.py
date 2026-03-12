
import pandas as pd
import re
import unicodedata
from pathlib import Path

# 設定路徑
BASE_DIR = Path(__file__).parent
RAW_DIR = BASE_DIR / "raw"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 讀取資料
roster_raw = pd.read_excel(RAW_DIR / "retirement_roster.xlsx", dtype=str)
nhi_raw = pd.read_excel(RAW_DIR / "nhi_transfer_list.xlsx", dtype=str)
payment_raw = pd.read_excel(RAW_DIR / "payment_notification.xlsx", dtype=str)

# 備份原始資料供異常檢查使用
roster_orig = roster_raw.copy()
nhi_orig = nhi_raw.copy()
payment_orig = payment_raw.copy()

# ==========================================
# 任務一：自動清理資料格式
# ==========================================

def clean_name(s):
    return str(s).strip() if pd.notna(s) else ""

def clean_date(s):
    if pd.isna(s) or str(s).strip() == "":
        return ""
    return str(s).strip().replace("/", "-")

def clean_tw_id(s):
    return str(s).strip().upper() if pd.notna(s) else ""

def clean_remark(s):
    if pd.isna(s) or str(s).strip() == "":
        return ""
    s = str(s)
    # 移除零寬字元 (\u200b) 和 BOM (\ufeff)
    return s.replace("\u200b", "").replace("\ufeff", "").strip()

# 清理退休名冊
roster = roster_raw.copy()
roster["姓名"] = roster["姓名"].apply(clean_name)
roster["身分證字號"] = roster["身分證字號"].apply(clean_tw_id)
for col in ["出生日期", "到職日", "預計退休日"]:
    roster[col] = roster[col].apply(clean_date)

# 清理健保轉出清單
nhi = nhi_raw.copy()
nhi["姓名"] = nhi["姓名"].apply(clean_name)
nhi["身分證字號"] = nhi["身分證字號"].apply(clean_tw_id)
nhi["轉出日期"] = nhi["轉出日期"].apply(clean_date)

# 清理給付通知
payment = payment_raw.copy()
payment["姓名"] = payment["姓名"].apply(clean_name)
payment["備註"] = payment.get("備註", pd.Series([""]*len(payment))).apply(clean_remark)

# 輸出清理後檔案
roster.to_excel(OUTPUT_DIR / "cleaned_retirement_roster.xlsx", index=False)
nhi.to_excel(OUTPUT_DIR / "cleaned_nhi_transfer_list.xlsx", index=False)
payment.to_excel(OUTPUT_DIR / "cleaned_payment_notification.xlsx", index=False)

# ==========================================
# 任務二：產出異常報告 anomaly_report.xlsx
# ==========================================

TW_ID_PATTERN = re.compile(r"^[A-Z][0-9]{9}$")
DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
BANK_CODE_PATTERN = re.compile(r"^\d{3}$")

# Sheet 1: 格式異常
format_errors = []

def add_format_err(src, row, col, val, desc):
    format_errors.append({"來源檔案": src, "列號": row, "欄位名稱": col, "原始值": val, "問題描述": desc})

# 檢查 Roster
for idx, row in roster_orig.iterrows():
    rnum = idx + 2
    tid = str(row["身分證字號"]).upper() if pd.notna(row["身分證字號"]) else ""
    if not TW_ID_PATTERN.match(tid):
        add_format_err("retirement_roster.xlsx", rnum, "身分證字號", row["身分證字號"], "身分證格式不符（1碼大寫英文+9碼數字）")
    for col in ["出生日期", "到職日", "預計退休日"]:
        dval = clean_date(row[col])
        if dval and not DATE_PATTERN.match(dval):
            add_format_err("retirement_roster.xlsx", rnum, col, row[col], "日期格式不符 YYYY-MM-DD")

# 檢查 NHI
for idx, row in nhi_orig.iterrows():
    rnum = idx + 2
    tid = str(row["身分證字號"]).upper() if pd.notna(row["身分證字號"]) else ""
    if not TW_ID_PATTERN.match(tid):
        add_format_err("nhi_transfer_list.xlsx", rnum, "身分證字號", row["身分證字號"], "身分證格式不符")
    try:
        amt = float(row["投保金額"])
        if amt < 0:
            add_format_err("nhi_transfer_list.xlsx", rnum, "投保金額", row["投保金額"], "投保金額為負數")
    except: pass

# 檢查 Payment
for idx, row in payment_orig.iterrows():
    rnum = idx + 2
    try:
        amt = float(row["應付金額"])
        if amt == 0:
            add_format_err("payment_notification.xlsx", rnum, "應付金額", row["應付金額"], "應付金額為 0")
    except: pass
    bank = str(row["銀行代碼"])
    if not BANK_CODE_PATTERN.match(bank):
        add_format_err("payment_notification.xlsx", rnum, "銀行代碼", row["銀行代碼"], "銀行代碼非 3 碼數字")
    acct = str(row["帳號"])
    if not (10 <= len(acct) <= 14):
        add_format_err("payment_notification.xlsx", rnum, "帳號", row["帳號"], f"帳號長度不在 10-14 碼（實際{len(acct)}碼）")

df_format = pd.DataFrame(format_errors)

# Sheet 2: 邏輯異常
logic_errors = []

for idx, row in roster_orig.iterrows():
    rnum = idx + 2
    d1, d2 = clean_date(row["到職日"]), clean_date(row["預計退休日"])
    if d1 and d2 and d2 < d1:
        logic_errors.append({"來源檔案": "retirement_roster.xlsx", "列號": rnum, "問題描述": "預計退休日早於到職日", "相關欄位值": f"到職:{d1}, 退休:{d2}"})

for idx, row in nhi_orig.iterrows():
    rnum = idx + 2
    if pd.isna(row["轉出日期"]) or str(row["轉出日期"]).strip() == "":
        logic_errors.append({"來源檔案": "nhi_transfer_list.xlsx", "列號": rnum, "問題描述": "轉出日期為空值", "相關欄位值": f"員工編號:{row['員工編號']}"})

df_logic = pd.DataFrame(logic_errors)

# Sheet 3: 跨檔比對異常
cross_errors = []
roster_names = dict(zip(roster["員工編號"], roster["姓名"]))
nhi_names = dict(zip(nhi["員工編號"], nhi["姓名"]))
pay_names = dict(zip(payment["員工編號"], payment["姓名"]))

all_ids = set(roster_names.keys()) | set(nhi_names.keys()) | set(pay_names.keys())
for eid in sorted(all_ids):
    names = []
    if eid in roster_names: names.append(("退休名冊", roster_names[eid]))
    if eid in nhi_names: names.append(("健保清單", nhi_names[eid]))
    if eid in pay_names: names.append(("給付通知", pay_names[eid]))
    
    unique_names = set(n for f, n in names)
    if len(unique_names) > 1:
        cross_errors.append({
            "員工編號": eid, 
            "問題描述": "跨檔姓名不一致", 
            "檔案一值": f"{names[0][0]}:{names[0][1]}", 
            "檔案二值": f"{names[1][0]}:{names[1][1]}"
        })

df_cross = pd.DataFrame(cross_errors)

# Sheet 4: 重複資料
dup_data = []
for df, name in [(roster_orig, "retirement_roster.xlsx"), (nhi_orig, "nhi_transfer_list.xlsx"), (payment_orig, "payment_notification.xlsx")]:
    counts = df["員工編號"].value_counts()
    dups = counts[counts > 1]
    for eid, count in dups.items():
        indices = [str(i+2) for i in df[df["員工編號"] == eid].index]
        dup_data.append({"來源檔案": name, "員工編號": eid, "出現次數": count, "相關列號": ", ".join(indices)})

df_dup = pd.DataFrame(dup_data)

with pd.ExcelWriter(OUTPUT_DIR / "anomaly_report.xlsx") as writer:
    df_format.to_excel(writer, sheet_name="格式異常", index=False)
    df_logic.to_excel(writer, sheet_name="邏輯異常", index=False)
    df_cross.to_excel(writer, sheet_name="跨檔比對異常", index=False)
    df_dup.to_excel(writer, sheet_name="重複資料", index=False)

# ==========================================
# 任務三：產出摘要報表 summary_report.xlsx
# ==========================================

# Sheet 1: 總覽儀表板
summary_stats = [
    {"項目": "退休名冊總筆數", "數值": len(roster_orig)},
    {"項目": "健保清單總筆數", "數值": len(nhi_orig)},
    {"項目": "給付通知總筆數", "數值": len(payment_orig)},
    {"項目": "格式異常總數", "數值": len(df_format)},
    {"項目": "邏輯異常總數", "數值": len(df_logic)},
    {"項目": "跨檔異常總數", "數值": len(df_cross)},
    {"項目": "重複資料總數", "數值": len(df_dup)},
    {"項目": "資料總筆數", "數值": len(roster_orig) + len(nhi_orig) + len(payment_orig)},
]
df_dashboard = pd.DataFrame(summary_stats)

# Sheet 2: 退休人員總表 (Inner Join)
# 先去重以確保 Join 結果乾淨
r_sub = roster.drop_duplicates(subset=["員工編號"])[["員工編號", "姓名", "身分證字號", "部門", "職等", "預計退休日"]]
n_sub = nhi.drop_duplicates(subset=["員工編號"])[["員工編號", "轉出日期", "投保金額"]]
p_sub = payment.drop_duplicates(subset=["員工編號"])[["員工編號", "給付類型", "應付金額"]]

master_list = r_sub.merge(n_sub, on="員工編號").merge(p_sub, on="員工編號")

# Sheet 3: 給付明細
master_list["應付金額"] = pd.to_numeric(master_list["應付金額"], errors="coerce").fillna(0)
dept_summary = master_list.groupby(["部門", "給付類型"])["應付金額"].sum().reset_index()
avg_payment = master_list["應付金額"].mean()

with pd.ExcelWriter(OUTPUT_DIR / "summary_report.xlsx") as writer:
    df_dashboard.to_excel(writer, sheet_name="總覽儀表板", index=False)
    master_list.to_excel(writer, sheet_name="退休人員總表", index=False)
    dept_summary.to_excel(writer, sheet_name="給付明細", index=False)
    # 寫入平均值到 Sheet 3 下方
    pd.DataFrame([{"部門": "全體平均給付金額", "應付金額": avg_payment}]).to_excel(writer, sheet_name="給付明細", startrow=len(dept_summary)+2, index=False)

# ==========================================
# 產出通知信 notification_letters.txt
# ==========================================

def fmt_date(d):
    if not d or d == "": return "（未定）"
    p = str(d).split("-")
    return f"{p[0]}年{p[1]}月{p[2]}日" if len(p) == 3 else d

letters = []
for _, row in master_list.iterrows():
    # 獲取發放日期
    pay_date_val = payment[payment["員工編號"] == row["員工編號"]]["發放日期"].values
    p_date = clean_date(pay_date_val[0]) if len(pay_date_val) > 0 else ""
    
    amt = f"{int(float(row['應付金額'])):,}"
    
    letter = f"""【退休離職通知】
{row['姓名']} 先生/女士 您好：
感謝您在本公司 {row['部門']} 服務多年。您的退休生效日為 {fmt_date(row['預計退休日'])}。
相關退休給付金額為新台幣 {amt} 元整，將於 {fmt_date(p_date)} 撥入您指定帳戶。
健保將於 {fmt_date(row['轉出日期'])} 辦理轉出。
如有任何疑問，請洽人力資源部。"""
    letters.append(letter)

(OUTPUT_DIR / "notification_letters.txt").write_text("\n---\n".join(letters), encoding="utf-8")

print("Processing complete.")
