import openpyxl

files = [
    ("/Users/design_euini/Downloads/3min_biz_english_marketer_v2.xlsx",  "marketer"),
    ("/Users/design_euini/Downloads/3min_biz_english_designer_v2.xlsx",  "designer"),
    ("/Users/design_euini/Downloads/3min_biz_english_developer_v2.xlsx", "developer"),
    ("/Users/design_euini/Downloads/3min_biz_english_pm_v2.xlsx",        "pm"),
]

def esc(val):
    if val is None:
        return "NULL"
    return "'" + str(val).replace("'", "''") + "'"

rows = []
for filepath, job in files:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        # cols: #, 난이도, 직군, 상황, 어색한표현(wrong1), 정답, 오답1(wrong2), 오답2(unused), 정답이유
        _, difficulty, _, situation, wrong1, correct, wrong2, _, reason = row
        is_premium = "false"
        rows.append(
            f"  ({esc(job)},{esc(difficulty)},{esc(situation)},{esc(correct)},NULL,{esc(wrong1)},{esc(wrong2)},{esc(reason)},{is_premium})"
        )

sql = "truncate table public.lessons restart identity;\n\n"
sql += "insert into public.lessons (job, difficulty, situation, correct, correct_ko, wrong1, wrong2, reason, is_premium) values\n"
sql += ",\n".join(rows)
sql += ";\n"

out = "/Users/design_euini/Desktop/buis-eng/supabase-seed.sql"
with open(out, "w", encoding="utf-8") as f:
    f.write(sql)

print(f"생성 완료: {out} ({len(rows)}개 rows)")
