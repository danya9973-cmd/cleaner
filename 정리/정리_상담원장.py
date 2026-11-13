"""상담원장 정리 — ../messy/상담원장_2025.xlsx → 정리본/

결정 기록: my-wiki/decisions/d-009-상담원장-정리의-다섯-결정.md
  d-009-1  «종결» = «완료», «대기» = «보류» — 합친다 (CS기획팀 답)
  d-009-2  처리시간 음수는 지우지 않는다. 값 그대로 + _표시 에 이유
  d-009-3  만족도 6~10 은 지우지 않는다. 값 그대로 + _표시 에 이유. 평균 만족도에서는 표시한 줄을 뺀다
  d-009-4  칸 이름은 원본 그대로 둔다
  d-009-5  고객번호 앞뒤 공백을 없앤다

[작성기준] 이라고 적은 처리는 원본 «작성기준» 시트에 적힌 것이다.
[결정 기록 없음] 이라고 적은 처리는 결정 기록이 아직 없다 — d-009 «아직 정하지 않은 것».

원본은 읽기만 한다. 결과: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
같은 원본이면 몇 번을 돌려도 세 파일이 똑같이 나온다.

    python 정리/정리_상담원장.py
"""
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent.parent          # cleaner/
SRC = HERE.parent / "messy" / "상담원장_2025.xlsx"
OUT = HERE / "정리본"
FLAG = "_표시"

raw = pd.read_excel(SRC, sheet_name="원장", dtype=object)
COLS = list(raw.columns)                                # d-009-4: 칸 이름은 원본 그대로
raw["_원본행"] = raw.index + 2                          # 엑셀 행 번호 (지운 이유에만 쓰고 결과에는 넣지 않는다)
summary = []                                            # (무엇을, 몇 줄, 근거)


def note(what, n, basis):
    summary.append((what, int(n), basis))


# ── 1. 앞뒤 공백 ───────────────────────────────────────────
df = raw.copy()
cust_ws = df["고객번호"].map(lambda v: isinstance(v, str) and v != v.strip())
other_ws = df[COLS].drop(columns="고객번호").map(lambda v: isinstance(v, str) and v != v.strip())
for c in COLS:
    df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) else v)

# ── 2. 접수일 표기 — [결정 기록 없음] ───────────────────────
df["접수일"] = pd.to_datetime(df["접수일"].str.replace(r"[/.]", "-", regex=True),
                              format="%Y-%m-%d").dt.strftime("%Y-%m-%d")

# ── 3. 지울 줄 ─────────────────────────────────────────────
reason = pd.Series("", index=df.index)
# 3-1 완전 중복 [작성기준 «12월 초 두 번 추출»] — 공백 · 날짜 표기를 맞춘 뒤 같은 줄, 앞 줄을 남긴다
first = {}
for i, k in df[COLS].map(str).agg("\x1f".join, axis=1).items():
    if k in first:
        reason[i] = f"완전 중복 — 날짜 표기·앞뒤 공백을 맞추면 원본 {df.at[first[k], '_원본행']}행과 같음"
    else:
        first[k] = i
# 3-2 테스트 접수 [작성기준 «AG999 는 테스트 접수»]
reason[(reason == "") & (df["상담원ID"] == "AG999")] = "테스트 접수 — 상담원ID AG999"
# 3-3 상태 변경 전 옛 줄 [작성기준 «최종수정일이 늦은 줄이 최신»]
left = df[reason == ""].sort_values(["접수번호", "최종수정일", "_원본행"])
latest = left.groupby("접수번호").tail(1)
latest_row = dict(zip(latest["접수번호"], latest["_원본행"]))
for i in left.index.difference(latest.index):
    reason[i] = f"상태 변경 전 옛 줄 — 같은 접수번호의 최신 줄은 원본 {latest_row[df.at[i, '접수번호']]}행"

removed = raw.loc[reason != "", COLS].copy()
removed["지운 이유"] = reason[reason != ""]
removed = removed.assign(_r=raw.loc[removed.index, "_원본행"]).sort_values("_r").drop(columns="_r")
for kind in ["완전 중복", "테스트 접수", "상태 변경 전 옛 줄"]:
    note(f"지운 줄 — {kind}", reason.str.startswith(kind).sum(), "작성기준")

df = df[reason == ""].sort_values(["접수번호"]).copy()
keep = df.index

# d-009-5 고객번호 공백 / [결정 기록 없음] 다른 칸 공백 — 남은 줄 기준으로 센다
note("고객번호 앞뒤 공백 지움 (d-009-5)", cust_ws[keep].sum(), "d-009-5")
for c in other_ws.columns:
    n = other_ws.loc[keep, c].sum()
    if n:
        note(f"{c} 앞뒤 공백 지움", n, "결정 기록 없음")
note("접수일 표기를 YYYY-MM-DD 로", (raw.loc[keep, "접수일"] != df["접수일"]).sum(), "결정 기록 없음")

# ── 4. 표기 통일 ───────────────────────────────────────────
def remap(col, mapping, label, basis):
    hit = df[col].isin(mapping.keys())
    note(label, hit.sum(), basis)
    df[col] = df[col].replace(mapping)

remap("처리상태", {"종결": "완료"}, "처리상태 «종결» → «완료» (d-009-1)", "d-009-1")
remap("처리상태", {"대기": "보류"}, "처리상태 «대기» → «보류» (d-009-1)", "d-009-1")
remap("접수채널", {"유선": "전화", "전화상담": "전화", "TEL": "전화",
                  "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
                  "모바일앱": "앱", "app": "앱", "APP": "앱",
                  "메일": "이메일", "e-mail": "이메일", "E-mail": "이메일"},
      "접수채널 → 전화·챗봇·앱·이메일", "결정 기록 없음")
team = df["상담원팀"].str.contains(" ", regex=False)
note("상담원팀 띄어쓰기 지움 («요금 팀» → «요금팀»)", team.sum(), "결정 기록 없음")
df["상담원팀"] = df["상담원팀"].str.replace(" ", "", regex=False)
remap("지역", {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"},
      "지역 «경기도»→«경기» · «서울시»·«서울특별시»→«서울»", "결정 기록 없음")
won = df["월정액"].map(lambda v: isinstance(v, str))
note("월정액 «33,000원» → 33000", won.sum(), "결정 기록 없음")
df["월정액"] = df["월정액"].map(lambda v: int(str(v).replace(",", "").replace("원", "")) if pd.notna(v) else v)
YES, NO = {"Y", "y", "예", "TRUE"}, {"N", "n", "아니오", "FALSE"}
rq = df["재문의"].map(lambda v: "Y" if v in YES else ("N" if v in NO else v))
note("재문의 → Y / N", (rq.astype(str) != df["재문의"].astype(str)).sum(), "결정 기록 없음")
df["재문의"] = rq

# ── 5. 상담시간 — 2025-07-01 이후 전화 채널은 초 → 분 [작성기준] ──
sec = (df["접수채널"] == "전화") & (df["접수일"] >= "2025-07-01")
df["상담시간"] = df["상담시간"].astype(float)
df.loc[sec, "상담시간"] = df.loc[sec, "상담시간"] / 60
note("상담시간: 2025-07-01 이후 전화 채널 줄을 초 → 분 (÷60)", sec.sum(), "작성기준")

# ── 6. 만족도 ──────────────────────────────────────────────
flag = pd.Series([[] for _ in range(len(df))], index=df.index)
m = df["만족도"].astype(str)
jeom = m.str.endswith("점")
note("만족도 «5점» → 5", jeom.sum(), "결정 기록 없음")
nonresp = m.isin(["0", "-", "미응답"])
note("만족도 «0»·«-»·«미응답» → 빈칸", nonresp.sum(), "작성기준")
sat = df["만족도"].map(lambda v: None if pd.isna(v) or str(v) in ("0", "-", "미응답")
                     else float(str(v).replace("점", "")))
out_sat = sat.notna() & ~sat.between(1, 5)              # d-009-3: 값은 그대로 두고 표시
for i in df.index[out_sat]:
    flag[i].append(f"만족도 범위 밖({int(sat[i])}) — NPS를 잘못 넣은 것으로 보이나 확인 못 함 · 평균 만족도에서 뺌 (d-009-3)")
note("만족도 6~10 — 값 그대로 + _표시 (d-009-3)", out_sat.sum(), "d-009-3")
df["만족도"] = sat

# ── 7. 처리시간 음수 — 값 그대로 + 표시 (d-009-2) ─────────────
neg = pd.to_numeric(df["처리시간"]) < 0
for i in df.index[neg]:
    flag[i].append("처리시간 음수 — 틀린 값이지만 진짜 값을 모름 · 지우면 «완료인데 처리시간 없음» 모순이 생김 (d-009-2)")
note("처리시간 음수 — 값 그대로 + _표시 (d-009-2)", neg.sum(), "d-009-2")

df[FLAG] = flag.map(" · ".join)
df = df[COLS + [FLAG]]                                   # d-009-4: 원본 칸 이름 · 순서 그대로, 맨 끝에 _표시

# ── 쓰기 ───────────────────────────────────────────────────
OUT.mkdir(exist_ok=True)
df.to_csv(OUT / "상담원장_정리.csv", index=False, encoding="utf-8-sig", lineterminator="\n")
removed.to_csv(OUT / "지운줄.csv", index=False, encoding="utf-8-sig", lineterminator="\n")

# ── 다섯 숫자 ──────────────────────────────────────────────
n = len(df)
post = df[df["접수일"] >= "2025-03-01"]
sat_ok = df.loc[df[FLAG].str.contains("만족도 범위 밖") == False, "만족도"]   # d-009-3: 표시한 줄은 평균에서 뺀다
done = df["처리상태"].isin(["완료", "처리완료"])           # «처리완료» = 완료 — 결정 기록 없음
five = [
    ("상담 건수", f"{n:,}건"),
    ("평균 상담시간", f"{df['상담시간'].mean():.2f}분"),
    ("재문의율", f"{(post['재문의'] == 'Y').mean() * 100:.2f}% (3/1 이후 {len(post):,}건 중 {(post['재문의'] == 'Y').sum():,}건)"),
    ("평균 만족도", f"{sat_ok.mean():.2f}점 (응답 {sat_ok.notna().sum():,}건)"),
    ("완료율", f"{done.mean() * 100:.2f}% ({done.sum():,} / {n:,})"),
    ("고객 수", f"{df['고객번호'].nunique():,}명"),
]

lines = ["# 상담원장 변경요약", "",
         f"원본 `../messy/상담원장_2025.xlsx` (`원장` 시트) {len(raw):,}줄 → 정리본 `상담원장_정리.csv` {n:,}줄 · 지운 줄 `지운줄.csv` {len(removed):,}줄",
         f"빼기 확인: {len(raw):,} − {len(removed):,} = {len(raw) - len(removed):,}", "",
         "결정 기록: `my-wiki/decisions/d-009-상담원장-정리의-다섯-결정.md`", "",
         "| 무엇을 | 몇 줄 | 근거 |", "|---|---|---|"]
lines += [f"| {w} | {k:,} | {b} |" for w, k, b in summary]
lines += ["", "- 줄 수는 정리본에 남은 줄 기준이다(지운 줄은 맨 위 세 줄).",
          "- **상담시간 칸의 단위**: 2025-07-01 이후 접수된 전화 채널 줄은 원본에 초 단위로 적혀 있어서 60으로 나눠 분으로 바꿨다. 정리본의 상담시간은 모든 줄이 분이다.",
          "- 근거가 «결정 기록 없음»인 처리는 d-009 «아직 정하지 않은 것»에 해당한다.", "",
          "## 이 정리본으로 낸 숫자", "", "| 숫자 | 값 |", "|---|---|"]
lines += [f"| {a} | {b} |" for a, b in five]
lines += ["", "- 완료율은 `완료`(«종결» 포함, d-009-1) · `처리완료` 를 완료로 센다. `처리완료` 를 완료로 보는 것은 결정 기록이 없다.",
          "- 평균 만족도는 `_표시` 에 «만족도 범위 밖»이 있는 줄을 뺀다 (d-009-3)."]
(OUT / "변경요약.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

for a, b in five:
    print(f"{a}\t{b}")
