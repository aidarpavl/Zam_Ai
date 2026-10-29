"""Zam AI — басшы орынбасарының цифрлық көмекшісі.

Іске қосу:  streamlit run app.py
"""
from __future__ import annotations

import datetime as dt
import json
import uuid

import pandas as pd
import streamlit as st

from zam_core.analyzers import PROFILE_NAMES, PROFILE_TO_TASK, detect_profile, run_profile, to_str, xnorm, xpct
from zam_core.excel_io import DEMO_ROWS, detect_header_row, read_file, scan_privacy, table_from_rows
from zam_core.prompts import (base_prompt, default_goal, default_theme, excel_prompt_for, excel_report_for,
                              file_stem, report_docx, table_to_tsv, to_csv_bytes)
from zam_core.tasks import GOLDEN, LEVELS, OUTPUT_OPTIONS, STATUSES, TASK_BY_ID, TASKS

st.set_page_config(page_title="Zam AI — мектеп менеджменті", page_icon="📊", layout="wide")

# ================================================================ Стиль
st.markdown("""
<style>
.block-container{padding-top:1.6rem;max-width:1500px}
.za-hero{background:linear-gradient(135deg,#10213e,#1e4ea9 72%,#3979ee);color:#fff;padding:32px 34px;border-radius:22px;margin-bottom:18px}
.za-hero h1{color:#fff;font-size:32px;line-height:1.15;margin:0 0 10px}
.za-hero p{color:#dbe8ff;font-size:16px;margin:0;max-width:850px}
.za-banner{background:linear-gradient(135deg,#f0f5ff,#fff);border:1px solid #cbdafa;border-radius:16px;padding:14px 18px;margin-bottom:16px}
.za-banner b{font-size:17px}.za-banner span{color:#6b7890}
.za-step{background:#fff;border:1px solid #dfe6f1;border-radius:14px;padding:13px;min-height:86px}
.za-step b{display:block;font-size:14px;margin-bottom:4px;color:#17233b}.za-step span{color:#6b7890;font-size:12px}
.za-num{display:inline-grid;place-items:center;width:30px;height:30px;border-radius:9px;background:#edf3ff;color:#1f5eff;font-weight:800;margin-right:8px}
.za-tag{display:inline-block;background:#edf3ff;color:#315aa6;border-radius:999px;padding:3px 9px;font-size:11px;font-weight:700}
.za-muted{color:#6b7890;font-size:13px}
.za-pilot{border-left:4px solid #1f5eff;padding:10px 14px;background:#f1f5ff;border-radius:0 12px 12px 0;color:#314a78;margin-bottom:14px}
.za-arrow{text-align:center;color:#91a0b9;font-size:22px;padding-top:26px}
</style>
""", unsafe_allow_html=True)

PAGES = {"home": "⌂ Басты бет", "dashboard": "◫ Жұмыс панелі", "catalog": "▦ Тапсырмалар каталогы",
         "builder": "✦ Промт конструкторы", "excel": "▥ Excel → талдау", "workflow": "↻ Басқарушылық цикл"}

ss = st.session_state


# ================================================================ Күй
def init_state():
    defaults = {
        "nav": "builder", "analyses": [], "decisions": [], "profile": {"school": "", "region": ""},
        "b_task": "2.1", "b_period": "", "b_scope": "", "b_output": OUTPUT_OPTIONS[0], "b_stage": "3",
        "b_extra": "", "b_data": "",
        "x_sheets": [], "x_file_name": "", "x_last_upload": None, "x_upload_ver": 0,
        "x_sheet": None, "x_header": 1, "x_profile": "auto", "x_period": "", "x_theme": "", "x_basis": "",
        "x_object": "", "x_goal": "", "x_analysis": None, "x_privacy": None, "x_prompt": "", "x_report": "",
        "x_batch": None, "x_error": "", "dec_ver": 0,
    }
    for k, v in defaults.items():
        ss.setdefault(k, v)


init_state()

# Streamlit бетте көрсетілмеген виджеттердің мәндерін өшіреді — беттер арасында ауысқанда
# толтырылған өрістер сақталуы үшін оларды әр іске қосуда қайта бекітеміз.
PERSIST_KEYS = ["nav", "b_task", "b_period", "b_scope", "b_output", "b_stage", "b_extra", "b_data", "x_sheet",
                "x_header", "x_profile", "x_period", "x_theme", "x_basis", "x_object", "x_goal"]
for _k in PERSIST_KEYS:
    if _k in ss:
        ss[_k] = ss[_k]


def go(page):
    ss.nav = page


def select_task(tid):
    ss.b_task = tid
    ss.nav = "builder"


def matrix_df(matrix, limit=None):
    if not matrix:
        return pd.DataFrame()
    head = [to_str(h) for h in matrix[0]]
    body = matrix[1:] if limit is None else matrix[1:limit + 1]
    rows = [[to_str(r[i]) if i < len(r) else "" for i in range(len(head))] for r in body]
    return pd.DataFrame(rows, columns=head)


def new_decision(**o):
    return {"id": "d" + uuid.uuid4().hex[:10], "problem": o.get("problem", ""), "action": o.get("action", ""),
            "responsible": o.get("responsible") or "[жауапты]", "deadline": o.get("deadline") or "[мерзімі]",
            "indicator": o.get("indicator") or "[күтілетін нәтиже/индикатор]",
            "control": o.get("control") or "Қайталама бақылау", "status": o.get("status") or "Жоспарланған",
            "createdAt": dt.datetime.now().isoformat(timespec="seconds")}


# ================================================================ Сайдбар
with st.sidebar:
    st.markdown("### 📊 Zam AI\n<span class='za-muted'>Мектеп менеджменті</span>", unsafe_allow_html=True)
    st.radio("Жұмыс кеңістігі", list(PAGES), format_func=PAGES.get, key="nav")
    st.divider()
    with st.expander("Мектеп профилі (промтқа қосылады)"):
        ss.profile["school"] = st.text_input("Мектеп", value=ss.profile.get("school", ""),
                                             placeholder="Мектеп атын жаз")
        ss.profile["region"] = st.text_input("Өңір", value=ss.profile.get("region", ""),
                                             placeholder="Мысалы: Павлодар облысы")
    st.caption("Деректер тек осы сессияда сақталады. Жұмысты сақтау үшін «Жұмыс панелі» бетінде "
               "JSON экспортын пайдаланыңыз.")


# ================================================================ Ортақ компоненттер
def flow(steps):
    cols = st.columns([4, 1] * (len(steps) - 1) + [4])
    for i, (t, d) in enumerate(steps):
        cols[i * 2].markdown(f"<div class='za-step'><b>{t}</b><span>{d}</span></div>", unsafe_allow_html=True)
        if i < len(steps) - 1:
            cols[i * 2 + 1].markdown("<div class='za-arrow'>→</div>", unsafe_allow_html=True)


def task_card(t, key_prefix):
    with st.container(border=True):
        st.markdown(f"**Промт {t['id']}** &nbsp; <span class='za-tag'>{t['category']}</span>", unsafe_allow_html=True)
        st.markdown(f"#### {t['title']}")
        st.markdown(f"<span class='za-muted'><b>Не жүктеу керек:</b> {t['upload']}<br><b>Нәтиже:</b> {t['result']}</span>",
                    unsafe_allow_html=True)
        st.caption("ЖИ рөлді, талдау ретін, нәтиже форматын және дәлелденбеген қорытындыларға тыйымды алады.")
        st.button("Пайдалану", key=f"{key_prefix}_use_{t['id']}", type="primary",
                  on_click=select_task, args=(t["id"],))
        with st.expander("Негізін көшіру"):
            st.code(base_prompt(t, output=OUTPUT_OPTIONS[0]), language=None, wrap_lines=True)


# ================================================================ Беттер
def page_home():
    st.markdown("<div class='za-banner'><b>Zam AI</b><br><span>Нақты сынақтан өткізуге арналған нұсқа: Excel-ді "
                "ақылды импорттау, 11 мамандандырылған талдау профилі, талдаулар мұрағаты, басқарушылық шешімдер "
                "тізілімі және қайталама бақылау.</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='za-hero'><h1>Алғашқы деректерден — негізделген басқарушылық шешімге дейін</h1>"
                "<p>Zam AI мектеп директорының орынбасарына жасанды интеллектке тапсырманы дұрыс қоюға көмектеседі: "
                "деректерді тексеру, есептеулер жүргізу, фактіні болжамнан ажырату, шешім мен құжат жобасын дайындау.</p></div>",
                unsafe_allow_html=True)
    c = st.columns(3)
    c[0].button("Excel-ді талдау", on_click=go, args=("excel",), width="stretch")
    c[1].button("Промт жасау", on_click=go, args=("builder",), width="stretch")
    c[2].button("Басқарушылық тапсырманы таңдау", on_click=go, args=("catalog",), type="primary",
                width="stretch")

    k = st.columns(4)
    k[0].metric("Тапсырмалар банкі", len(TASKS), help="мектеп менеджментінің негізгі бағыты")
    k[1].metric("Құжатты тексеру", 16, help="сараптамалық чек-лист өлшемі")
    k[2].metric("Жұмыс нысандары", 4, help="МІБ картасы, шешімдер, қайталама бақылау, аудит")
    k[3].metric("Excel-профильдер", 11, help="автоматты талдағыш + әмбебап режим")

    st.subheader("Басқарушылық цикл")
    st.caption("«Анықтама жазу» емес, дәлелді циклден өту.")
    flow([("1. Деректер", "Кесте, бақылау, тізілім, нәтижелер"), ("2. Салыстыру", "Кезеңдер, сыныптар, пәндер, топтар"),
          ("3. Ауытқу", "Нақты назар аударуды қажет ететін мәселе"),
          ("4. Шешім", "Әрекет, жауапты, мерзім, индикатор, бақылау")])

    st.subheader("Жылдам бастау")
    st.caption("Директор орынбасарының ең жиі кездесетін тапсырмалары.")
    cols = st.columns(3)
    for col, tid in zip(cols, ["2.1", "2.3", "8.1"]):
        with col:
            task_card(TASK_BY_ID[tid], "home")

    st.subheader("Талдамалық анықтаманың «Алтын қаңқасы»")
    st.caption("Сапалы құжат құрылымын айқындайтын 12 элемент.")
    with st.container(border=True):
        cols = st.columns(2)
        for i, g in enumerate(GOLDEN):
            cols[i % 2].markdown(f"<span class='za-num'>{i + 1}</span><b>{g}</b> "
                                 f"<span class='za-muted'>— құрылымның міндетті элементі</span>", unsafe_allow_html=True)


def page_catalog():
    st.header("Басқарушылық тапсырмалар каталогы")
    st.caption("Тапсырманы таңдаңыз — конструктор рөлді, талдау логикасын, күтілетін нәтижені және шектеулерді автоматты қояды.")
    c1, c2 = st.columns([3, 1])
    q = c1.text_input("Іздеу", placeholder="Іздеу: үлгерім, МІБ, тамақтану, БЖБ/ТЖБ...", label_visibility="collapsed")
    cats = ["Барлық бағыттар"] + sorted({t["category"] for t in TASKS})
    cat = c2.selectbox("Бағыт", cats, label_visibility="collapsed")
    ql = q.strip().lower()
    arr = [t for t in TASKS if (cat == cats[0] or t["category"] == cat)
           and (not ql or ql in " ".join([t["title"], t["category"], t["upload"], t["result"], t["id"]]).lower())]
    if not arr:
        st.info("Берілген шарттар бойынша ештеңе табылмады.")
        return
    for i in range(0, len(arr), 3):
        cols = st.columns(3)
        for col, t in zip(cols, arr[i:i + 3]):
            with col:
                task_card(t, "cat")


def load_builder_example():
    ss.b_task = "2.1"
    ss.b_period = "2026–2027 оқу жылының I тоқсаны"
    ss.b_scope = "5-сыныптар"
    ss.b_data = ("Сынып\tN\t5\t4\t3\t2\tБір 3-і бар\n5А\t27\t4\t12\t11\t0\t3\n"
                 "5Б\t26\t2\t8\t15\t1\t6\n5В\t25\t5\t11\t9\t0\t2")


def page_builder():
    st.header("Промт конструкторы")
    st.caption("Сайт сұрауды «Рөл–Нысан–Логика–Нәтиже–Контекст» моделі бойынша жинап, дәлелді талдау ережелерін қосады.")
    left, right = st.columns([1.05, .95], gap="large")
    with left, st.container(border=True):
        st.selectbox("Басқарушылық тапсырма", [t["id"] for t in TASKS],
                     format_func=lambda i: f"{i} · {TASK_BY_ID[i]['title']}", key="b_task")
        c1, c2 = st.columns(2)
        c1.text_input("Кезең", key="b_period", placeholder="Мысалы: 2026–2027 оқу жылының I тоқсаны")
        c2.text_input("Талдау нысаны", key="b_scope", placeholder="Мысалы: 5–9 сыныптар / 7-сыныптар параллелі")
        c1, c2 = st.columns(2)
        c1.selectbox("Не алу керек", OUTPUT_OPTIONS, key="b_output")
        c2.selectbox("ЖИ жұмыс режимі", ["3", "1"], key="b_stage",
                     format_func={"3": "3 кезең: аудитор → талдаушы → директор орынбасары",
                                  "1": "Бір кешенді сұрау"}.get)
        st.text_input("Қосымша шарттар", key="b_extra",
                      placeholder="Мысалы: алдыңғы тоқсанмен салыстыру, сапа резервін бөліп көрсету")
        st.text_area("Алғашқы деректер", key="b_data", height=200,
                     placeholder="Excel/Google Sheets кестесін, CSV немесе мәтіндік деректерді қойыңыз.")
        st.error("")
        st.button("Мысал жүктеу", on_click=load_builder_example)
    with right:
        prompt = base_prompt(TASK_BY_ID[ss.b_task], ss.b_period.strip(), ss.b_scope.strip(), ss.b_extra.strip(),
                             ss.b_output, ss.b_stage == "3", ss.b_data.strip(), ss.profile)
        st.markdown("**Дайын промт** <span class='za-muted'>(көшіру батырмасы — блоктың жоғарғы оң жағында)</span>",
                    unsafe_allow_html=True)
        st.code(prompt, language=None, wrap_lines=True, height=620)
        st.download_button("Промтты TXT жүктеу", prompt.encode("utf-8"), file_name=f"ZamAI_prompt_{ss.b_task}.txt",
                           mime="text/plain")


# ---------------------------------------------------------------- Excel
def current_sheet():
    for sh in ss.x_sheets:
        if sh["name"] == ss.x_sheet:
            return sh
    return ss.x_sheets[0] if ss.x_sheets else None


def current_table():
    sh = current_sheet()
    if not sh:
        return None
    hr = max(0, int(ss.x_header or 1) - 1)
    if hr >= len(sh["rows"]):
        hr = 0
    return table_from_rows(sh["rows"], hr)


def _clear_results():
    ss.x_analysis = ss.x_privacy = ss.x_batch = None
    ss.x_prompt = ss.x_report = ""


def _sync_theme():
    t = current_table()
    if t is not None and ss.x_profile == "auto":
        ss.x_theme = default_theme(detect_profile(t["headers"]))


def load_sheets(sheets, name):
    ss.x_sheets = sheets
    ss.x_file_name = name
    ss.x_sheet = sheets[0]["name"]
    ss.x_header = detect_header_row(sheets[0]["rows"]) + 1
    ss.x_profile = "auto"
    ss.x_error = ""
    _clear_results()
    _sync_theme()


def on_sheet_change():
    sh = current_sheet()
    if sh:
        ss.x_header = detect_header_row(sh["rows"]) + 1
    _sync_theme()


def on_profile_change():
    t = current_table()
    p = detect_profile(t["headers"]) if (ss.x_profile == "auto" and t) else ss.x_profile
    ss.x_theme = default_theme(p if p != "auto" else "generic")
    ss.x_goal = default_goal(p if p != "auto" else "generic")


def run_analysis():
    t = current_table()
    if t is None:
        return
    profile = detect_profile(t["headers"]) if ss.x_profile == "auto" else ss.x_profile
    if not ss.x_theme.strip():
        ss.x_theme = default_theme(profile)
    if not ss.x_goal.strip():
        ss.x_goal = default_goal(profile)
    a = run_profile(t, profile)
    meta = {"period": ss.x_period.strip(), "theme": ss.x_theme.strip(), "object": ss.x_object.strip(),
            "goal": ss.x_goal.strip(), "basis": ss.x_basis.strip()}
    ss.x_analysis = a
    ss.x_privacy = scan_privacy(t)
    ss.x_prompt = excel_prompt_for(a, meta, ss.profile)
    ss.x_report = excel_report_for(a, meta)


def load_demo():
    load_sheets([{"name": "Білім сапасы", "rows": [list(r) for r in DEMO_ROWS]}], "Нұсқаулықтан демо")
    ss.x_period, ss.x_object = "I тоқсан", "5-сыныптар"
    ss.x_basis = "МІБ жоспары / [негіздемені көрсету]"
    run_analysis()


def reset_excel():
    ss.x_sheets, ss.x_file_name, ss.x_sheet, ss.x_last_upload = [], "", None, None
    ss.x_upload_ver += 1
    for k in ("x_period", "x_theme", "x_basis", "x_object", "x_goal", "x_error"):
        ss[k] = ""
    _clear_results()


def analyze_all_sheets():
    out = []
    for sh in ss.x_sheets:
        hi = detect_header_row(sh["rows"])
        t = table_from_rows(sh["rows"], hi)
        p = detect_profile(t["headers"])
        a = run_profile(t, p)
        out.append([sh["name"], PROFILE_NAMES.get(p, p), len(t["rows"]), to_str(a["kpis"][0][1]) if a["kpis"] else "—",
                    sum(1 for x in a["audit"] if x["level"] == "error"),
                    sum(1 for x in a["audit"] if x["level"] == "warn")])
    ss.x_batch = [["Парақ", "Автопрофиль", "Жолдар", "Негізгі көрсеткіш", "Қателер", "Ескертулер"], *out]


def excel_to_builder():
    a = ss.x_analysis
    if not a:
        return
    if a["profile"] in PROFILE_TO_TASK:
        ss.b_task = PROFILE_TO_TASK[a["profile"]]
    ss.b_period, ss.b_scope = ss.x_period, ss.x_object
    ss.b_data = table_to_tsv(a["table"])
    ss.b_extra = "Жергілікті есептелген кестені пайдалан. Себептерді ойлап таппа. Әр қорытындыны нақты көрсеткіш бойынша тексер."
    go("builder")


def save_analysis():
    a = ss.x_analysis
    if not a:
        return
    ss.analyses.insert(0, {"id": "a" + uuid.uuid4().hex[:10], "createdAt": dt.datetime.now().isoformat(timespec="seconds"),
                           "profile": a["profile"], "title": ss.x_theme or PROFILE_NAMES[a["profile"]],
                           "fileName": ss.x_file_name, "sheetName": ss.x_sheet, "period": ss.x_period,
                           "object": ss.x_object, "facts": a["facts"], "problems": a["problems"],
                           "decisions": a["decisions"], "report": ss.x_report, "prompt": ss.x_prompt})
    ss.analyses = ss.analyses[:60]
    st.toast("Талдау сақталды")


def analysis_to_decisions():
    a = ss.x_analysis
    if not a:
        return
    for i, d in enumerate(a["decisions"]):
        prob = a["problems"][i] if i < len(a["problems"]) else (ss.x_theme or PROFILE_NAMES[a["profile"]])
        ss.decisions.insert(0, new_decision(problem=prob, action=d))
    ss.dec_ver += 1
    go("dashboard")


def page_excel():
    st.header("Excel → автоматты талдау → промт → талдамалық анықтама")
    st.caption("XLSX немесе CSV жүктеңіз. Сайт құрылымды тексереді, көрсеткіштерді есептейді, дәлелденген ауытқуларды "
               "бөледі және екі нәтиже жасайды — ЖИ үшін қауіпсіз промт және МІБ анықтамасының жобасы.")

    with st.container(border=True):
        up = st.file_uploader("Excel-файлды осында сүйреңіз (XLSX немесе CSV)", type=["xlsx", "xlsm", "csv"],
                              key=f"x_upload_{ss.x_upload_ver}")
        b1, b2 = st.columns([1, 5])
        b1.button("Демо: білім сапасы", on_click=load_demo)
        if up is not None:
            fid = (up.name, up.size)
            if ss.x_last_upload != fid:
                ss.x_last_upload = fid
                try:
                    load_sheets(read_file(up.name, up.getvalue()), up.name)
                except Exception as e:  # noqa: BLE001
                    ss.x_sheets = []
                    ss.x_error = f"Файлды оқу мүмкін болмады: {e}"
        if ss.x_error:
            st.error(ss.x_error)
        chips = [f"📄 {ss.x_file_name or 'Файл таңдалмаған'}",
                 f"Оқылған парақтар: {len(ss.x_sheets)}" if ss.x_sheets else "Деректер күтілуде"]
        if ss.x_privacy:
            chips.append("🔴 ДД тәуекелі анықталды" if ss.x_privacy["hasRisk"] else "🟢 Анық ДД табылмады")
        else:
            chips.append("ЖД тексерілмеген")
        b2.caption(" · ".join(chips))
        st.caption("Ескерту: Streamlit Cloud-та файл қолданба серверінде өңделеді және сақталмайды. Толық жергілікті "
                   "өңдеу үшін қолданбаны өз компьютеріңізде іске қосыңыз.")

    if not ss.x_sheets:
        return

    t = current_table()
    auto = detect_profile(t["headers"]) if t else "generic"
    with st.container(border=True):
        h1, h2 = st.columns([3, 2])
        h1.subheader("1. Талдауды баптау")
        h2.info(f"Автоанықтау: **{PROFILE_NAMES[auto]}**")
        c = st.columns(4)
        c[0].selectbox("Excel парағы", [s["name"] for s in ss.x_sheets], key="x_sheet", on_change=on_sheet_change)
        c[1].number_input("Тақырып жолы", min_value=1, step=1, key="x_header", on_change=_sync_theme)
        c[2].selectbox("Талдау түрі", ["auto", *PROFILE_NAMES], key="x_profile", on_change=on_profile_change,
                       format_func=lambda p: "Автоматты анықтау" if p == "auto" else PROFILE_NAMES[p])
        c[3].text_input("Кезең", key="x_period", placeholder="2026–2027 оқу жылының I тоқсаны")
        c = st.columns(3)
        c[0].text_input("Анықтама тақырыбы", key="x_theme", placeholder="Үлгерім қорытындыларын талдау")
        c[1].text_input("Негіздеме", key="x_basis", placeholder="МІБ жоспары / №__ бұйрық, __ күні")
        c[2].text_input("Талдау нысаны", key="x_object", placeholder="Мысалы: 5-сыныптар")
        st.text_input("Бақылау мақсаты", key="x_goal",
                      placeholder="Нақты нәтижелерді, ауытқуларды және келесі бақылау аймақтарын анықтау")
        c = st.columns([2, 2, 1, 3])
        c[0].button("Автоматты талдауды іске қосу", type="primary", on_click=run_analysis, width="stretch")
        c[1].button("Барлық парақты талдау", on_click=analyze_all_sheets, width="stretch")
        c[2].button("Тазалау", on_click=reset_excel, width="stretch")
        st.info("**Дәлелдік принципі:** модуль себептерді ойлап таппайды. Ол тек кестеден есептелген фактілерді "
                "тұжырымдайды. Себептер берілмесе, анықтамада қосымша зерттеу қажет екені көрсетіледі.")

    if ss.x_batch:
        with st.container(border=True):
            st.subheader(f"Кітапты пакеттік талдау · парақтар: {len(ss.x_batch) - 1}")
            st.dataframe(matrix_df(ss.x_batch), hide_index=True, width="stretch")

    if t is not None:
        with st.expander(f"Бастапқы деректерді алдын ала қарау — {len(t['rows'])} жол × {len(t['headers'])} баған · "
                         f"тақырып: {t['header_index'] + 1}-жол", expanded=ss.x_analysis is None):
            st.dataframe(matrix_df([t["headers"], *t["rows"]], 20), hide_index=True, width="stretch")

    a = ss.x_analysis
    if not a:
        return
    render_results(a, ss.x_privacy)


def render_results(a, privacy):
    err = sum(1 for x in a["audit"] if x["level"] == "error")
    warn = sum(1 for x in a["audit"] if x["level"] == "warn")
    st.subheader("2. Аудит және есептеулер")
    st.caption(f"Алдымен деректерді тексеру, содан кейін талдау. Қателер: {err} · ескертулер: {warn}")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        st.markdown("**Деректер аудиті**")
        for x in a["audit"]:
            {"error": st.error, "warn": st.warning}.get(x["level"], st.success)(x["text"])
    with c2, st.container(border=True):
        st.markdown("**Деректерді қорғау**")
        if privacy["hasRisk"]:
            parts = []
            if privacy["badHeaders"]:
                parts.append(f"Өрістер: {', '.join(privacy['badHeaders'])}.")
            if privacy["iin"]:
                parts.append(f"ЖСН-ге ұқсас мәндер: {privacy['iin']}.")
            if privacy["email"]:
                parts.append(f"Email: {privacy['email']}.")
            if privacy["phone"]:
                parts.append(f"Телефондар: {privacy['phone']}.")
            st.error("**Дербес деректер тәуекелі бар.**\n\n" + " ".join(parts) +
                     "\n\nДайын промтқа бастапқы жолдар қосылмайды: тек жиынтық есептеулер қолданылады. "
                     "Сыртқы ЖИ-ге жібермес бұрын бәрібір жасырындандыруды орындаңыз.")
        else:
            st.success("**Экспресс-тексеру анық ЖСН, телефон, email немесе қауіпті өріс атауларын таппады.**\n\n"
                       "Бұл дербес және құпия деректерді қолмен тексеруді алмастырмайды.")

    st.subheader("3. Негізгі көрсеткіштер")
    cols = st.columns(len(a["kpis"]))
    for col, (k, v) in zip(cols, a["kpis"]):
        col.metric(k, to_str(v))

    with st.container(border=True):
        h1, h2 = st.columns([4, 1])
        h1.markdown("**Есептеу кестесі**")
        h2.download_button("Есептеулерді CSV жүктеу", to_csv_bytes(a["table"]),
                           file_name=file_stem(a["profile"]) + "_esepteuler.csv", mime="text/csv",
                           width="stretch")
        st.dataframe(matrix_df(a["table"], 300), hide_index=True, width="stretch")

    st.subheader("4. Дәлелденген қорытындылар")
    st.caption("Факт болжамнан бөлінген; басқарушылық шаралар анықталған ауытқуларға байланған.")
    cols = st.columns(3)
    for col, (title, items) in zip(cols, [("Фактілер", a["facts"]), ("Оң үрдістер", a["positive"]),
                                          ("Назар аударатын аймақтар", a["problems"])]):
        with col, st.container(border=True):
            st.markdown(f"**{title}**")
            st.markdown("\n".join(f"- {x}" for x in (items or ["Автоматты тұжырымдалған тармақтар жоқ."])))

    st.subheader("5. Дайын промт және талдамалық анықтама")
    st.caption("Промтқа бастапқы файл емес, есептелген жиынтық деректер беріледі.")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        st.markdown("**ЖИ үшін дайын промт**")
        b = st.columns(2)
        b[0].button("Конструкторға", on_click=excel_to_builder, width="stretch")
        b[1].button("Талдауды сақтау", on_click=save_analysis, width="stretch")
        st.code(ss.x_prompt, language=None, wrap_lines=True, height=560)
    with c2, st.container(border=True):
        st.markdown("**Талдамалық анықтама жобасы**")
        b = st.columns(3)
        b[0].button("Шешімдер тізіліміне", on_click=analysis_to_decisions, width="stretch")
        stem = file_stem(a["profile"])
        b[1].download_button("DOCX жүктеу", report_docx(ss.x_report, a["table"]), file_name=stem + "_anyqtama.docx",
                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                             width="stretch")
        b[2].download_button("TXT жүктеу", ss.x_report.encode("utf-8"), file_name=stem + "_anyqtama.txt",
                             mime="text/plain", width="stretch", type="primary")
        st.code(ss.x_report, language=None, wrap_lines=True, height=560)
    st.error("**Қол қою алдында:** нормативтік негіздемелерді, дербес деректерді, есептеулерді, себеп-салдарлық "
             "тұжырымдарды және тапсырмалардың нақтылығын тексеріңіз. ЖИ және автоматты модуль жоба қалыптастырады, "
             "директор орынбасарының шешімін алмастырмайды.")


# ---------------------------------------------------------------- Пилоттық панель
def delete_analysis(aid):
    ss.analyses = [x for x in ss.analyses if x["id"] != aid]


def import_workspace(data: dict):
    if isinstance(data.get("analyses"), list):
        known = {x["id"] for x in ss.analyses}
        ss.analyses = ss.analyses + [x for x in data["analyses"] if x.get("id") not in known]
    elif isinstance(data, list):  # бұрынғы нұсқаның талдаулар экспорты
        ss.analyses = ss.analyses + data
    if isinstance(data, dict) and isinstance(data.get("decisions"), list):
        known = {x["id"] for x in ss.decisions}
        ss.decisions = ss.decisions + [x for x in data["decisions"] if x.get("id") not in known]
    if isinstance(data, dict) and isinstance(data.get("profile"), dict):
        ss.profile.update(data["profile"])
    ss.dec_ver += 1


def page_dashboard():
    st.header("Пилоттық панель")
    st.markdown("<div class='za-pilot'>Панель мектептің ақпараттық жүйесі болып табылмайды. Ол тек пайдаланушы осы "
                "сессияда сақтаған жасырындандырылған деректер мен жобаларды көрсетеді. Сессия жабылғанда деректер "
                "жойылады — жұмыс кеңістігін JSON түрінде экспорттап, кейін импорттаңыз.</div>", unsafe_allow_html=True)
    dd = ss.decisions
    open_n = sum(1 for d in dd if d["status"] != "Орындалды")
    k = st.columns(4)
    k[0].metric("Сақталған талдаулар", len(ss.analyses))
    k[1].metric("Басқарушылық шешімдер", len(dd))
    k[2].metric("Ашық шешімдер", open_n)
    k[3].metric("Каталогтағы тапсырмалар", len(TASKS))

    ws = {"version": 1, "exportedAt": dt.datetime.now().isoformat(timespec="seconds"),
          "profile": ss.profile, "analyses": ss.analyses, "decisions": ss.decisions}
    with st.container(border=True):
        c1, c2 = st.columns([1, 2])
        c1.download_button("Жұмыс кеңістігін JSON экспорттау",
                           json.dumps(ws, ensure_ascii=False, indent=2).encode("utf-8"),
                           file_name=f"ZamAI_workspace_{dt.date.today().isoformat()}.json",
                           mime="application/json", width="stretch")
        imp = c2.file_uploader("JSON импорттау", type=["json"], key=f"ws_import_{ss.dec_ver}",
                               label_visibility="collapsed")
        if imp is not None:
            try:
                import_workspace(json.loads(imp.getvalue().decode("utf-8-sig")))
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"JSON оқылмады: {e}")

    left, right = st.columns([1.1, .9])
    with left, st.container(border=True):
        h1, h2 = st.columns([3, 1])
        h1.markdown("**Соңғы талдаулар**")
        if ss.analyses and h2.button("Тазалау", width="stretch"):
            ss.analyses = []
            st.rerun()
        if not ss.analyses:
            st.caption("Талдаулар әзірге жоқ.")
        for a in ss.analyses[:8]:
            with st.expander(f"{a['title']} · {a['createdAt'].replace('T', ' ')} · "
                             f"{PROFILE_NAMES.get(a['profile'], a['profile'])}"):
                st.caption(" · ".join(x for x in (a.get("period"), a.get("object"), a.get("fileName")) if x))
                tab1, tab2 = st.tabs(["Талдамалық анықтама", "Промт"])
                tab1.code(a.get("report", ""), language=None, wrap_lines=True)
                tab2.code(a.get("prompt", ""), language=None, wrap_lines=True)
                st.button("Жою", key=f"del_{a['id']}", on_click=delete_analysis, args=(a["id"],))
    with right, st.container(border=True):
        st.markdown("**Орындаушылық тәртіп**")
        done = sum(1 for d in dd if d["status"] == "Орындалды")
        p = done / len(dd) * 100 if dd else 0
        st.markdown(f"<div style='font-size:32px;font-weight:800'>{xpct(p, 0)}</div>", unsafe_allow_html=True)
        st.caption(f"{len(dd)}-ның {done}-ы орындалды")
        st.progress(min(1.0, p / 100))
        st.caption(f"Ашық: {open_n}. Мәртебе расталған бақылау нәтижелері бойынша қолмен жаңартылады.")

    render_decisions()


def render_decisions():
    with st.container(border=True):
        st.markdown("**Басқарушылық шешімдер тізілімі**")
        st.caption("Мәселе → әрекет → жауапты → мерзім → индикатор → қайталама бақылау.")
        with st.expander("➕ Шешім қосу"):
            with st.form("add_decision", clear_on_submit=True):
                problem = st.text_input("Мәселе / шешім негіздемесі *")
                action = st.text_input("Нақты әрекет *")
                c = st.columns(3)
                resp = c[0].text_input("Жауапты", placeholder="[жауапты]")
                dl = c[1].text_input("Мерзімі", placeholder="[мерзімі]")
                status = c[2].selectbox("Мәртебе", STATUSES)
                ind = st.text_input("Күтілетін нәтиже / индикатор", placeholder="[индикатор]")
                ctrl = st.text_input("Бақылау нысаны", value="Қайталама бақылау")
                if st.form_submit_button("Қосу", type="primary"):
                    if problem.strip() and action.strip():
                        ss.decisions.insert(0, new_decision(problem=problem.strip(), action=action.strip(),
                                                            responsible=resp.strip(), deadline=dl.strip(),
                                                            indicator=ind.strip(), control=ctrl.strip(),
                                                            status=status))
                        ss.dec_ver += 1
                        st.rerun()
                    else:
                        st.warning("Мәселе мен әрекетті толтырыңыз.")

        c1, c2, c3 = st.columns([3, 2, 1])
        q = c1.text_input("Іздеу", placeholder="Мәселе, әрекет, жауапты бойынша іздеу...", label_visibility="collapsed")
        f = c2.selectbox("Мәртебе", ["Барлық мәртебелер", *STATUSES], label_visibility="collapsed")
        rows = [["Мәселе", "Шешім", "Жауапты", "Мерзімі", "Индикатор", "Бақылау нысаны", "Мәртебе"],
                *[[d["problem"], d["action"], d["responsible"], d["deadline"], d["indicator"], d["control"], d["status"]]
                  for d in ss.decisions]]
        c3.download_button("CSV жүктеу", to_csv_bytes(rows),
                           file_name=f"ZamAI_sheshimder_{dt.date.today().isoformat()}.csv", mime="text/csv",
                           width="stretch", disabled=not ss.decisions)

        qn = xnorm(q)
        view = [d for d in ss.decisions if (f == "Барлық мәртебелер" or d["status"] == f)
                and (not qn or qn in xnorm(" ".join([d["problem"], d["action"], d["responsible"]])))]
        if not view:
            st.caption("Таңдалған сүзгі бойынша шешімдер жоқ.")
            return
        df = pd.DataFrame([{"id": d["id"], "Жою": False, "Мәселе": d["problem"], "Әрекет": d["action"],
                            "Жауапты": d["responsible"], "Мерзімі": d["deadline"], "Индикатор": d["indicator"],
                            "Бақылау": d["control"], "Мәртебе": d["status"],
                            "Құрылды": d["createdAt"][:10]} for d in view]).set_index("id")
        edited = st.data_editor(
            df, hide_index=True, width="stretch", key=f"dec_editor_{ss.dec_ver}",
            column_config={"Мәртебе": st.column_config.SelectboxColumn("Мәртебе", options=STATUSES, required=True),
                           "Жою": st.column_config.CheckboxColumn("Жою", width="small"),
                           "Құрылды": st.column_config.TextColumn("Құрылды", disabled=True)})
        if st.button("Өзгерістерді сақтау", type="primary"):
            by_id = {d["id"]: d for d in ss.decisions}
            to_del = set()
            for did, r in edited.iterrows():
                d = by_id.get(did)
                if not d:
                    continue
                if r["Жою"]:
                    to_del.add(did)
                    continue
                d.update(problem=r["Мәселе"], action=r["Әрекет"], responsible=r["Жауапты"], deadline=r["Мерзімі"],
                         indicator=r["Индикатор"], control=r["Бақылау"], status=r["Мәртебе"])
            ss.decisions = [d for d in ss.decisions if d["id"] not in to_del]
            ss.dec_ver += 1
            st.rerun()


# ---------------------------------------------------------------- Цикл
def page_workflow():
    st.header("ЖИ-мен жұмыстың үш кезеңді технологиясы")
    st.caption("Әр кезеңнің жеке рөлі және тексерілетін нәтижесі бар.")
    stages = [("1-кезең", "Деректер аудиторы",
               "Толықтықты, арифметиканы, логикалық қайшылықтарды, жетіспейтін өрістерді тексереді. Құжатты әзірге жазбайды.",
               "Нәтиже: қателер, олқылықтар және деректер шектеулері туралы есеп.", st.info),
              ("2-кезең", "Талдаушы",
               "Көрсеткіштерді қайта есептейді, сыныптар/кезеңдерді салыстырады, ауытқулар мен назар аймақтарын анықтайды. Себептерді ойлап таппайды.",
               "Нәтиже: есептеулер кестесі, салыстырулар және дәлелді қорытындылар.", st.info),
              ("3-кезең", "Директор орынбасары",
               "Расталған нәтижелер негізінде ғана құжат жобасын және басқарушылық шешімдерді қалыптастырады.",
               "Нәтиже: анықтама жобасы + мекенжайлы шешімдер + қайталама бақылау.", st.success)]
    for col, (tag, title, text, res, box) in zip(st.columns(3), stages):
        with col, st.container(border=True):
            st.markdown(f"<span class='za-tag'>{tag}</span>", unsafe_allow_html=True)
            st.markdown(f"### {title}")
            st.write(text)
            box(res)
    st.subheader("Басқарушылық шешім формуласы")
    with st.container(border=True):
        flow([("Не істеу керек?", "Нақты әрекет"), ("Кім?", "Жауапты"), ("Қай мерзімге?", "Тексерілетін күн"),
              ("Қалай тексереміз?", "Индикатор және қайталама бақылау нысаны")])
    st.subheader("Дәлелді талдаудың 5 деңгейі")
    with st.container(border=True):
        cols = st.columns(2)
        for i, (n, t, d) in enumerate(LEVELS):
            cols[i % 2].markdown(f"<span class='za-num'>{n}</span><b>{t}</b> <span class='za-muted'>— {d}</span>",
                                 unsafe_allow_html=True)


# ================================================================ Маршрутизация
{"home": page_home, "dashboard": page_dashboard, "catalog": page_catalog, "builder": page_builder,
 "excel": page_excel, "workflow": page_workflow}[ss.nav]()

st.divider()
st.caption("Zam AI · мектеп менеджменті")
