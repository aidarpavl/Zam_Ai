"""Excel-деректерді талдау: бағандарды тану, профильді анықтау, 11 мамандандырылған талдағыш."""
from __future__ import annotations

import datetime as _dt
import math
import re

PROFILE_NAMES = {
    "grades": "Білім сапасы және үлгерім",
    "diagnostic": "Диагностикалық жұмыстар",
    "sor": "БЖБ / ТЖБ",
    "attendance": "Жалпыға бірдей білім беру және қатысу",
    "correction": "Олқылықтарды толтыру",
    "nutrition": "Тамақтануды ұйымдастыру",
    "safety": "Қауіпсіздік",
    "lesson": "Сабаққа қатысу және талдау",
    "attestation": "Қорытынды аттестаттау / ҰБТ",
    "upbringing": "Тәрбие жұмысы",
    "docs": "Педагог құжаттамасы",
    "generic": "Әмбебап кесте",
}

# Excel профилі → каталогтағы тапсырма
PROFILE_TO_TASK = {"grades": "2.1", "diagnostic": "2.2", "sor": "2.3", "correction": "2.4", "attendance": "1.1",
                   "nutrition": "1.2", "safety": "6.1", "lesson": "3.1", "attestation": "4.1",
                   "upbringing": "5.1", "docs": "7.1"}


# ---------------------------------------------------------------- көмекші функциялар
def is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def to_str(v) -> str:
    """Ұяшық мәнін мәтінге айналдыру (JS String(v ?? "") аналогы)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        if not math.isfinite(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return repr(v)
    if isinstance(v, _dt.datetime):
        return v.strftime("%d.%m.%Y") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v, _dt.date):
        return v.strftime("%d.%m.%Y")
    return str(v)


def s(v) -> str:
    return to_str(v).strip()


def xnorm(v) -> str:
    t = to_str(v).lower().replace("ё", "е")
    t = re.sub(r"[«»\"'`]", "", t)
    t = re.sub(r"[%№]", " ", t)
    t = re.sub(r"[^a-zа-яәғқңөұүһі0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _clean_num(n: float):
    return int(n) if float(n).is_integer() else n


def xnum(v, percent: bool = False):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if not math.isfinite(v):
            return None
        n = _clean_num(v)
        return n * 100 if percent and abs(n) <= 1 else n
    t = to_str(v).strip()
    if not t:
        return None
    t = re.sub(r"\s", "", t).replace(",", ".", 1)
    had = "%" in t
    t = t.replace("%", "", 1)
    t = re.sub(r"[^0-9+\-.]", "", t)
    if not t:
        return None
    try:
        n = float(t)
    except ValueError:
        return None
    if not math.isfinite(n):
        return None
    n = _clean_num(n)
    return n * 100 if percent and not had and abs(n) <= 1 else n


def xfmt(v, d: int = 1) -> str:
    if v is None or not is_num(v):
        return "—"
    out = f"{v:,.{d}f}"
    return out.replace(",", " ").replace(".", ",")


def xpct(v, d: int = 1) -> str:
    return xfmt(v, d) + "%" if v is not None and is_num(v) else "—"


def sdiv(a, b, k=1.0, default=0.0):
    return a / b * k if b else default


def cell(r, i):
    return r[i] if 0 <= i < len(r) else None


def issue(level, text):
    return {"level": level, "text": text}


def num_show(v):
    return _clean_num(v) if is_num(v) else v


# ---------------------------------------------------------------- бағандарды тану
def find_col(headers, aliases) -> int:
    hn = [xnorm(h) for h in headers]
    aa = [xnorm(a) for a in aliases]
    for a in aa:
        if a in hn:
            return hn.index(a)
    for a in [a for a in aa if len(a) >= 4]:
        for i, h in enumerate(hn):
            if (a in h) or (h in a and len(h) >= 4):
                return i
    return -1


COL_ALIASES = {
    "klass": ["класс", "сынып"],
    "subject": ["предмет", "дисциплина", "пән"],
    "n": ["n", "всего учащихся", "количество учащихся", "число учащихся", "всего обучающихся", "всего",
          "оқушылар саны", "барлық оқушы", "жалпы оқушы саны", "барлығы"],
    "g5": ["5", "оценка 5", "баға 5"], "g4": ["4", "оценка 4", "баға 4"],
    "g3": ["3", "оценка 3", "баға 3"], "g2": ["2", "оценка 2", "баға 2"],
    "one3": ["с одной 3", "одна 3", "резерв качества", "бір 3", "бір 3 і бар", "сапа резерві"],
    "errors": ["типичные ошибки", "проблемные темы", "проблемная тема", "типтік қателер", "проблемалық тақырыптар"],
    "sor": ["качество сор", "сор качество", "бжб сапасы", "сапасы бжб"],
    "soch": ["качество соч", "соч качество", "тжб сапасы", "сапасы тжб"],
    "high": ["высокий уровень заданий", "высокий уровень", "жоғары деңгей тапсырмалары", "жоғары деңгей"],
    "teacher": ["учитель", "педагог", "мұғалім"],
    "covered": ["охвачены обучением", "охвачено обучением", "оқумен қамтылған"],
    "absences": ["пропуски всего", "всего пропусков", "барлық жіберілген сабақ", "жалпы жіберілген сабақ",
                 "жіберілген сабақ барлығы"],
    "unjust": ["без уважительной причины", "неуважительные пропуски", "дәлелсіз себеппен", "дәлелсіз"],
    "longAbsent": ["длительно отсутствуют", "ұзақ уақыт қатыспайтындар"],
    "measures": ["принятые меры", "меры", "қабылданған шаралар"],
    "initial": ["первичный результат", "первоначальный результат", "бастапқы нәтиже"],
    "repeat": ["повторный результат", "итоговый результат", "қайталама нәтиже"],
    "correction": ["коррекция проведена", "коррекционная работа", "түзету жұмысы"],
    "id": ["id", "код", "коды"],
    "hot": ["получают горячее питание", "горячее питание", "ыстық тамақ алады", "ыстық тамақ"],
    "susn": ["учащиеся сусн", "сусн", "әот оқушылары", "әот"],
    "susnFree": ["сусн получают бесплатное питание", "бесплатное питание сусн", "әот тегін тамақ алады",
                 "әот тегін тамақ"],
    "remarks": ["замечания комиссии", "замечания", "комиссия ескертулері", "ескертулер"],
    "object": ["объект контроля", "оборудование", "бақылау нысаны", "жабдық"],
    "working": ["исправно", "исправные", "жарамды"],
    "broken": ["неисправно", "неисправные", "жарамсыз"],
    "deadline": ["срок устранения", "срок", "жою мерзімі", "мерзім"],
    "goalClear": ["цель урока понятна учащимся", "цель урока понятна", "цель понятна", "сабақ мақсаты түсінікті"],
    "differentiation": ["дифференциация", "саралау"],
    "formative": ["формативное оценивание", "фо", "формативті бағалау"],
    "activity": ["активность учащихся", "активность", "оқушылар белсенділігі", "белсенділік"],
    "reflection": ["рефлексия"],
    "test1": ["тест 1", "тест №1", "тест1"], "test2": ["тест 2", "тест №2", "тест2"],
    "test3": ["тест 3", "тест №3", "тест3"],
    "profileSubjects": ["профильные предметы", "профиль", "бейіндік пәндер"],
    "problemSection": ["проблемный раздел", "проблемный блок", "проблемалық бөлім"],
    "direction": ["направление", "бағыт"],
    "planned": ["запланировано", "план", "жоспарланған", "жоспар"],
    "conducted": ["проведено", "факт", "өткізілген"],
    "coverage": ["охват учащихся", "охват", "қамту"],
    "feedback": ["результат обратная связь", "результат/обратная связь", "обратная связь", "нәтиже кері байланыс",
                 "кері байланыс"],
    "ksp": ["ксп", "қмж"],
    "ktp": ["соответствие ктп", "ктп", "ктж"],
    "journal": ["журнал заполнен", "журнал", "журнал толтырылған"],
    "homework": ["домашнее задание", "үй тапсырмасы"],
    "gradesSet": ["оценки выставлены", "оценки", "бағалар қойылған"],
}


def col_map(headers) -> dict:
    return {k: find_col(headers, v) for k, v in COL_ALIASES.items()}


def detect_profile(headers) -> str:
    c = col_map(headers)
    ok = lambda *ks: all(c[k] >= 0 for k in ks)  # noqa: E731
    if ok("sor", "soch"):
        return "sor"
    if ok("initial", "repeat"):
        return "correction"
    if ok("covered", "absences", "unjust"):
        return "attendance"
    if ok("hot", "n"):
        return "nutrition"
    if ok("object", "working", "broken"):
        return "safety"
    if c["differentiation"] >= 0 and (c["formative"] >= 0 or c["activity"] >= 0):
        return "lesson"
    if ok("test1", "test2"):
        return "attestation"
    if ok("direction", "planned", "conducted"):
        return "upbringing"
    if c["ksp"] >= 0 and (c["journal"] >= 0 or c["ktp"] >= 0):
        return "docs"
    if ok("subject", "n", "g5", "g4", "g3", "g2"):
        return "diagnostic"
    if ok("n", "g5", "g4", "g3", "g2"):
        return "grades"
    return "generic"


def base_audit(table, profile):
    a = []
    if not table["rows"]:
        a.append(issue("error", "Тақырып жолынан кейін деректер жолдары жоқ."))
    if len(table["headers"]) < 2:
        a.append(issue("error", "Кестеде екіден аз баған бар."))
    en = sum(1 for h in table["headers"] if h.startswith("Бағана"))
    if en:
        a.append(issue("warn", f"Атауы жоқ бағандар: {en}."))
    a.append(issue("ok", f"Танылған деректер жолдары: {len(table['rows'])}; бағандар: {len(table['headers'])}."))
    a.append(issue("ok", f"Талдау профилі: {PROFILE_NAMES[profile]}."))
    return a


def empty_analysis(audit, msg):
    audit.append(issue("error", msg))
    return {"audit": audit, "kpis": [("Мәртебе", "Деректер жоқ")], "table": [["Нәтиже"], [msg]], "facts": [msg],
            "positive": [], "problems": [msg], "decisions": [], "profile": "generic"}


def _ln(table, i):
    return i + table["header_index"] + 2


def _label(r):
    return " · ".join(x for x in (r.get("subject", ""), r.get("klass", "")) if x)


def _signed(v, d=1):
    return ("+" if v >= 0 else "") + xfmt(v, d)


# ---------------------------------------------------------------- 1–2. Үлгерім / диагностика
def analyze_grades(table, diag=False):
    prof = "diagnostic" if diag else "grades"
    c = col_map(table["headers"])
    audit = base_audit(table, prof)
    rows, mism = [], 0
    for i, r in enumerate(table["rows"]):
        N, g5, g4, g3, g2 = (xnum(cell(r, c[k])) for k in ("n", "g5", "g4", "g3", "g2"))
        if any(v is None for v in (N, g5, g4, g3, g2)):
            audit.append(issue("warn", f"{_ln(table, i)}-жол: N/5/4/3/2 сандық өрістерінің бәрі танылмады."))
            continue
        if N <= 0:
            audit.append(issue("error", f"{_ln(table, i)}-жол: N нөлден үлкен болуы керек."))
            continue
        tot = num_show(g5 + g4 + g3 + g2)
        if tot != N:
            mism += 1
            audit.append(issue("error", f"{_ln(table, i)}-жол: 5+4+3+2 = {tot}, бірақ N = {N}."))
        rows.append({"subject": s(cell(r, c["subject"])), "klass": s(cell(r, c["klass"])) or f"{i + 1}-жол",
                     "N": N, "g5": g5, "g4": g4, "g3": g3, "g2": g2,
                     "one3": xnum(cell(r, c["one3"])) or 0, "errors": s(cell(r, c["errors"])),
                     "quality": (g5 + g4) / N * 100, "success": (g5 + g4 + g3) / N * 100,
                     "sou": (g5 * 100 + g4 * 64 + g3 * 36 + g2 * 16) / N})
    if not mism:
        audit.append(issue("ok", "Танылған барлық жолдарда бағалар қосындысы N-ге сәйкес келеді."))
    if not rows:
        return empty_analysis(audit, "Бағалары бар жолдарды тану мүмкін болмады.")

    sm = lambda k: num_show(sum(r[k] for r in rows))  # noqa: E731
    total, g5, g4, g3, g2, reserve = sm("N"), sm("g5"), sm("g4"), sm("g3"), sm("g2"), sm("one3")
    oQ, oS = (g5 + g4) / total * 100, (g5 + g4 + g3) / total * 100
    oSou = (g5 * 100 + g4 * 64 + g3 * 36 + g2 * 16) / total
    srt = sorted(rows, key=lambda r: r["quality"])
    low, high = srt[0], srt[-1]
    if diag:
        f0 = (f"{len(rows)} талдау жолы бойынша жиынтық көлем N = {total}. Бір оқушылар әр түрлі пәндер бойынша "
              f"қайталанса, бұл бірегей оқушылар саны емес. Жолдар бойынша өлшенген сапа — {xpct(oQ)}, "
              f"үлгерім — {xpct(oS)}, СОУ — {xpct(oSou)}.")
    else:
        f0 = (f"{len(rows)} талдау жолы бойынша {total} оқушы есепке алынды. Деректер жиынтығы бойынша білім "
              f"сапасы — {xpct(oQ)}, үлгерім — {xpct(oS)}, СОУ — {xpct(oSou)}.")
    facts = [f0,
             f"Ең жоғары сапа: {_label(high)} — {xpct(high['quality'])}; ең төмен: {_label(low)} — "
             f"{xpct(low['quality'])}. Айырма — {xfmt(high['quality'] - low['quality'])} п.п.",
             f"Кестедегі қанағаттанарлықсыз нәтижелер: {g2}. Сапа резерві (бір «3»-і барлар): {reserve}."]
    if diag:
        errs = [f"{_label(r)}: {r['errors']}" for r in rows if r["errors"]]
        if errs:
            facts.append(f"Тіркелген типтік қателер: {'; '.join(errs)}.")
    positive = [f"{_label(r)}: сапа {xpct(r['quality'])}, үлгерім {xpct(r['success'])}."
                for r in rows if r["quality"] >= oQ and r["g2"] == 0][:6]
    problems = []
    for r in sorted([r for r in rows if r["g2"] > 0 or r["quality"] < oQ], key=lambda r: r["quality"])[:8]:
        problems.append(f"{_label(r)}: сапа {xpct(r['quality'])}" + (f", «2» — {r['g2']}" if r["g2"] else "")
                        + (f"; расталған қиындық тақырыбы — {r['errors']}" if r["errors"] else "") + ".")
    if diag:
        headers = ["Пән", "Сынып", "N", "5", "4", "3", "2", "Сапа", "Үлгерім", "СОУ", "Типтік қателер"]
        out = [[r["subject"], r["klass"], r["N"], r["g5"], r["g4"], r["g3"], r["g2"], xpct(r["quality"]),
                xpct(r["success"]), xpct(r["sou"]), r["errors"]] for r in rows]
        decisions = ["Нәтижесі ең төмен жолдар бойынша жұмыстарға элементтік талдау жасап, түзетуді жоспарлау үшін тек тіркелген типтік қателерді пайдалану.",
                     "Түзету жұмысынан кейін сол оқу мақсаттары бойынша [күні] дейін қайталама әкімшілік тілім өткізу."]
    else:
        headers = ["Сынып", "N", "5", "4", "3", "2", "Сапа", "Үлгерім", "СОУ", "Сапа резерві"]
        out = [[r["klass"], r["N"], r["g5"], r["g4"], r["g3"], r["g2"], xpct(r["quality"]), xpct(r["success"]),
                xpct(r["sou"]), r["one3"]] for r in rows]
        decisions = ["Сапасы жиынтық көрсеткіштен төмен жолдар бойынша нәтижелерді пәндік тұрғыда егжей-тегжейлеу; төмендеу себептерін тек жұмыстар мен оқу тақырыптарын талдағаннан кейін анықтау.",
                     f"Сапа резервін ({reserve} оқушы) бөлек қарап, себептер туралы автоматты қорытынды жасамай, нақты пәндер бойынша мекенжайлы шараларды анықтау.",
                     "Түзетуден кейін [күні] дейін қайталама диагностика/тілім өткізіп, динамиканы салыстыру."]
    return {"audit": audit,
            "kpis": [("N қосындысы" if diag else "Оқушылар", total), ("Сапа", xpct(oQ)), ("Үлгерім", xpct(oS)),
                     ("Қанағаттанарлықсыз" if diag else "Сапа резерві", g2 if diag else reserve)],
            "table": [headers, *out], "facts": facts,
            "positive": positive or ["Берілген ереже бойынша оң үрдіс автоматты түрде бөлінбеді; мазмұнды түсіндіру қажет."],
            "problems": problems or ["Ұсынылған таңдамада бағалар бойынша критикалық ауытқулар анықталмады."],
            "decisions": decisions, "profile": prof}


# ---------------------------------------------------------------- 3. БЖБ / ТЖБ
def analyze_sor(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "sor")
    rows = []
    for i, r in enumerate(table["rows"]):
        so, sh = xnum(cell(r, c["sor"]), True), xnum(cell(r, c["soch"]), True)
        if so is None or sh is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: БЖБ/ТЖБ сапасы танылмады."))
            continue
        rows.append({"subject": s(cell(r, c["subject"])), "klass": s(cell(r, c["klass"])), "sor": so, "soch": sh,
                     "delta": so - sh, "abs": abs(so - sh), "high": xnum(cell(r, c["high"]), True)})
    if not rows:
        return empty_analysis(audit, "БЖБ/ТЖБ көрсеткіштерін тану мүмкін болмады.")
    flagged = sorted([r for r in rows if r["abs"] >= 20], key=lambda r: -r["abs"])
    mx = max(rows, key=lambda r: r["abs"])
    mean = sum(r["abs"] for r in rows) / len(rows)
    facts = [f"{len(rows)} пән/сынып үйлесімі талданды. БЖБ–ТЖБ арасындағы орташа абсолютті айырма — {xfmt(mean)} п.п.",
             f"Ең үлкен айырма: {_label(mx)} — {xfmt(mx['delta'])} п.п. (модулі бойынша {xfmt(mx['abs'])} п.п.).",
             f"Абсолютті дельтасы кемінде 20 п.п. жағдайлар: {len(flagged)}. Дельтаның болуы өз алдына бағалаудың объективсіздігін дәлелдемейді."]
    positive = [f"{_label(r)}: дельта {xfmt(r['delta'])} п.п., қосымша тексерудің 20 п.п. шегінен төмен."
                for r in rows if r["abs"] < 20][:8]
    problems = [f"{_label(r)}: БЖБ {xpct(r['sor'])}, ТЖБ {xpct(r['soch'])}, дельта {xfmt(r['delta'])} п.п."
                + (f" Жоғары деңгей тапсырмалары: {xpct(r['high'])}." if r["high"] is not None else "")
                for r in flagged]
    ordered = sorted(rows, key=lambda r: -r["abs"])
    return {"audit": audit,
            "kpis": [("Талдау жолдары", len(rows)), ("Орташа |дельта|", xfmt(mean) + " п.п."),
                     ("≥ 20 п.п.", len(flagged)), ("Максимум", xfmt(mx["abs"]) + " п.п.")],
            "table": [["Пән", "Сынып", "БЖБ сапасы", "ТЖБ сапасы", "Дельта, п.п.", "|Дельта|", "Жоғары деңгей"],
                      *[[r["subject"], r["klass"], xpct(r["sor"]), xpct(r["soch"]), xfmt(r["delta"]), xfmt(r["abs"]),
                         xpct(r["high"]) if r["high"] is not None else "—"] for r in ordered]],
            "facts": facts,
            "positive": positive or ["Танылған барлық жол белгіленген шек бойынша қосымша тексеруді қажет етеді."],
            "problems": problems or ["Дельтасы ≥20 п.п. жағдайлар анықталмады."],
            "decisions": (["Дельтасы ең үлкен жағдайлар бойынша пәндік модерация ұйымдастыру: БЖБ/ТЖБ тапсырмаларын, критерийлерді, дескрипторларды, оқушы жұмыстарын және модерация материалдарын зерттеу.",
                           "Қосымша тексеру нәтижелерін жауапты директор орынбасарына [күні] дейін ұсыну."]
                          if flagged else
                          ["20 п.п. шегі бойынша елеулі айырмашылықтардың жоқтығын тіркеп, бағалау объективтілігінің жоспарлы мониторингін жалғастыру."]),
            "profile": "sor"}


# ---------------------------------------------------------------- 4. Қатысу
def analyze_attendance(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "attendance")
    rows = []
    for i, r in enumerate(table["rows"]):
        N, cov, ab, unj = (xnum(cell(r, c[k])) for k in ("n", "covered", "absences", "unjust"))
        if any(v is None for v in (N, cov, ab, unj)):
            audit.append(issue("warn", f"{_ln(table, i)}-жол: қатысу көрсеткіштерінің бәрі танылмады."))
            continue
        if cov > N:
            audit.append(issue("error", f"{_ln(table, i)}-жол: оқумен қамтылғандар ({cov}) жалпы саннан ({N}) көп."))
        if unj > ab:
            audit.append(issue("error", f"{_ln(table, i)}-жол: дәлелсіз себеппен жіберілген сабақ ({unj}) жалпы санынан ({ab}) көп."))
        rows.append({"klass": s(cell(r, c["klass"])), "N": N, "cov": cov, "abs": ab, "unj": unj,
                     "long": xnum(cell(r, c["longAbsent"])) or 0, "measures": s(cell(r, c["measures"])),
                     "coverage": sdiv(cov, N, 100), "unjustShare": sdiv(unj, ab, 100)})
    if not rows:
        return empty_analysis(audit, "Қатысу көрсеткіштерін тану мүмкін болмады.")
    sm = lambda k: num_show(sum(r[k] for r in rows))  # noqa: E731
    total, cov, ab, unj = sm("N"), sm("cov"), sm("abs"), sm("unj")
    mx = max(rows, key=lambda r: r["unjustShare"])
    if all(abs(r["coverage"] - 100) < .01 for r in rows):
        audit.append(issue("ok", "Танылған барлық жолда оқумен қамту 100% құрайды."))
    else:
        audit.append(issue("warn", "Оқумен қамтуы 100%-дан төмен сыныптар бар; бастапқы деректер мен жағдайды тексеру қажет."))
    facts = [f"Деректер жиынтығы бойынша оқумен қамту — {xpct(sdiv(cov, total, 100))} ({total} оқушының {cov}-і).",
             f"Жіберілген сабақтар барлығы: {ab}; дәлелсіз себеппен: {unj} (барлық жіберілген сабақтың {xpct(sdiv(unj, ab, 100)) if ab else '—'}).",
             f"Негізсіз жіберілген сабақтардың ең жоғары үлесі: {mx['klass']} — {xpct(mx['unjustShare'])} ({mx['abs']}-тің {mx['unj']}-і)."]
    positive = [f"{r['klass']}: қамту 100%, негізсіз жіберілген сабақ тіркелмеген."
                for r in rows if r["unj"] == 0 and r["coverage"] == 100]
    problems = [f"{r['klass']}: қамту {xpct(r['coverage'])}, негізсіз жіберілген сабақ {r['unj']} (барлық жіберілгеннің {xpct(r['unjustShare'])})."
                for r in sorted([r for r in rows if r["unj"] > 0 or r["coverage"] < 100], key=lambda r: -r["unjustShare"])]
    return {"audit": audit,
            "kpis": [("Оқушылар", total), ("Қамту", xpct(sdiv(cov, total, 100))), ("Жіберілген сабақ", ab),
                     ("Дәлелсіз себеппен", unj)],
            "table": [["Сынып", "Барлығы", "Қамтылған", "Қамту", "Жіберілген", "Дәлелсіз", "Дәлелсіз үлесі",
                       "Ұзақ қатыспайтындар", "Қабылданған шаралар"],
                      *[[r["klass"], r["N"], r["cov"], xpct(r["coverage"]), r["abs"], r["unj"], xpct(r["unjustShare"]),
                         r["long"], r["measures"]] for r in rows]],
            "facts": facts,
            "positive": positive or ["Таңдамада негізсіз жіберілген сабағы жоқ сыныптар бөлінбеді."],
            "problems": problems or ["Оқумен қамту және негізсіз сабақ жіберу бойынша ауытқулар анықталмады."],
            "decisions": ["Негізсіз жіберілген сабақ үлесі ең жоғары сыныптар бойынша расталған себептерге мекенжайлы талдау жүргізу; себептерді тек сабақ жіберу санына қарап анықтамау.",
                          "Сынып жетекшісіне/әлеуметтік педагогқа [аты-жөні] қабылданған шаралар туралы ақпаратты [күні] дейін ұсыну.",
                          "[мерзім] өткен соң негізсіз жіберілген сабақтардың жалпы санын және үлесін қайта салыстыру."],
            "profile": "attendance"}


# ---------------------------------------------------------------- 5. Олқылықтарды толтыру
def analyze_correction(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "correction")
    rows = []
    for i, r in enumerate(table["rows"]):
        a, b = xnum(cell(r, c["initial"]), True), xnum(cell(r, c["repeat"]), True)
        if a is None or b is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: бастапқы/қайталама нәтижелер танылмады."))
            continue
        d = b - a
        cat = "Айқын оң" if d > 15 else "Шамалы" if d > 0 else "Динамикасыз" if d == 0 else "Теріс"
        rid = cell(r, c["id"])
        rows.append({"id": to_str(rid) if rid is not None else f"{i + 1}-жол", "klass": to_str(cell(r, c["klass"])),
                     "subject": to_str(cell(r, c["subject"])), "topic": to_str(cell(r, c["errors"])),
                     "initial": a, "repeat": b, "delta": d, "cat": cat,
                     "correction": to_str(cell(r, c["correction"]))})
    if not rows:
        return empty_analysis(audit, "Нәтижелер динамикасын тану мүмкін болмады.")
    ex = sum(1 for r in rows if r["delta"] > 15)
    neg = sum(1 for r in rows if r["delta"] < 0)
    none = sum(1 for r in rows if r["delta"] == 0)
    avg = sum(r["delta"] for r in rows) / len(rows)
    facts = [f"{len(rows)} жеке жазба талданды. Орташа динамика — {xfmt(avg)} п.п.",
             f"Айқын оң динамика (>15 п.п.) — {ex}; динамикасыз — {none}; теріс динамика — {neg}.",
             "Түзету фактісі мен нәтиже өзгерісінің сәйкес келуі себеп-салдарлық байланыстың жеткілікті дәлелі болып табылмайды."]
    positive = [f"{r['id']}" + (f" ({r['klass']})" if r["klass"] else "")
                + f": {xpct(r['initial'])} → {xpct(r['repeat'])}, +{xfmt(r['delta'])} п.п." for r in rows if r["delta"] > 15]
    problems = [f"{r['id']}" + (f" · {r['topic']}" if r["topic"] else "")
                + f": {xpct(r['initial'])} → {xpct(r['repeat'])}, {_signed(r['delta'])} п.п. — {r['cat']}."
                for r in sorted([r for r in rows if r["delta"] <= 15], key=lambda r: r["delta"])]
    return {"audit": audit,
            "kpis": [("Жазбалар", len(rows)), (">15 п.п.", ex), ("Жоқ/теріс", none + neg),
                     ("Орташа динамика", xfmt(avg) + " п.п.")],
            "table": [["ID", "Сынып", "Пән", "Проблемалық тақырып", "Бастапқы", "Түзету", "Қайталама", "Динамика, п.п.", "Санат"],
                      *[[r["id"], r["klass"], r["subject"], r["topic"], xpct(r["initial"]), r["correction"],
                         xpct(r["repeat"]), xfmt(r["delta"]), r["cat"]] for r in rows]],
            "facts": facts,
            "positive": positive or ["Айқын оң динамика (>15 п.п.) анықталмады."],
            "problems": problems or ["Динамикасы жеткіліксіз немесе теріс жазбалар анықталмады."],
            "decisions": ["Динамикасы теріс, нөлдік немесе шамалы жазбалар бойынша [күні] дейін сол мазмұн элементтері бойынша қосымша диагностика жүргізу.",
                          "Жұмыстарды талдау негізінде меңгерілмеген мазмұн элементтерін нақтылап, жеке маршруттарды түзету.",
                          "Келесі диагностикалық тілімді өткізіп, сол оқу мақсаттары бойынша динамиканы тіркеу."],
            "profile": "correction"}


# ---------------------------------------------------------------- 6. Тамақтану
def analyze_nutrition(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "nutrition")
    rows = []
    for i, r in enumerate(table["rows"]):
        N, hot = xnum(cell(r, c["n"])), xnum(cell(r, c["hot"]))
        if N is None or hot is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: тамақпен қамту көрсеткіштері танылмады."))
            continue
        susn, free = xnum(cell(r, c["susn"])), xnum(cell(r, c["susnFree"]))
        if hot > N:
            audit.append(issue("error", f"{_ln(table, i)}-жол: тамақ алатындар ({hot}) жалпы саннан ({N}) көп."))
        if susn is not None and free is not None and free > susn:
            audit.append(issue("error", f"{_ln(table, i)}-жол: ӘОТ тегін тамағы ({free}) ӘОТ санынан ({susn}) көп."))
        rows.append({"klass": to_str(cell(r, c["klass"])), "N": N, "hot": hot, "coverage": sdiv(hot, N, 100),
                     "susn": susn or 0, "free": free or 0,
                     "susnCoverage": (free or 0) / susn * 100 if susn else None, "remarks": s(cell(r, c["remarks"]))})
    if not rows:
        return empty_analysis(audit, "Тамақтану көрсеткіштерін тану мүмкін болмады.")
    sm = lambda k: num_show(sum(r[k] for r in rows))  # noqa: E731
    total, hot, susn, free = sm("N"), sm("hot"), sm("susn"), sm("free")
    low = min(rows, key=lambda r: r["coverage"])
    no_rem = lambda r: (not r["remarks"]) or xnorm(r["remarks"]) in ("нет", "жоқ")  # noqa: E731
    remarks = [r for r in rows if not no_rem(r)]
    facts = [f"Ыстық тамақпен жалпы қамту — {xpct(sdiv(hot, total, 100))} ({total} оқушының {hot}-і).",
             f"Ұсынылған деректердегі ең төмен қамту: {low['klass']} — {xpct(low['coverage'])}.",
             f"ӘОТ оқушыларын тегін тамақпен қамту — {xpct(free / susn * 100)} ({susn}-ның {free}-і)." if susn
             else "ӘОТ бойынша деректер ұсынылмаған немесе танылмады.",
             f"Комиссия ескертулері тіркелген жолдар: {len(remarks)}."]
    positive = [f"{r['klass']}: ұсынылған көрсеткіштер бойынша толық қамту, ескертулер тіркелмеген."
                for r in rows if r["coverage"] == 100 and (not r["susn"] or r["susnCoverage"] == 100) and no_rem(r)]
    problems = [f"{r['klass']}: қамту {xpct(r['coverage'])}" + (f"; ескерту: {r['remarks']}" if r["remarks"] else "") + "."
                for r in rows if r["coverage"] < 100 or not no_rem(r)]
    return {"audit": audit,
            "kpis": [("Оқушылар", total), ("Ыстық тамақ", xpct(sdiv(hot, total, 100))),
                     ("ӘОТ тегін", xpct(free / susn * 100) if susn else "—"), ("Ескертулер", len(remarks))],
            "table": [["Сынып", "Барлығы", "Ыстық тамақ", "Қамту", "ӘОТ", "ӘОТ тегін", "ӘОТ қамту", "Ескертулер"],
                      *[[r["klass"], r["N"], r["hot"], xpct(r["coverage"]), r["susn"], r["free"],
                         xpct(r["susnCoverage"]) if r["susnCoverage"] is not None else "—", r["remarks"]] for r in rows]],
            "facts": facts,
            "positive": positive or ["Толық қамтылған және ескертусіз жолдар автоматты түрде бөлінбеді."],
            "problems": problems or ["Ұсынылған деректерде қамту және ескертулер бойынша ауытқулар анықталмады."],
            "decisions": ["Қамтуы толық емес сыныптар бойынша расталған ұйымдастырушылық себептерді нақтылап, [күні] дейін шараларды тіркеу.",
                          "Комиссия ескертулері бойынша нақты әрекетті, жауаптыны, жою мерзімін және қайталама тексеру нысанын анықтау.",
                          "Қамту мен ескертулердің орындалуына қайталама мониторинг жүргізу."],
            "profile": "nutrition"}


# ---------------------------------------------------------------- 7. Қауіпсіздік
def analyze_safety(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "safety")
    rows = []
    for i, r in enumerate(table["rows"]):
        N, w, b = xnum(cell(r, c["n"])), xnum(cell(r, c["working"])), xnum(cell(r, c["broken"]))
        if N is None or w is None or b is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: жабдық саны танылмады."))
            continue
        if w + b != N:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: жарамды + жарамсыз = {num_show(w + b)}, барлығы = {N}."))
        rows.append({"object": to_str(cell(r, c["object"])), "N": N, "w": w, "b": b, "rate": sdiv(w, N, 100),
                     "remarks": s(cell(r, c["remarks"])), "deadline": s(cell(r, c["deadline"]))})
    if not rows:
        return empty_analysis(audit, "Қауіпсіздік көрсеткіштерін тану мүмкін болмады.")
    total, w, b = (num_show(sum(r[k] for r in rows)) for k in ("N", "w", "b"))
    worst = min(rows, key=lambda r: r["rate"])
    facts = [f"Ұсынылған жабдық бойынша жалпы жарамдылық — {xpct(sdiv(w, total, 100))} ({total}-тің {w}-ы).",
             f"Жарамсыз бірліктер: {b}. Жарамдылықтың ең төмен пайызы: {worst['object']} — {xpct(worst['rate'])}.",
             "Модуль анықталған фактілерді қолданылатын НҚА-ны тексермей, нормативтік талаптарды бұзу деп саралаймайды."]
    positive = [f"{r['object']}: {r['N']}-тің {r['w']}-ы жарамды (100%)." for r in rows if r["b"] == 0]
    problems = [f"{r['object']}: {r['N']}-тің {r['b']}-ы жарамсыз; жарамдылық {xpct(r['rate'])}"
                + (f"; ескерту: {r['remarks']}" if r["remarks"] else "")
                + (f"; мерзімі: {r['deadline']}" if r["deadline"] else "") + "." for r in rows if r["b"] > 0]
    return {"audit": audit,
            "kpis": [("Барлық бірлік", total), ("Жарамдылық", xpct(sdiv(w, total, 100))), ("Жарамсыз", b),
                     ("Ақауы бар нысандар", sum(1 for r in rows if r["b"] > 0))],
            "table": [["Бақылау нысаны", "Барлығы", "Жарамды", "Жарамсыз", "Жарамдылық", "Ескерту", "Жою мерзімі"],
                      *[[r["object"], r["N"], r["w"], r["b"], xpct(r["rate"]), r["remarks"], r["deadline"]] for r in rows]],
            "facts": facts,
            "positive": positive or ["100% жарамды нысандар автоматты түрде бөлінбеді."],
            "problems": problems or ["Ұсынылған деректерде жарамсыздық тіркелмеген."],
            "decisions": ["Әр тіркелген жарамсыздық бойынша жауапты мен жою мерзімін анықтау/тексеру; кестеде жоқ бұзушылықтарды қоспау.",
                          "Белгіленген мерзім өткен соң қайталама техникалық тексеру жүргізіп, жою фактісін тіркеу."],
            "profile": "safety"}


# ---------------------------------------------------------------- 8. Әмбебап
def analyze_generic(table):
    audit = base_audit(table, "generic")
    stats = []
    for j, h in enumerate(table["headers"]):
        nums = [v for v in (xnum(cell(r, j)) for r in table["rows"]) if v is not None]
        if len(nums) >= max(2, math.ceil(len(table["rows"]) * .4)):
            stats.append({"h": h, "n": len(nums), "min": min(nums), "max": max(nums), "mean": sum(nums) / len(nums)})
    return {"audit": audit,
            "kpis": [("Жолдар", len(table["rows"])), ("Бағандар", len(table["headers"])),
                     ("Сандық өрістер", len(stats)), ("Профиль", "Әмбебап")],
            "table": [["Сандық баған", "Мәндер саны", "Минимум", "Максимум", "Орташа"],
                      *[[x["h"], x["n"], xfmt(x["min"], 2), xfmt(x["max"], 2), xfmt(x["mean"], 2)] for x in stats]],
            "facts": [f"Әмбебап кесте автоматты танылды: {len(table['rows'])} жол, {len(table['headers'])} баған.",
                      f"Негізгі статистикаға жарамды сандық бағандар: {len(stats)}.",
                      "Пәндік басқарушылық қорытындылар үшін құрылым қолдау көрсетілетін үлгілердің біріне сәйкес келсе, профильді қолмен таңдаңыз."],
            "positive": ["Автоматты статистикалық құрылым мазмұнды себеп-салдарлық қорытындысыз құрылды."],
            "problems": ["Пәндік түсіндіру шектеулі: құрылым қолдау көрсетілетін профильге сәйкес келмейді немесе әмбебап режим таңдалған."],
            "decisions": ["Басқарушылық тапсырма түрін нақтылап, қажет болса, талдау профилін қолмен таңдау.",
                          "Ресми құжатты дайындағанға дейін әр көрсеткіштің мағынасын, өлшем бірліктерін және жолдардың салыстырмалылығын тексеру."],
            "profile": "generic"}


# ---------------------------------------------------------------- 9. Сабаққа қатысу
YES = {"да", "есть", "выполнено", "соответствует", "полностью", "иә", "бар", "орындалған", "сәйкес", "толық"}
PART = {"частично", "не полностью", "ішінара", "толық емес"}
NO = {"нет", "не выполнено", "не соответствует", "жоқ", "орындалмаған", "сәйкес емес"}


def xyes(v):
    t = xnorm(v)
    if t in YES:
        return 1
    if t in PART:
        return 0.5
    if t in NO:
        return 0
    return None


def cell_score(v):
    return {1: "Иә", 0.5: "Ішінара", 0: "Жоқ"}.get(v, "—") if v is not None else "—"


def _avg(rows, k):
    a = [r[k] for r in rows if r[k] is not None]
    return sum(a) / len(a) if a else None


def analyze_lesson(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "lesson")
    rows = [{"teacher": to_str(cell(r, c["teacher"])), "klass": to_str(cell(r, c["klass"])),
             "goal": xyes(cell(r, c["goalClear"])), "diff": xyes(cell(r, c["differentiation"])),
             "fo": xyes(cell(r, c["formative"])), "activity": xnum(cell(r, c["activity"]), True),
             "refl": xyes(cell(r, c["reflection"])), "remarks": s(cell(r, c["remarks"]))} for r in table["rows"]]
    pct = lambda k: None if _avg(rows, k) is None else _avg(rows, k) * 100  # noqa: E731
    act = _avg(rows, "activity")
    pd, pf, pr = pct("diff"), pct("fo"), pct("refl")
    facts = [f"Бақылаулар саны: {len(rows)}.",
             f"«Иә/ішінара/жоқ» шкаласы бойынша саралау: интегралдық көрсеткіш {xpct(pd)}. Формативті бағалау: {xpct(pf)}. Рефлексия: {xpct(pr)}.",
             f"Көрсетілген жолдар бойынша оқушылардың орташа белсенділігі: {xpct(act)}."]
    probs = []
    if pd is not None and pd < 70:
        probs.append("Саралау бақылаулар сериясында біркелкі емес іске асады және әрі қарай әдістемелік зерттеуді қажет етеді.")
    if pf is not None and pf < 70:
        probs.append("Формативті бағалау сапасы бақылаулар сериясы бойынша қосымша назарды қажет етеді.")
    if pr is not None and pr < 70:
        probs.append("Рефлексия барлық бақылауда көрсетілмеген.")
    return {"audit": audit,
            "kpis": [("Бақылаулар", len(rows)), ("Саралау", xpct(pd)), ("ФБ", xpct(pf)), ("Белсенділік", xpct(act))],
            "table": [["Педагог", "Сынып", "Мақсат", "Саралау", "ФБ", "Белсенділік", "Рефлексия", "Ескертулер"],
                      *[[r["teacher"], r["klass"], cell_score(r["goal"]), cell_score(r["diff"]), cell_score(r["fo"]),
                         xpct(r["activity"]), cell_score(r["refl"]), r["remarks"]] for r in rows]],
            "facts": facts,
            "positive": ["Талдау педагог тұлғасы бойынша емес, әдістемелік аспектілер бойынша орындалды."],
            "problems": probs or ["Қайталанатын әдістемелік назар аймақтары автоматты түрде анықталмады."],
            "decisions": ["Қайталанатын өсу аймақтары бойынша әдістемелік талдау ұйымдастырып, қолдаудың нақты форматын анықтау.",
                          "Бірыңғай чек-лист бойынша қайталама тақырыптық бақылаулар сериясын өткізу.",
                          "Жекелеген бақылау негізінде жеке педагогтың кәсіби құзыреттілігі туралы қорытынды жасамау."],
            "profile": "lesson"}


# ---------------------------------------------------------------- 10. Аттестаттау / ҰБТ
def analyze_attestation(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "attestation")
    rows = []
    for i, r in enumerate(table["rows"]):
        t1, t2, t3 = (xnum(cell(r, c[k])) for k in ("test1", "test2", "test3"))
        if t1 is None or t2 is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: динамиканы есептеуге нәтиже жеткіліксіз."))
            continue
        vals = [v for v in (t1, t2, t3) if v is not None]
        if len(vals) >= 3:
            trend = ("Тұрақты өсу" if vals[1] > vals[0] and vals[2] > vals[1] else
                     "Тұрақты төмендеу" if vals[1] < vals[0] and vals[2] < vals[1] else "Тұрақсыз динамика")
        else:
            trend = "Өсу" if t2 > t1 else "Төмендеу" if t2 < t1 else "Динамикасыз"
        last = vals[-1]
        rows.append({"id": to_str(cell(r, c["id"])), "klass": to_str(cell(r, c["klass"])), "t1": t1, "t2": t2, "t3": t3,
                     "last": last, "delta": last - t1, "trend": trend, "risk": last < 50,
                     "section": to_str(cell(r, c["problemSection"]))})
    if not rows:
        return empty_analysis(audit, "Тестілеулер сериясын тану мүмкін болмады.")
    risk = [r for r in rows if r["risk"]]
    decl = [r for r in rows if "төмендеу" in r["trend"].lower()]
    who = lambda r: r["id"] or r["klass"]  # noqa: E731
    return {"audit": audit,
            "kpis": [("Қатысушылар", len(rows)), ("Тәуекел тобы", len(risk)), ("Төмендеу", len(decl)),
                     ("Орташа динамика", xfmt(sum(r["delta"] for r in rows) / len(rows), 1))],
            "table": [["ID", "Сынып", "Тест 1", "Тест 2", "Тест 3", "Динамика", "Тренд", "Проблемалық бөлім"],
                      *[[r["id"], r["klass"], r["t1"], r["t2"], r["t3"] if r["t3"] is not None else "—",
                         _signed(r["delta"]), r["trend"], r["section"]] for r in rows]],
            "facts": [f"Талданған оқушылар/идентификаторлар: {len(rows)}.",
                      f"Прототип ережесі бойынша тәуекел тобы (соңғы тестілеуде <50 балл): {len(risk)}.",
                      f"Тұрақты/тіркелген теріс динамика: {len(decl)}."],
            "positive": [f"{who(r)}: тұрақты өсу, динамика {_signed(r['delta'])}." for r in rows if r["trend"] == "Тұрақты өсу"][:12],
            "problems": ([f"{who(r)}: соңғы нәтиже {r['last']} (<50), мекенжайлы бақылау қажет." for r in risk]
                         + [f"{who(r)}: {r['trend'].lower()}, динамика {xfmt(r['delta'], 1)}." for r in decl])[:20],
            "decisions": ["Тәуекел тобы үшін көрсетілген проблемалық бөлімдер бойынша мекенжайлы диагностика жүргізу.",
                          "Төмендеу жағдайлары бойынша сол бөлімдердегі жұмыстарды салыстырып, дайындық жоспарын түзету.",
                          "Келесі салыстырмалы тестілеуде динамиканы қайта бағалау."],
            "profile": "attestation"}


# ---------------------------------------------------------------- 11. Тәрбие жұмысы
def analyze_upbringing(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "upbringing")
    rows = []
    for i, r in enumerate(table["rows"]):
        p, f, cov = xnum(cell(r, c["planned"])), xnum(cell(r, c["conducted"])), xnum(cell(r, c["coverage"]), True)
        if p is None or f is None:
            audit.append(issue("warn", f"{_ln(table, i)}-жол: жоспар/факт танылмады."))
            continue
        rows.append({"direction": to_str(cell(r, c["direction"])), "planned": p, "done": f, "exec": sdiv(f, p, 100),
                     "coverage": cov, "feedback": to_str(cell(r, c["feedback"])), "remarks": to_str(cell(r, c["remarks"]))})
    if not rows:
        return empty_analysis(audit, "Тәрбие жұмысының көрсеткіштерін тану мүмкін болмады.")
    planned, done = num_show(sum(r["planned"] for r in rows)), num_show(sum(r["done"] for r in rows))
    covs = [r["coverage"] for r in rows if r["coverage"] is not None]
    probs = [f"{r['direction']}: {r['planned']}-тің {r['done']}-ы орындалды" + (f"; мәселе/ескерту: {r['remarks']}" if r["remarks"] else "") + "."
             for r in rows if r["done"] < r["planned"] or r["remarks"]]
    return {"audit": audit,
            "kpis": [("Жоспарланған", planned), ("Өткізілген", done),
                     ("Орындалуы", xpct(done / planned * 100) if planned else "—"), ("Бағыттар", len(rows))],
            "table": [["Бағыт", "Жоспар", "Факт", "Орындалуы", "Қамту", "Нәтиже/кері байланыс", "Мәселелер"],
                      *[[r["direction"], r["planned"], r["done"], xpct(r["exec"]), xpct(r["coverage"]), r["feedback"],
                         r["remarks"]] for r in rows]],
            "facts": [f"Жоспардың орындалуы: {xpct(done / planned * 100) if planned else '—'} ({planned}-тің {done}-ы).",
                      f"Көрсетілген бағыттар бойынша орташа қамту: {xpct(sum(covs) / len(covs)) if covs else '—'}.",
                      "Тиімділік өткізілген іс-шаралар санымен теңестірілмейді: тек ұсынылған нәтижелер/кері байланыс ескеріледі."],
            "positive": [f"{r['direction']}: ұсынылған деректер бойынша жоспар орындалған." for r in rows
                         if r["done"] >= r["planned"] and not r["remarks"]],
            "problems": probs or ["Жоспардың орындалмауы немесе тіркелген мәселелер автоматты түрде анықталмады."],
            "decisions": ["Жоспар орындалмаған бағыттар бойынша расталған себептерді анықтап, түзету әрекеттерін белгілеу.",
                          "Өлшенетін нәтижесі жоқ бағыттар бойынша кері байланыс/диагностика нысанын анықтау.",
                          "Қайталама бақылауды тек іс-шара өткізілу фактісі бойынша емес, нәтиже мен қамту бойынша жүргізу."],
            "profile": "upbringing"}


# ---------------------------------------------------------------- 12. Педагог құжаттамасы
def analyze_docs(table):
    c = col_map(table["headers"])
    audit = base_audit(table, "docs")
    fields = ["ksp", "ktp", "journal", "homework", "grades"]
    src = {"ksp": "ksp", "ktp": "ktp", "journal": "journal", "homework": "homework", "grades": "gradesSet"}
    rows = []
    for r in table["rows"]:
        row = {"teacher": to_str(cell(r, c["teacher"])), "remarks": s(cell(r, c["remarks"]))}
        for f in fields:
            row[f] = xyes(cell(r, c[src[f]]))
        rows.append(row)
    vals = [r[k] for r in rows for k in fields if r[k] is not None]
    score = sum(vals) / len(vals) * 100 if vals else None
    rr = [r for r in rows if r["remarks"]]
    return {"audit": audit,
            "kpis": [("Жазбалар", len(rows)), ("Сәйкестік", xpct(score)), ("Ескертулермен", len(rr)),
                     ("Ескертусіз", len(rows) - len(rr))],
            "table": [["Педагог", "ҚМЖ", "КТЖ", "Журнал", "Үй тапсырмасы", "Бағалар", "Ескертулер"],
                      *[[r["teacher"], *[cell_score(r[k]) for k in fields], r["remarks"]] for r in rows]],
            "facts": [f"Тексерілген жазбалар: {len(rows)}.",
                      f"Танылған өрістер бойынша сәйкестіктің интегралдық үлесі: {xpct(score)}.",
                      f"Ескертулері тіркелген жолдар: {len(rr)}. Жүйелілік бір фактімен анықталмайды."],
            "positive": [f"{r['teacher'] or 'Жол'}: ескертулер тіркелмеген." for r in rows
                         if not r["remarks"] and all(r[k] in (1, None) for k in fields)][:12],
            "problems": [f"{r['teacher'] or 'Жол'}: {r['remarks']}" for r in rr][:20],
            "decisions": ["Қайталанатын ескертулерді түрлері бойынша топтап, мекенжайлы жою мерзімдерін анықтау.",
                          "Ескертуді жеткілікті тексеру санынсыз жүйелі деп сараламау.",
                          "Белгіленген мерзімнен кейін сол өлшемдер бойынша қайталама бақылау жүргізу."],
            "profile": "docs"}


ANALYZERS = {
    "grades": lambda t: analyze_grades(t, False), "diagnostic": lambda t: analyze_grades(t, True),
    "sor": analyze_sor, "attendance": analyze_attendance, "correction": analyze_correction,
    "nutrition": analyze_nutrition, "safety": analyze_safety, "lesson": analyze_lesson,
    "attestation": analyze_attestation, "upbringing": analyze_upbringing, "docs": analyze_docs,
}


def run_profile(table, profile):
    return ANALYZERS.get(profile, analyze_generic)(table)
