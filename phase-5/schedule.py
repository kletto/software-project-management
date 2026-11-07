"""
Фаза 5: календарное планирование проекта с ограниченными ресурсами (эвристики + случайные порядки).

Запуск:  uv run python phase-5/schedule.py
Читает листы «Задачи» и «Параметры», строит расписание последовательной схемой (serial SGS) для каждого из 12 правил приоритета
и для заданного числа случайных порядков, выбирает лучшее и записывает листы «Эвристики», «Расписание», «Этапы», «Загрузка»
(даты и проверки – формулами; начала задач – значения решателя). Рисует диаграмму Ганта.
"""
import sys
import datetime as dt
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

RULES = { # имя: (направление, описание)
	"SLK": ("asc", "по возрастанию общего резерва LST − EST"),
	"FREE": ("asc", "по возрастанию свободного резерва min(EST последователей) − EFT"),
	"LST": ("asc", "по возрастанию позднего начала"),
	"LFT": ("asc", "по возрастанию позднего завершения"),
	"LSTLFT": ("asc", "по возрастанию LST + LFT"),
	"GRPW": ("desc", "по убыванию длительности задачи и её прямых последователей"),
	"LPT": ("desc", "по убыванию длительности"),
	"MIS": ("desc", "по убыванию числа прямых последователей"),
	"GRD": ("desc", "по убыванию длительности × суммарные затраты ресурсов"),
	"GRWC": ("desc", "по убыванию суммарных затрат ресурсов"),
	"GCRWC": ("desc", "по убыванию суммарных затрат ресурсов задачи и её последователей"),
	"ROT": ("desc", "по убыванию Σ(затраты/запас) / длительность"),
}

def read(path):
	wb = load_workbook(path, data_only=True)
	t, p = wb["Задачи"], wb["Параметры"]
	names, stages, d, req, preds = [], [], [], [], []
	r = 5
	while isinstance(t.cell(row=r, column=1).value, (int, float)):
		names.append(t.cell(row=r, column=2).value); stages.append(t.cell(row=r, column=3).value)
		d.append(int(t.cell(row=r, column=4).value))
		req.append([int(t.cell(row=r, column=c).value or 0) for c in (5, 6, 7)])
		preds.append([int(t.cell(row=r, column=c).value) - 1 for c in range(9, 15) if t.cell(row=r, column=c).value])
		r += 1
	start = p["B3"].value
	start = start.date() if isinstance(start, dt.datetime) else start
	par = {"start": start, "hpd": int(p["B4"].value), "nrand": int(p["B5"].value), "seed": int(p["B6"].value),
			   "cap": [int(p.cell(row=9 + k, column=2).value) for k in range(3)], "resnames": [p.cell(row=9 + k, column=1).value for k in range(3)]}
	return wb, names, stages, np.array(d), np.array(req), preds, par

def cpm(d, preds):
	n = len(d)
	succ = [[] for _ in range(n)]
	for j, ps in enumerate(preds):
		for i in ps: succ[i].append(j)
	est, eft = np.zeros(n), np.zeros(n)
	for j in range(n): # задачи заданы в топологическом порядке (предшественники имеют меньший номер)
		est[j] = max([eft[i] for i in preds[j]], default=0); eft[j] = est[j] + d[j]
	T = eft.max()
	lst, lft = np.zeros(n), np.zeros(n)
	for j in reversed(range(n)):
		lft[j] = min([lst[k] for k in succ[j]], default=T); lst[j] = lft[j] - d[j]
	return succ, est, eft, lst, lft, T

def priorities(d, req, preds, cap):
	succ, est, eft, lst, lft, T = cpm(d, preds)
	n = len(d)
	tot = req.sum(axis=1)
	free = np.array([min([est[k] for k in succ[j]], default=T) - eft[j] for j in range(n)])
	v = {
		"SLK": lst - est, "FREE": free, "LST": lst, "LFT": lft, "LSTLFT": lst + lft,
		"GRPW": np.array([d[j] + sum(d[k] for k in succ[j]) for j in range(n)], float),
		"LPT": d.astype(float), "MIS": np.array([len(succ[j]) for j in range(n)], float),
		"GRD": d * tot.astype(float), "GRWC": tot.astype(float),
		"GCRWC": np.array([tot[j] + sum(tot[k] for k in succ[j]) for j in range(n)], float),
		"ROT": np.array([(req[j] / np.array(cap)).sum() / d[j] for j in range(n)]),
	}
	return {k: (x if RULES[k][0] == "asc" else -x) for k, x in v.items()}, T

def sgs(key, d, req, preds, cap):
	"""Последовательная схема: из доступных задач берём с наименьшим ключом, ставим в самое раннее допустимое время."""
	n = len(d); H = int(d.sum()) + 1
	use = np.zeros((len(cap), H), int)
	start = np.full(n, -1); fin = np.zeros(n, int)
	done = np.zeros(n, bool)
	for _ in range(n):
		elig = [j for j in range(n) if not done[j] and all(done[i] for i in preds[j])]
		j = min(elig, key=lambda x: (key[x], x))
		t = max([fin[i] for i in preds[j]], default=0)
		while any((use[k, t:t + d[j]] + req[j, k] > cap[k]).any() for k in range(len(cap))): t += 1
		for k in range(len(cap)): use[k, t:t + d[j]] += req[j, k]
		start[j] = t; fin[j] = t + d[j]; done[j] = True
	return start, int(fin.max())

def cal_date(start, k): # k-й рабочий день от даты начала (k = 1, 2, ...)
	return np.busday_offset(np.datetime64(start), k - 1, roll="forward")

if __name__ == "__main__":
	path = sys.argv[1] if len(sys.argv) > 1 else "phase-5/phase-5.xlsx"
	png = sys.argv[2] if len(sys.argv) > 2 else "phase-5/gantt.png"
	wb, names, stages, d, req, preds, par = read(path)
	cap, hpd, n = np.array(par["cap"]), par["hpd"], len(d)
	keys, T = priorities(d, req, preds, par["cap"])
	results = {}
	for rule, key in keys.items():
		results[rule] = sgs(key, d, req, preds, cap)
	rng = np.random.default_rng(par["seed"])
	best_rand = None; rand_ms = []
	for _ in range(par["nrand"]):
		key = rng.random(n)
		s, ms = sgs(key, d, req, preds, cap)
		rand_ms.append(ms)
		if best_rand is None or ms < best_rand[1]: best_rand = (s, ms)
	rand_ms = np.array(rand_ms)
	cands = list(results.items()) + [("Случайный", best_rand)]
	best_name, (best_start, best_ms) = min(cands, key=lambda x: x[1][1])
	lb_res = max(int(np.ceil((d * req[:, k]).sum() / cap[k])) for k in range(3))
	print(f"Критический путь (без ресурсов): {T:.0f} ч; нижняя оценка по ресурсам: {lb_res} ч")
	for r_, (s, ms) in results.items(): print(f"{r_:>7}: {ms} ч")
	print(f"Случайные ({par['nrand']}): лучший {rand_ms.min()} ч, средний {rand_ms.mean():.1f} ч, достигли лучшего: {(rand_ms == rand_ms.min()).mean():.1%}")
	print(f"Лучшее решение: {best_name}, срок {best_ms} ч = {best_ms/hpd:.2f} раб. дн.")

	# ---------------- запись листов
	wb = load_workbook(path) # открываем с формулами (read() даёт только значения)
	for nm in ("Эвристики", "Расписание", "Этапы", "Загрузка"):
		if nm in wb.sheetnames: del wb[nm]
	F = "Arial"
	bold, norm, title = Font(name=F, size=11, bold=True), Font(name=F, size=11), Font(name=F, size=14, bold=True)
	blue = Font(name=F, size=11, color="0000FF"); green = Font(name=F, size=11, color="008000")
	fh = PatternFill("solid", fgColor="D9E1F2"); fr = PatternFill("solid", fgColor="E2EFDA"); fi = PatternFill("solid", fgColor="FFF9DB")
	thin = Side(style="thin", color="999999"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
	ctr = Alignment(horizontal="center", vertical="center", wrap_text=True); wrap = Alignment(vertical="center", wrap_text=True)

	def put(ws, ref, v, font=norm, fmt=None, fill=None, align=None, border=True):
		c = ws[ref]; c.value = v; c.font = font
		if fmt: c.number_format = fmt
		if fill: c.fill = fill
		if align: c.alignment = align
		if border: c.border = box
		return c
	hdr = lambda ws, ref, t: put(ws, ref, t, bold, fill=fh, align=ctr)

	# --- Эвристики
	we = wb.create_sheet("Эвристики")
	put(we, "A1", "Сравнение эвристик", title, border=False)
	put(we, "A2", f"Значения записаны скриптом schedule.py (seed {par['seed']}); при изменении данных перезапустите скрипт.", border=False)
	for col, t in zip("ABCD", ["Правило", "Описание", "Срок проекта, ч", "Рабочих дней"]): hdr(we, f"{col}4", t)
	r = 5
	for rule, (s, ms) in results.items():
		put(we, f"A{r}", rule, bold if rule == best_name else norm); put(we, f"B{r}", RULES[rule][1], align=wrap)
		put(we, f"C{r}", ms, bold if ms == best_ms else norm, "0", fr if ms == best_ms else None, ctr)
		put(we, f"D{r}", f"=C{r}/Параметры!$B$4", norm, "0.00", align=ctr); r += 1
	put(we, f"A{r}", "Случайные"); put(we, f"B{r}", f"лучший из {par['nrand']} случайных порядков", align=wrap)
	put(we, f"C{r}", int(rand_ms.min()), bold if rand_ms.min() == best_ms else norm, "0", fr if rand_ms.min() == best_ms else None, ctr)
	put(we, f"D{r}", f"=C{r}/Параметры!$B$4", norm, "0.00", align=ctr); r += 2
	put(we, f"A{r}", "Лучшее решение", bold, border=False); put(we, f"C{r}", f"{best_name}", bold, fill=fr, align=ctr); r += 1
	put(we, f"A{r}", "Критический путь без учёта ресурсов, ч", border=False); put(we, f"C{r}", int(T), norm, "0", align=ctr); r += 1
	put(we, f"A{r}", "Нижняя оценка по загрузке специальностей, ч", border=False); put(we, f"C{r}", lb_res, norm, "0", align=ctr); r += 1
	put(we, f"A{r}", "Случайные порядки: средний срок, ч", border=False); put(we, f"C{r}", round(float(rand_ms.mean()), 1), norm, "0.0", align=ctr); r += 1
	put(we, f"A{r}", "Случайные порядки: доля достигших лучшего срока", border=False); put(we, f"C{r}", round(float((rand_ms == rand_ms.min()).mean()), 4), norm, "0.0%", align=ctr)
	we.column_dimensions["A"].width = 44; we.column_dimensions["B"].width = 62; we.column_dimensions["C"].width = 18; we.column_dimensions["D"].width = 14

	# --- Расписание
	wr = wb.create_sheet("Расписание")
	put(wr, "A1", f"Расписание проекта (лучшее решение: {best_name})", title, border=False)
	put(wr, "A2", "Начало задач – результат решателя (часы от начала проекта); окончание, даты и проверка предшествования считаются формулами.", border=False)
	for col, t in zip("ABCDEFGH", ["№", "Задача", "Этап", "Начало, ч", "Окончание, ч", "Дата начала", "Дата окончания", "Предшествование"]): hdr(wr, f"{col}4", t)
	for k in range(6): hdr(wr, f"{'JKLMNO'[k]}4", f"Окончание предш. {k+1}, ч")
	for j in range(n):
		r = 5 + j
		put(wr, f"A{r}", j + 1, align=ctr); put(wr, f"B{r}", f"=Задачи!B{r}", green, align=wrap); put(wr, f"C{r}", f"=Задачи!C{r}", green, align=wrap)
		put(wr, f"D{r}", int(best_start[j]), blue, "0", fi, ctr)
		put(wr, f"E{r}", f"=D{r}+Задачи!D{r}", norm, "0", align=ctr)
		put(wr, f"F{r}", f"=WORKDAY(Параметры!$B$3-1,INT(D{r}/Параметры!$B$4)+1)", norm, "DD.MM.YYYY", align=ctr)
		put(wr, f"G{r}", f"=WORKDAY(Параметры!$B$3-1,ROUNDUP(E{r}/Параметры!$B$4,0))", norm, "DD.MM.YYYY", align=ctr)
		for k, col in enumerate("JKLMNO"):
			pc = "IJKLMN"[k]
			put(wr, f"{col}{r}", f'=IF(Задачи!{pc}{r}="","",INDEX($E$5:$E${4+n},Задачи!{pc}{r}))', Font(name=F, size=11, color="808080"), "0", align=ctr)
		put(wr, f"H{r}", f'=IF(D{r}>=MAX(J{r}:O{r}),"ок","нарушено")', bold, align=ctr)
	for col, w in zip("ABCDEFGHIJKLMNO", [6, 52, 28, 12, 14, 14, 16, 18, 4, 14, 14, 14, 14, 14, 14]): wr.column_dimensions[col].width = w
	wr.column_dimensions["I"].width = 3
	wr.freeze_panes = "C5"

	# --- Этапы
	wsg = wb.create_sheet("Этапы")
	put(wsg, "A1", "Плановые сроки этапов и проекта", title, border=False)
	for col, t in zip("ABCDEF", ["Этап", "Начало, ч", "Окончание, ч", "Дата начала", "Дата окончания", "Рабочих дней"]): hdr(wsg, f"{col}4", t)
	uniq = list(dict.fromkeys(stages))
	for i, st in enumerate(uniq):
		r = 5 + i
		put(wsg, f"A{r}", st, align=wrap)
		put(wsg, f"B{r}", f"=_xlfn.MINIFS(Расписание!$D$5:$D${4+n},Расписание!$C$5:$C${4+n},A{r})", norm, "0", align=ctr)
		put(wsg, f"C{r}", f"=_xlfn.MAXIFS(Расписание!$E$5:$E${4+n},Расписание!$C$5:$C${4+n},A{r})", norm, "0", align=ctr)
	rt = 5 + len(uniq)
	put(wsg, f"A{rt}", "Проект в целом", bold)
	put(wsg, f"B{rt}", f"=MIN(Расписание!$D$5:$D${4+n})", bold, "0", align=ctr); put(wsg, f"C{rt}", f"=MAX(Расписание!$E$5:$E${4+n})", bold, "0", fill=fr, align=ctr)
	for r in range(5, rt + 1):
		b = bold if r == rt else norm
		put(wsg, f"D{r}", f"=WORKDAY(Параметры!$B$3-1,INT(B{r}/Параметры!$B$4)+1)", b, "DD.MM.YYYY", align=ctr)
		put(wsg, f"E{r}", f"=WORKDAY(Параметры!$B$3-1,ROUNDUP(C{r}/Параметры!$B$4,0))", b, "DD.MM.YYYY", fr if r == rt else None, ctr)
		put(wsg, f"F{r}", f"=ROUNDUP(C{r}/Параметры!$B$4,0)-INT(B{r}/Параметры!$B$4)", b, "0", align=ctr)
	wsg.column_dimensions["A"].width = 34
	for col in "BCDEF": wsg.column_dimensions[col].width = 16
	put(wsg, f"A{rt+2}", "Этапы пересекаются: задачи разных этапов выполняются параллельно, пока позволяют зависимости и ресурсы.", border=False)

	# --- Загрузка
	wl = wb.create_sheet("Загрузка")
	put(wl, "A1", "Загрузка специальностей по часам проекта", title, border=False)
	put(wl, "A2", "Занято людей каждой специальности в каждый час; не должно превышать число людей на листе «Параметры».", border=False)
	for col, t in zip("ABCDE", ["Час", "Frontend", "Backend", "Тестировщик", "Проверка"]): hdr(wl, f"{col}4", t)
	for t_ in range(best_ms):
		r = 5 + t_
		put(wl, f"A{r}", t_, norm, "0", align=ctr)
		for k, col in enumerate("BCD"):
			rc = "EFG"[k]
			put(wl, f"{col}{r}", f"=SUMPRODUCT((Расписание!$D$5:$D${4+n}<=$A{r})*(Расписание!$E$5:$E${4+n}>$A{r})*Задачи!${rc}$5:${rc}${4+n})", norm, "0", align=ctr)
		put(wl, f"E{r}", f'=IF(AND(B{r}<=Параметры!$B$9,C{r}<=Параметры!$B$10,D{r}<=Параметры!$B$11),"ок","превышение")', bold, align=ctr)
	re_ = 4 + best_ms
	put(wl, "G4", "Итоговая проверка", bold, fill=fh, align=ctr)
	put(wl, "G5", f'=IF(COUNTIF(E5:E{re_},"ок")={best_ms},"ресурсы не превышены","есть превышение")', bold, fill=fr, align=ctr)
	put(wl, "G6", f'=IF(COUNTIF(Расписание!H5:H{4+n},"ок")={n},"предшествование соблюдено","есть нарушения")', bold, fill=fr, align=ctr)
	for col, w in zip("ABCDEG", [8, 12, 12, 14, 14, 30]): wl.column_dimensions[col].width = w
	wl.freeze_panes = "A5"
	wb.save(path)

	# --- диаграмма Ганта
	import matplotlib; matplotlib.use("Agg")
	import matplotlib.pyplot as plt
	colors = dict(zip(dict.fromkeys(stages), ["#8FAADC", "#A9D18E", "#F4B183", "#FFD966", "#C9A0DC"]))
	fig, ax = plt.subplots(figsize=(13, 8))
	for j in range(n):
		ax.barh(n - j, d[j], left=best_start[j], color=colors[stages[j]], edgecolor="#555555", height=0.7)
		who = ", ".join(f"{'FBQ'[k]}" for k in range(3) if req[j, k] for _ in range(req[j, k]))
		ax.text(best_start[j] + d[j] / 2, n - j, f"{j+1}", ha="center", va="center", fontsize=8)
	ax.set_yticks(range(n, 0, -1)); ax.set_yticklabels([f"{j+1}. {names[j]}" for j in range(n)], fontsize=8)
	days = int(np.ceil(best_ms / hpd))
	ax.set_xticks([k * hpd for k in range(days + 1)])
	labels = [str(np.busday_offset(np.datetime64(par["start"]), k, roll="forward").astype(dt.date).strftime("%d.%m")) for k in range(days + 1)]
	ax.set_xticklabels(labels, fontsize=8, rotation=45)
	ax.set_xlim(0, days * hpd); ax.grid(axis="x", linestyle=":", alpha=0.6)
	ax.set_xlabel("Рабочие дни (начало дня)")
	ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c) for c in colors.values()], labels=list(colors.keys()), loc="upper right", fontsize=8)
	ax.set_title(f"Диаграмма Ганта: срок {best_ms} ч ≈ {days} рабочих дней")
	plt.tight_layout(); plt.savefig(png, dpi=150)
	end_date = cal_date(par["start"], int(np.ceil(best_ms / hpd)))
	print("Окончание проекта:", end_date, "| файлы записаны:", path, png)
