"""
Фаза 4: прогноз срока выполнения работ фазы 5 методом Монте-Карло (две модели).

Запуск:  uv run python phase-4/simulation.py
Читает листы «Задачи», «История», «Параметры», пишет результаты на лист «Результаты» (значения, не формулы).

Модель 1 (по числу задач). Каждую неделю число выполненных задач берётся случайно из истории прошлых недель
  (дискретное распределение). Недели считаются, пока сумма выполненных задач не достигнет числа задач фазы 5.
Модель 2 (по трудозатратам). Трудозатраты каждой задачи берутся из треугольного распределения (О, В, П).
  Часы, которые команда уделяет предмету в неделю, берутся случайно из истории (дискретное распределение).
  Недели считаются, пока накопленные часы не покроют суммарные трудозатраты (задачи выполняются подряд, часы общие).
Срок с достоверностью 80 % – наименьшее целое число недель W, для которого доля запусков с результатом ≤ W не меньше 0,8.
"""
import sys
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def read(path):
	wb = load_workbook(path, data_only=True)
	t, h, p = wb["Задачи"], wb["История"], wb["Параметры"]
	tasks = []
	r = 5
	while isinstance(t.cell(row=r, column=1).value, (int, float)):
		tasks.append(tuple(float(t.cell(row=r, column=c).value) for c in (4, 5, 6)))
		r += 1
	hist_tasks, hist_hours = [], []
	r = 5
	while h.cell(row=r, column=1).value not in (None, "Среднее"):
		hist_tasks.append(float(h.cell(row=r, column=2).value)); hist_hours.append(float(h.cell(row=r, column=3).value))
		r += 1
	return wb, np.array(tasks), np.array(hist_tasks), np.array(hist_hours), int(p["B3"].value), int(p["B4"].value), float(p["B5"].value)

def model1(n_tasks, hist, runs, rng, max_weeks=200):
	weeks = np.zeros(runs, int)
	done = np.zeros(runs)
	active = np.ones(runs, bool)
	w = 0
	while active.any() and w < max_weeks:
		w += 1
		done[active] += rng.choice(hist, size=active.sum())
		finished = active & (done >= n_tasks)
		weeks[finished] = w
		active &= ~finished
	weeks[active] = max_weeks
	return weeks

def model2(tasks, hist_hours, runs, rng, max_weeks=200):
	effort = rng.triangular(tasks[:, 0][None, :], tasks[:, 1][None, :], tasks[:, 2][None, :], size=(runs, len(tasks))).sum(axis=1)
	weeks = np.zeros(runs, int)
	done = np.zeros(runs)
	active = np.ones(runs, bool)
	w = 0
	while active.any() and w < max_weeks:
		w += 1
		done[active] += rng.choice(hist_hours, size=active.sum())
		finished = active & (done >= effort)
		weeks[finished] = w
		active &= ~finished
	weeks[active] = max_weeks
	return weeks, effort

def pct(weeks, q):
	s = np.sort(weeks)
	return int(s[int(np.ceil(q * len(s))) - 1])

if __name__ == "__main__":
	path = sys.argv[1] if len(sys.argv) > 1 else "phase-4/phase-4.xlsx"
	wb, tasks, ht, hh, runs, seed, q = read(path)
	rng = np.random.default_rng(seed)
	w1 = model1(len(tasks), ht, runs, rng)
	w2, effort = model2(tasks, hh, runs, rng)
	stats = {}
	for name, w in (("Модель 1: по числу задач", w1), ("Модель 2: по трудозатратам", w2)):
		stats[name] = dict(mean=w.mean(), p50=pct(w, 0.5), p80=pct(w, q), p90=pct(w, 0.9), mn=int(w.min()), mx=int(w.max()))
		print(name, {k: (round(v, 2) if isinstance(v, float) else v) for k, v in stats[name].items()})
	print("Трудозатраты фазы 5, ч: среднее %.1f, P80 %.1f" % (effort.mean(), np.percentile(effort, 80)))

	wb = load_workbook(path) # открываем с формулами (read() даёт только значения)
	ws = wb["Результаты"]
	ws.delete_rows(1, ws.max_row + 1)
	F = "Arial"
	bold, norm, title = Font(name=F, size=11, bold=True), Font(name=F, size=11), Font(name=F, size=14, bold=True)
	fh = PatternFill("solid", fgColor="D9E1F2"); fr = PatternFill("solid", fgColor="E2EFDA")
	thin = Side(style="thin", color="999999"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
	ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)

	def put(ref, v, font=norm, fmt=None, fill=None, align=None, border=True):
		c = ws[ref]; c.value = v; c.font = font
		if fmt: c.number_format = fmt
		if fill: c.fill = fill
		if align: c.alignment = align
		if border: c.border = box

	put("A1", "Результаты симуляции Монте-Карло", title, border=False)
	put("A2", f"Значения записаны скриптом simulation.py ({runs} запусков, seed {seed}); при изменении данных перезапустите скрипт.", border=False)
	for col, t in zip("ABCDEFG", ["Модель", "Среднее, недель", "Медиана (50 %)", f"Достоверность {q:.0%}", "90 %", "Минимум", "Максимум"]):
		put(f"{col}4", t, bold, fill=fh, align=ctr)
	for i, (name, s) in enumerate(stats.items()):
		r = 5 + i
		put(f"A{r}", name)
		put(f"B{r}", round(float(s["mean"]), 2), fmt="0.00", align=ctr)
		for col, k in zip("CDEFG", ("p50", "p80", "p90", "mn", "mx")):
			put(f"{col}{r}", s[k], bold if k == "p80" else norm, "0", fr if k == "p80" else None, ctr)
	put("A8", f"Трудозатраты фазы 5 в модели 2, ч (среднее / {q:.0%})", border=False)
	put("D8", f"{effort.mean():.1f} / {np.percentile(effort, 100*q):.1f}", bold, align=ctr, border=False)
	put("A10", "Распределение числа недель", bold, border=False)
	for col, t in zip("ABCDE", ["Недель", "Модель 1, доля", "Модель 1, накопленная", "Модель 2, доля", "Модель 2, накопленная"]):
		put(f"{col}11", t, bold, fill=fh, align=ctr)
	top = int(max(w1.max(), w2.max()))
	c1 = c2 = 0.0
	for k, wk in enumerate(range(1, min(top, 20) + 1)):
		p1, p2 = float((w1 == wk).mean()), float((w2 == wk).mean())
		c1 += p1; c2 += p2
		r = 12 + k
		put(f"A{r}", wk, align=ctr)
		for col, v in zip("BCDE", (p1, c1, p2, c2)): put(f"{col}{r}", round(v, 4), fmt="0.0%", align=ctr)
	ws.column_dimensions["A"].width = 32
	for col in "BCDEFG": ws.column_dimensions[col].width = 20
	wb.save(path)
	print("Результаты записаны в", path)
