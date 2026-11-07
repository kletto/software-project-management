"""
Фаза 2: распределение задач фазы 3 между участниками (задача о назначениях, целочисленное программирование).

Запуск: uv run solve.py

Скрипт читает листы «Задачи», «Предпочтения», «Параметры», решает задачу (scipy.optimize.milp, решатель HiGHS),
записывает 0/1 в ячейки C5:E… листа «Распределение» и печатает результат и анализ чувствительности к допуску Δ.
Формулы листа пересчитаются при открытии файла в Excel.

Модель:
	x[i,j] = 1, если задачу i выполняет участник j
	max   Σ p[i,j] · x[i,j]                         – суммарная предпочтительность
	при   Σ_j x[i,j] = 1                            – у каждой задачи ровно один исполнитель
		  |Σ_i h[i]·x[i,j] − H/m| ≤ Δ               – нагрузка каждого близка к средней
		  Σ_i x[i,j] ≥ k                            – у каждого не меньше k задач
		  x[a,j] = x[b,j]  для заданных пар (a, b)  – связанные задачи делает один человек
"""

import sys
import numpy as np
from openpyxl import load_workbook
from scipy.optimize import milp, LinearConstraint, Bounds

FIRST = 5 # первая строка с индивидуальными задачами

def read(path):
	wb = load_workbook(path, data_only=True)
	t, p, par = wb["Задачи"], wb["Предпочтения"], wb["Параметры"]
	last = FIRST
	while isinstance(t.cell(row=last, column=1).value, (int, float)):
		last += 1
	last -= 1
	names = [p.cell(row=4, column=c).value for c in (3, 4, 5)]
	tasks = [t.cell(row=r, column=2).value for r in range(FIRST, last + 1)]
	h = np.array([t.cell(row=r, column=4).value for r in range(FIRST, last + 1)], float)
	pref = np.array([[p.cell(row=r, column=c).value for c in (3, 4, 5)] for r in range(FIRST, last + 1)], float)
	delta, kmin = par["B3"].value, par["B4"].value
	pairs = [(par.cell(row=r, column=2).value, par.cell(row=r, column=3).value) for r in (7, 8)
			 if par.cell(row=r, column=2).value]
	return names, tasks, h, pref, float(delta), int(kmin), pairs

def solve(h, pref, delta, kmin, pairs):
	n, m = pref.shape
	idx = lambda i, j: i * m + j
	c = -pref.flatten() # максимизируем -> минимизируем минус
	A, lo, hi = [], [], []

	def add(row, l, u):
		A.append(row); lo.append(l); hi.append(u)

	for i in range(n): # одна задача – один исполнитель
		row = np.zeros(n * m); row[[idx(i, j) for j in range(m)]] = 1
		add(row, 1, 1)
	avg = h.sum() / m
	for j in range(m):
		row = np.zeros(n * m); row[[idx(i, j) for i in range(n)]] = h
		add(row, avg - delta, avg + delta) # баланс нагрузки
		row = np.zeros(n * m); row[[idx(i, j) for i in range(n)]] = 1
		add(row, kmin, n) # минимум задач на человека
	for a, b in pairs: # связанные задачи – один исполнитель
		for j in range(m):
			row = np.zeros(n * m); row[idx(a - 1, j)] = 1; row[idx(b - 1, j)] = -1
			add(row, 0, 0)
	res = milp(c, constraints=LinearConstraint(np.array(A), lo, hi),
			   integrality=np.ones(n * m), bounds=Bounds(0, 1))
	if not res.success:
		return None, None
	return np.round(res.x).astype(int).reshape(n, m), -res.fun

if __name__ == "__main__":
	path = sys.argv[1] if len(sys.argv) > 1 else "phase-2.xlsx"
	names, tasks, h, pref, delta, kmin, pairs = read(path)
	x, total = solve(h, pref, delta, kmin, pairs)
	if x is None:
		raise SystemExit("Решение не найдено: увеличьте допуск Δ или ослабьте другие ограничения.")
	wb_f = load_workbook(path) # открываем с формулами, чтобы не потерять их
	ws = wb_f["Распределение"]
	for i in range(x.shape[0]):
		for j in range(x.shape[1]):
			ws.cell(row=FIRST + i, column=3 + j, value=int(x[i, j]))
	wb_f.save(path)

	print(f"Суммарная предпочтительность: {total:.0f}\n")
	for j, nm in enumerate(names):
		mine = [i for i in range(len(tasks)) if x[i, j]]
		print(f"{nm}: {h[mine].sum():.1f} ч, задач: {len(mine)}, предпочтительность: {pref[mine, j].sum():.0f}")
		for i in mine:
			print(f"   {i+1:>2}. {tasks[i]} ({h[i]:g} ч, оценка {pref[i, j]:g})")
	print("\nЧувствительность к допуску Δ (k и пары задач без изменений):")
	print("Δ, ч | цель | нагрузка (ч)")
	for d in (1, 2, 3, 5, None):
		xs, ts = solve(h, pref, d if d is not None else 1e9, kmin, pairs)
		label = "без ограничения" if d is None else f"{d}"
		if xs is None:
			print(f"{label:>15} | нет решения")
		else:
			loads = ", ".join(f"{(h[xs[:, j] == 1]).sum():.1f}" for j in range(xs.shape[1]))
			print(f"{label:>15} | {ts:.0f} | {loads}")
