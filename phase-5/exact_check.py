"""
Проверка оптимальности расписания фазы 5 точным методом: целочисленное программирование на временной сетке (scipy, HiGHS).
Запуск:  uv run python phase-5/exact_check.py [горизонт_в_часах]
Горизонт по умолчанию – срок лучшего эвристического расписания с листа «Эвристики» (нужно передать числом, если лист изменён).
x[j,t] = 1, если задача j начинается в час t; минимизируем срок проекта при ограничениях предшествования и ресурсов.
"""
import sys
import numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds
from scipy.sparse import lil_matrix
from schedule import read

path = sys.argv[1] if len(sys.argv) > 1 else "phase-5/phase-5.xlsx"
wb, names, stages, d, req, preds, par = read(path)
cap, n = par["cap"], len(d)
H = int(sys.argv[2]) if len(sys.argv) > 2 else int(d.sum())

# переменные x[j, t] – старт задачи j в час t; CM – вспомогательная переменная срока проекта
idx, c = {}, 0
for j in range(n):
	for t in range(H - d[j] + 1):
		idx[(j, t)] = c; c += 1
CM = c
A = lil_matrix((n + sum(len(p) for p in preds) + n + 3 * H, c + 1)); lo, hi = [], []; r = 0

for j in range(n): # каждая задача стартует ровно один раз
	for t in range(H - d[j] + 1): A[r, idx[(j, t)]] = 1
	lo.append(1); hi.append(1); r += 1

for j in range(n): # предшествование
	for i in preds[j]:
		for t in range(H - d[j] + 1): A[r, idx[(j, t)]] += t
		for t in range(H - d[i] + 1): A[r, idx[(i, t)]] -= t + d[i]
		lo.append(0); hi.append(np.inf); r += 1

for j in range(n): # срок проекта не меньше окончания любой задачи
	A[r, CM] = 1
	for t in range(H - d[j] + 1): A[r, idx[(j, t)]] -= t + d[j]
	lo.append(0); hi.append(np.inf); r += 1

for k in range(3): # ресурсы
	for tau in range(H):
		for j in range(n):
			if req[j, k]:
				for t in range(max(0, tau - d[j] + 1), min(tau, H - d[j]) + 1): A[r, idx[(j, t)]] += req[j, k]
		lo.append(-np.inf); hi.append(cap[k]); r += 1

cost = np.zeros(c + 1); cost[CM] = 1
integ = np.ones(c + 1); integ[CM] = 0
res = milp(cost, constraints=LinearConstraint(A[:r].tocsr(), lo, hi), integrality=integ,
		   bounds=Bounds(0, [1] * c + [H]), options={"time_limit": 300})

print(res.message)
print(f"Оптимальный срок проекта: {res.fun:.0f} ч")
