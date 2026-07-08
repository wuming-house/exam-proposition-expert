import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

os.makedirs("figs", exist_ok=True)

# 图1：角的示意图（∠AOB=90°，求∠BOC）
fig, ax = plt.subplots(figsize=(3.2, 3.0))
ax.set_aspect("equal")
ax.axis("off")
ax.plot([0, 2.2], [0, 0], "k-", lw=1.5)          # OA
ax.plot([0, 0], [0, 2.2], "k-", lw=1.5)          # OB
ang = np.deg2rad(130)
C = (2.0 * np.cos(ang), 2.0 * np.sin(ang))
ax.plot([0, C[0]], [0, C[1]], "k-", lw=1.5)      # OC
for lab, pt in [("O", (0, 0)), ("A", (2.2, 0)), ("B", (0, 2.2)), ("C", C)]:
    ax.annotate(lab, pt, xytext=(5, 5), textcoords="offset points", fontsize=13)
ax.text(0.5, 0.18, r"$90^\circ$", fontsize=11)
plt.savefig("figs/q2.png", dpi=150, bbox_inches="tight")
plt.close()

# 图2：直角三角形 ABC（∠B=90°）
fig, ax = plt.subplots(figsize=(3.4, 2.6))
ax.set_aspect("equal")
ax.axis("off")
ax.plot([0, 4], [0, 0], "k-", lw=1.5)            # AB
ax.plot([4, 4], [0, 3], "k-", lw=1.5)            # BC
ax.plot([0, 4], [0, 3], "k-", lw=1.5)            # AC
ax.annotate("A", (0, 0), xytext=(-14, 2), fontsize=13)
ax.annotate("B", (4, 0), xytext=(8, -4), fontsize=13)
ax.annotate("C", (4, 3), xytext=(8, 2), fontsize=13)
plt.savefig("figs/q5.png", dpi=150, bbox_inches="tight")
plt.close()

print("figs done:", os.listdir("figs"))
