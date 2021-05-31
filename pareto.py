import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import matplotlib.lines as mlines

df = pd.read_excel('pareto_source.xlsx')
# df = pd.read_excel('test.csv', encoding='gb2312')
df = df.sort_values(by='数量', ascending=False)
df.loc[-1] = [' ', 0]
df.index = df.index + 1
df = df.sort_index()
index = df['关键词']
df["percentage"] = df["数量"].cumsum() / df["数量"].sum() * 100
df["currentPercent"] = df["数量"] / df["数量"].sum() * 100
print(df)

p = df["percentage"]
key = p[p >= 80].index[0]

fig, ax = plt.subplots()
ax.bar(index, df["数量"], color="green")
ax2 = ax.twinx()
ax2.plot(index, df["percentage"], color="C1", marker="D", ms=7)
ax2.yaxis.set_major_formatter(PercentFormatter())
ax2.set_ylim(bottom=0)
ax2.set_xlim(0)

ax.tick_params(axis="y", colors="green")
ax2.tick_params(axis="y", colors="C1")

for xtick in ax.get_xticklabels():
    xtick.set_rotation(50)

for i in range(1, len(p)):
    plt.annotate(format(p[i] / 100, '.2%'), xy=(i, p[i]), xytext=(i + 0.1, p[i] - 3))
# plt.annotate(format(p[key] / 100, '.4%'), xy=(key, p[key]), xytext=(key * 1.1, p[key] * 0.9),
#              arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=.2'))

# plt.axvline(key, color='r', linestyle="--", alpha=0.8)
plt.show()
