"""
小红书留学赛道内容策略分析
使用方法：python 小红书数据分析.py
依赖：pip install pandas matplotlib seaborn openpyxl
"""

# ── 导入库 ───────────────────────────────────────────────────────
# pandas：处理表格数据（类似Python里的Excel）
import pandas as pd
# matplotlib：画图的核心库
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
# seaborn：基于matplotlib的美化库（这里主要用它的风格）
import seaborn as sns
import warnings
warnings.filterwarnings("ignore")  # 忽略无关紧要的警告信息

# ── 字体设置（Colab/Linux环境下显示中文）────────────────────────
# Colab默认没有中文字体，需要先安装再注册
import subprocess, matplotlib
subprocess.run(["apt-get", "install", "-y", "fonts-wqy-zenhei"], capture_output=True)
import matplotlib.font_manager as fm
fm.fontManager.addfont("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc")
matplotlib.rcParams["font.family"] = "WenQuanYi Zen Hei"  # 指定使用刚安装的中文字体
matplotlib.rcParams["axes.unicode_minus"] = False  # 防止负号显示成方框

# ── 第一步：读取Excel数据 ────────────────────────────────────────
FILE = "小红书竞品数据收集模板.xlsx"
# read_excel：读取Excel文件，sheet_name指定读哪个工作表
df = pd.read_excel(FILE, sheet_name="笔记数据收集")
# df是DataFrame，可以理解为一张二维表格，行是每篇笔记，列是各个字段

# 去掉列名前后可能存在的空格，防止后面匹配出错
df.columns = df.columns.str.strip()

# ── 第二步：数据清洗 ─────────────────────────────────────────────
# pd.to_numeric：把列强制转换为数字类型
# errors="coerce"：遇到无法转换的值（如空格、文字）自动变成NaN
# fillna(0)：把NaN填充为0，避免后续计算报错
num_cols = ["点赞数", "评论数", "收藏数"]
for c in num_cols:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

# 阅读量只有自己账号后台才能看到，竞品账号这列会是空的
if "阅读量" in df.columns:
    df["阅读量"] = pd.to_numeric(df["阅读量"], errors="coerce")
    # 注意：这里不fillna(0)，因为后面要判断是否为空

# ── 第三步：计算核心指标——综合互动量 ────────────────────────────
# 为什么这样设计：
#   - 竞品账号的阅读量不对外显示，无法用"互动率=互动/阅读量"统一比较
#   - 改用绝对互动量作为统一指标，让自己和竞品可以放在同一把尺子下对比
#   - 收藏×2是因为小红书算法对收藏的权重高于点赞，收藏代表用户"真的觉得有用"
df["综合互动量"] = df["点赞数"] + df["收藏数"] * 2 + df["评论数"]

# ── 第四步：拆分自己账号和竞品账号 ─────────────────────────────
# 用布尔条件筛选行，类似Excel里的筛选功能
own  = df[df["账号类型"] == "自己账号"].copy()   # 自己账号的所有笔记
comp = df[df["账号类型"] != "自己账号"].copy()   # 竞品账号的所有笔记

# 自己账号有阅读量，额外计算互动率作为参考
if "阅读量" in own.columns:
    # apply：对每一行执行一个函数
    # lambda r: ... 是匿名函数，r代表当前行
    own["互动率(%)"] = own.apply(
        lambda r: round(r["综合互动量"] / r["阅读量"] * 100, 2)
        if pd.notna(r["阅读量"]) and r["阅读量"] > 0 else None, axis=1
    )

print("=" * 55)
print("  小红书留学赛道内容策略分析报告")
print("=" * 55)
print(f"  自己账号笔记数：{len(own)}  竞品笔记数：{len(comp)}")
print(f"  竞品账号：{', '.join(comp['账号类型'].unique())}")
print("=" * 55)

# ══════════════════════════════════════════════════════════════
# 图1：四张子图的综合分析
# ══════════════════════════════════════════════════════════════
# plt.subplots(2, 2)：创建一个2行2列共4张图的画布
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle("留学赛道小红书内容策略分析", fontsize=16, fontweight="bold", y=1.01)

colors = ["#2D6A9F", "#E8943A", "#3DAA6E", "#C95D5D", "#8B5CF6"]  # 自定义颜色列表

# ── 图1a：内容类型 vs 平均综合互动量 ────────────────────────────
ax = axes[0, 0]
# groupby("内容类型")：按内容类型分组，类似Excel的数据透视表
# ["综合互动量"].mean()：计算每组的平均值
# sort_values(ascending=False)：从高到低排序
type_avg = df.groupby("内容类型")["综合互动量"].mean().sort_values(ascending=False)
bars = ax.bar(type_avg.index, type_avg.values, color=colors[:len(type_avg)],
              edgecolor="white", linewidth=1.5)
# 在每个柱子顶部显示具体数值
for bar, val in zip(bars, type_avg.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f"{val:.1f}", ha="center", va="bottom", fontsize=10, fontweight="bold")
ax.set_title("① 内容类型 vs 平均综合互动量", fontsize=12, fontweight="bold")
ax.set_xlabel("内容类型")
ax.set_ylabel("平均综合互动量（点赞+收藏×2+评论）")
ax.set_facecolor("#F8F9FA")
ax.spines[["top", "right"]].set_visible(False)  # 去掉上边框和右边框，更简洁

# ── 图1b：是否含Emoji vs 平均综合互动量 ─────────────────────────
ax = axes[0, 1]
# groupby("是否含Emoji")：按是否含Emoji分成"是"和"否"两组
emoji_avg = df.groupby("是否含Emoji")["综合互动量"].mean()
# .get("否", 0)：取"否"组的值，如果不存在则默认为0
bars = ax.bar(["不含Emoji", "含Emoji"],
              [emoji_avg.get("否", 0), emoji_avg.get("是", 0)],
              color=["#94A3B8", "#E8943A"], edgecolor="white", linewidth=1.5, width=0.5)
for bar, val in zip(bars, [emoji_avg.get("否", 0), emoji_avg.get("是", 0)]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f"{val:.1f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax.set_title("② 标题含Emoji vs 平均综合互动量", fontsize=12, fontweight="bold")
ax.set_ylabel("平均综合互动量")
ax.set_facecolor("#F8F9FA")
ax.spines[["top", "right"]].set_visible(False)

# ── 图1c：是否含数字 vs 平均综合互动量 ──────────────────────────
ax = axes[1, 0]
num_avg = df.groupby("是否含数字")["综合互动量"].mean()
bars = ax.bar(["不含数字", "含数字"],
              [num_avg.get("否", 0), num_avg.get("是", 0)],
              color=["#94A3B8", "#2D6A9F"], edgecolor="white", linewidth=1.5, width=0.5)
for bar, val in zip(bars, [num_avg.get("否", 0), num_avg.get("是", 0)]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f"{val:.1f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax.set_title("③ 标题含数字 vs 平均综合互动量", fontsize=12, fontweight="bold")
ax.set_ylabel("平均综合互动量")
ax.set_facecolor("#F8F9FA")
ax.spines[["top", "right"]].set_visible(False)

# ── 图1d：各账号平均综合互动量横向对比 ──────────────────────────
ax = axes[1, 1]
# 按账号名称分组，计算每个账号的平均互动量，从低到高排序（横向柱状图从下到上）
account_avg = df.groupby("账号名称")["综合互动量"].mean().sort_values(ascending=True)
# 自己账号用蓝色，竞品用橙色，通过账号名字里是否有"Tina"来判断
bar_colors = ["#2D6A9F" if "Tina" in name else "#E8943A" for name in account_avg.index]
# 账号名太长就截断，避免图表拥挤
short_names = [n[:8] + "…" if len(n) > 8 else n for n in account_avg.index]
# barh：横向柱状图（h = horizontal）
bars = ax.barh(short_names, account_avg.values, color=bar_colors, edgecolor="white", linewidth=1.5)
for bar, val in zip(bars, account_avg.values):
    ax.text(val + 0.5, bar.get_y() + bar.get_height()/2,
            f"{val:.1f}", va="center", fontsize=10, fontweight="bold")
ax.set_title("④ 各账号平均综合互动量对比", fontsize=12, fontweight="bold")
ax.set_xlabel("平均综合互动量")
ax.set_facecolor("#F8F9FA")
ax.spines[["top", "right"]].set_visible(False)
from matplotlib.patches import Patch
legend = [Patch(color="#2D6A9F", label="自己账号"), Patch(color="#E8943A", label="竞品账号")]
ax.legend(handles=legend, loc="lower right", fontsize=9)

plt.tight_layout()
plt.savefig("分析图表.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n✅ 图表已保存：分析图表.png")

# ══════════════════════════════════════════════════════════════
# 图2：爆文分析（综合互动量排名前20%的笔记）
# ══════════════════════════════════════════════════════════════
# quantile(0.8)：计算第80百分位数，即"排名前20%的门槛值是多少"
threshold = df["综合互动量"].quantile(0.8)
top    = df[df["综合互动量"] >= threshold].copy()  # 爆文：互动量在门槛以上
bottom = df[df["综合互动量"] <  threshold].copy()  # 普通笔记：门槛以下

fig2, axes2 = plt.subplots(1, 2, figsize=(13, 5))
fig2.suptitle("爆文特征分析（综合互动量 Top 20%）", fontsize=14, fontweight="bold")

# ── 图2a：爆文内容类型分布（饼图）──────────────────────────────
ax = axes2[0]
# value_counts()：统计每种内容类型出现了多少次
top_type = top["内容类型"].value_counts()
wedge_colors = colors[:len(top_type)]
wedges, texts, autotexts = ax.pie(
    top_type.values,
    labels=top_type.index,
    autopct="%1.0f%%",      # 显示百分比，保留整数
    colors=wedge_colors,
    startangle=90,           # 从12点钟方向开始画
    pctdistance=0.75,        # 百分比文字距圆心的距离
    wedgeprops=dict(edgecolor="white", linewidth=2)  # 扇形之间的白色分割线
)
for at in autotexts:
    at.set_fontsize(11)
    at.set_fontweight("bold")
ax.set_title("爆文内容类型分布", fontsize=12, fontweight="bold")

# ── 图2b：爆文 vs 普通笔记 标题特征对比（分组柱状图）───────────
ax = axes2[1]
# 分别计算爆文和普通笔记中：含Emoji的比例、含数字的比例
# .eq("是")：判断每行是否等于"是"，返回True/False
# .mean()：True=1, False=0，mean()就是比例
metrics = {
    "含Emoji比例(%)": [
        top["是否含Emoji"].eq("是").mean() * 100,     # 爆文中含Emoji的比例
        bottom["是否含Emoji"].eq("是").mean() * 100,  # 普通笔记中含Emoji的比例
    ],
    "含数字比例(%)": [
        top["是否含数字"].eq("是").mean() * 100,
        bottom["是否含数字"].eq("是").mean() * 100,
    ],
}
x = range(len(metrics))
w = 0.35  # 每个柱子的宽度
# 爆文柱子向左偏移w/2，普通笔记向右偏移w/2，形成并排效果
bars1 = ax.bar([i - w/2 for i in x], [v[0] for v in metrics.values()],
               width=w, label="爆文 Top20%", color="#E8943A", edgecolor="white")
bars2 = ax.bar([i + w/2 for i in x], [v[1] for v in metrics.values()],
               width=w, label="普通笔记",    color="#94A3B8", edgecolor="white")
for bar in list(bars1) + list(bars2):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f"{bar.get_height():.0f}%", ha="center", va="bottom", fontsize=10, fontweight="bold")
ax.set_xticks(list(x))
ax.set_xticklabels(list(metrics.keys()))
ax.set_ylabel("占比(%)")
ax.set_title("爆文 vs 普通笔记 标题特征对比", fontsize=12, fontweight="bold")
ax.legend()
ax.set_facecolor("#F8F9FA")
ax.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
plt.savefig("爆文分析.png", dpi=150, bbox_inches="tight")
plt.close()
print("✅ 图表已保存：爆文分析.png")

# ══════════════════════════════════════════════════════════════
# 打印文字结论
# ══════════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  核心数据结论")
print("=" * 55)

print("\n【结论1】内容类型与互动表现")
for t, v in type_avg.items():
    print(f"  {t}：平均综合互动量 {v:.1f}")

# 计算Emoji带来的提升幅度：(有Emoji均值 - 无Emoji均值) / 无Emoji均值 × 100%
e_yes = emoji_avg.get("是", 0)
e_no  = emoji_avg.get("否", 0)
diff_e = ((e_yes - e_no) / e_no * 100) if e_no > 0 else 0
print(f"\n【结论2】标题含Emoji的笔记互动量比不含高 {diff_e:+.0f}%")
print(f"  含Emoji：{e_yes:.1f}  不含Emoji：{e_no:.1f}")

# 同理计算数字带来的提升幅度
n_yes = num_avg.get("是", 0)
n_no  = num_avg.get("否", 0)
diff_n = ((n_yes - n_no) / n_no * 100) if n_no > 0 else 0
print(f"\n【结论3】标题含数字的笔记互动量比不含高 {diff_n:+.0f}%")
print(f"  含数字：{n_yes:.1f}  不含数字：{n_no:.1f}")

# 自己账号 vs 竞品账号的差距
own_avg_val  = own["综合互动量"].mean()
comp_avg_val = comp["综合互动量"].mean()
gap = comp_avg_val / own_avg_val if own_avg_val > 0 else 0
print(f"\n【结论4】自己账号 vs 竞品账号")
print(f"  自己账号平均互动量：{own_avg_val:.1f}")
print(f"  竞品平均互动量：    {comp_avg_val:.1f}")
print(f"  差距倍数：竞品是自己账号的 {gap:.1f} 倍")

# 爆文特征总结
print(f"\n【结论5】爆文（Top20%，互动量≥{threshold:.0f}）共 {len(top)} 篇")
print(f"  主要内容类型：{top_type.index[0]}（占爆文 {top_type.values[0]/len(top)*100:.0f}%）")
print(f"  爆文中含Emoji比例：{top['是否含Emoji'].eq('是').mean()*100:.0f}%")
print(f"  爆文中含数字比例：{top['是否含数字'].eq('是').mean()*100:.0f}%")

# 自己账号里互动量最好的3篇，看看哪类内容相对表现好
print(f"\n【参考】自己账号综合互动量 Top 3：")
own_top = own.nlargest(3, "综合互动量")[["标题", "综合互动量", "内容类型"]]
for _, row in own_top.iterrows():
    print(f"  [{row['内容类型']}] {row['标题'][:25]}… 互动量={row['综合互动量']:.0f}")

print("\n" + "=" * 55)
print("  分析完成！请查看生成的两张图表。")
print("=" * 55)
