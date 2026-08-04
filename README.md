# Conference Tracker

追踪一批学术会议每年的举办时间、地点，核心是 **workshop proposal 提交截止时间**（申请举办 workshop 的 deadline，不是论文投稿 deadline）。

设计背景和取舍见 [PROJECT_BRIEF.md](PROJECT_BRIEF.md)（早期版本，当时还只覆盖 58 个 CCF-A 会议；现在清单已扩展到 85 个，含部分 CCF-B/C 会议，历史背景仍然有效）。

## 数据在哪

[data/Conferences.xlsx](data/Conferences.xlsx)，sheet **`2027A类会议`**，85 个会议，每行一条记录：

| 列 | 说明 | 谁维护 |
|---|---|---|
| 方向 | 领域分类（系统/网络/安全/软工/理论/数据库/AI/交互/多媒体/综合/半导体...） | 人工 |
| 刊物简称 / 刊物全称 | 缩写 / 全称 | 人工 |
| 会议地点（2027） | `City, (State), Country`，还没公布就是 `TBD` | 脚本 |
| 会议时间（2027） | `mm.dd-mm.dd`，还没公布就是 `TBD` | 脚本 |
| 所属地区（2027） | 北美/欧洲/亚太/拉美/非洲，根据"会议地点"自动归类，还没公布就是 `TBD` | 脚本 |
| Workshop Submission Deadline（2027） | **核心字段**，`yyyy.mm.dd`，还没公布就是 `TBD` | 脚本 |
| 出版方 | ACM / IEEE / USENIX / Springer 等 | 人工 |
| URL | dblp 链接（长期稳定，不随届次变化） | 人工 |
| 2027 URL | 本届官网链接（超链接） | 脚本 |
| 2027 Workshop Proposal URL | 抓到 workshop 信息的具体页面（超链接） | 脚本 |

另有一个 `变更日志` sheet，`check_workshop_deadlines.py` 每次跑如果 Workshop Deadline 有变化就追加一行，方便回溯。

会议的**数量、名字、顺序完全由你在 xlsx 里手动维护**，两个脚本都不会自动增删或重排行，也不会碰"方向"/"刊物简称"/"刊物全称"/"出版方"/"URL"这几列。

## 两个脚本

```bash
pip install -r scripts/requirements.txt

# 1. 同步基础字段（地点/时间/地区/本届官网），ccfddl 为主，官网直连兜底
python scripts/sync_base_info.py

# 2. 抓 workshop proposal deadline（核心）+ workshop 页面链接，直连官网
python scripts/check_workshop_deadlines.py
python scripts/check_workshop_deadlines.py --only PPoPP HPCA   # 只跑指定会议，按"刊物简称"，调试用
```

**行为规则**：

`sync_base_info.py` 只更新"会议地点（2027）"/"会议时间（2027）"/"所属地区（2027）"/"2027 URL"这四栏：
- 按"刊物简称"能在 ccfddl（[ccfddl/ccf-deadlines](https://github.com/ccfddl/ccf-deadlines)，众包整理自各会议官网的开源项目）里匹配上、且有 2027 届数据，就用 ccfddl 的数据刷新，这是主要数据源
- 匹配不上（不在 ccfddl 名单里的会议，如 ISSCC/IEDM/VLSI/WINE），或匹配上了但还没有 2027 届数据：退化到直接抓"2027 URL"里已有的官网链接，从页面正文里抠日期/地点；抓不到就保留原值不动，不会打回 TBD
- "所属地区"是脚本根据"会议地点"里的国家名自动归类的（国家词表见 `sync_base_info.py` 里的 `COUNTRY_REGION`），遇到没见过的国家名会保留原有地区值，不强行猜
- "Workshop Submission Deadline（2027）" / "2027 Workshop Proposal URL" 这两列不碰，那部分由 `check_workshop_deadlines.py` 维护

`check_workshop_deadlines.py` 则对 xlsx 里**当前所有行**一视同仁地去抓 workshop deadline 和 workshop 页面链接，不管这行是不是在 ccfddl 里能找到。**Brave Search 兜底目前在代码里硬关闭**（`ENABLE_BRAVE_SEARCH = False`），只走官网直连；抓不到就是 `TBD`。

> **不要直接在 Excel 里改"会议地点（2027）"/"会议时间（2027）"/"所属地区（2027）"/"2027 URL"这四列**——下次 `sync_base_info.py` 一跑，只要 ccfddl 有数据就会被覆盖掉（不在 ccfddl 名单里的会议改了倒是安全，因为脚本抓不到才会保留原值）。想手动订正错误数据，用下面的覆盖表。

## 手动增删 / 调整会议顺序

**直接在 xlsx 里改**——删行、拖顺序、加一行手打的会议都可以，`conference_id` 之类的辅助字段不需要维护。加一行新会议时，"刊物简称"填对应缩写即可，其余基础字段自己填或留空；下次跑完整流程（`sync_base_info.py` 然后 `check_workshop_deadlines.py`）会自动帮你把 Workshop Deadline 等列补上。

## 订正 ccfddl 的错误数据

ccfddl 是众包维护的项目，个别字段可能滞后或出错（比如实测发现过 HPCA 2027 的日期是错的）。发现问题不用去改 ccfddl 上游，在 [scripts/sync_base_info.py](scripts/sync_base_info.py) 顶部加进 `LOCATION_OVERRIDES` / `DATE_OVERRIDES` 就行，以后每次同步都不会被错误数据覆盖回去：

```python
LOCATION_OVERRIDES = {"chi": "Pittsburgh, (PA), USA"}
DATE_OVERRIDES = {"hpca": "03.20-03.24"}
```

key 是"刊物简称"小写去掉符号后的样子（比如 `S&P` -> `sp`）。

## 怎么判断抓到的信息是不是对的

没有任何一列能保证 100% 准，越关键的信息越应该点开 `2027 URL` / `2027 Workshop Proposal URL` 这两个超链接亲自确认。目前没有单独的"备注"列标注抓取来源可信度了，判断标准简化为：
- 官网直连正则抠出来的，相对可信，但也可能抓错（尤其地点/日期的抽取规则比较宽松）
- `TBD` —— 脚本自己都没抓到，别当真

## GitHub Actions 自动运行

[.github/workflows/check_deadlines.yml](.github/workflows/check_deadlines.yml)：

- **定时**：每周一 06:00 UTC（对应多伦多时间凌晨 2 点左右，夏令时/冬令时切换时有约 1 小时漂移，不影响周任务）
- 跑完依次执行 `sync_base_info.py` → `check_workshop_deadlines.py`，`data/Conferences.xlsx` 有变化就自动 commit + push

**需要的仓库设置**（一次性）：
1. `Settings → Actions → General → Workflow permissions` 选 **"Read and write permissions"**（否则最后一步 `git push` 会失败）
2. Brave Search 兜底目前不需要配置任何 secret（代码里硬关闭了）

**手动触发**：GitHub 仓库页面 → `Actions` 标签页 → 左侧选 "Update conference deadlines" → 右侧 "Run workflow"

## 关于 References*.xlsx

仓库根目录下如果出现 `References*.xlsx`，那是本地手动整理用的参考表格，已经加入 `.gitignore`，不会上传、也不参与自动化流程——两个脚本只读写 `data/Conferences.xlsx`。

## 已知局限

- 只覆盖 xlsx 里你手动确定的这份清单，严格只展示 `TARGET_YEAR`（明年要用记得改 [scripts/sync_base_info.py](scripts/sync_base_info.py) 顶部的 `TARGET_YEAR` 常量，两个脚本里都有）
- ccfddl 数据偶尔滞后/出错，靠 `LOCATION_OVERRIDES`/`DATE_OVERRIDES` 订正已知问题
- 官网直连兜底（ccfddl 没有这个会议，或者有会议但还没有 2027 届数据）会先按"最近一届 link 里的年份规律"猜一个 2027 候选 URL（比如 `pldi26.sigplan.org` -> `pldi27.sigplan.org`），猜出来的 URL 会先抓一下确认页面正文里真的出现"2027"才采信，猜不中/没有历史规律可猜就退化到抓"2027 URL"里已有的链接；地点抽取额外要求"2027"必须出现在匹配位置附近，避免把官网首页还没更新的旧一届信息当成新一届（实测抓到过 CoRL 这种官网不换域名、首页还停留在上一届信息的情况）。抽取规则本身还是比较简单，很多官网页面格式抠不出来，抠不出来会保留原值，不会主动报错
- 正则抽取 Workshop Deadline 无法保证 100% 准确，官网页面格式差异很大，已知会把不相关的日期（如 artifact 提交日期）误判成 workshop deadline 的情况发生过，已修复但不保证以后不会有新的类似情况
- Brave Search 兜底已临时关闭：现在（2026年年中）大部分 2027 届会议本身还没公布 workshop 信息，之前测试发现搜出来的假阳性比真结果多
