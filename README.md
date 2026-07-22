# Conference Tracker

追踪一批学术会议每年的举办时间、地点，核心是 **workshop proposal 提交截止时间**（申请举办 workshop 的 deadline，不是论文投稿 deadline）。

设计背景和取舍见 [PROJECT_BRIEF.md](PROJECT_BRIEF.md)。

## 数据在哪

[data/Conferences.xlsx](data/Conferences.xlsx)，sheet **`2027A类会议`**，61 个会议，每行一条记录：

| 列 | 说明 |
|---|---|
| 方向 | 领域分类（系统/网络/安全/软工/理论/数据库/AI/交互/多媒体/综合/半导体...） |
| 刊物简称 / 刊物全称 | 缩写 / 全称 |
| 出版方 | ACM / IEEE / USENIX / Springer 等 |
| 会议地点 | `City, (State), Country`，还没公布就是 `TBD` |
| 会议时间 | `mm.dd-mm.dd`，还没公布就是 `TBD` |
| Workshop Deadline | **核心字段**，`yyyy.mm.dd`，还没公布就是 `TBD`；确定或抓取到的都会自动做成可点击超链接对应的来源页 |
| 会议URL | 官网链接（超链接） |
| workshopURL | 抓到 workshop 信息的具体页面（超链接） |
| 备注 | 抓取来源/异常说明，走 Brave Search 兜底抓到的会标"需要人工审核" |

另有一个 `变更日志` sheet，`check_workshop_deadlines.py` 每次跑如果 Workshop Deadline 有变化就追加一行，方便回溯。

会议的**数量、名字、顺序完全由你在 xlsx 里手动维护**，两个脚本都不会自动增删或重排行。

## 两个脚本

```bash
pip install -r scripts/requirements.txt

# 1. 同步基础字段（地点/日期/官网），数据来自 ccfddl/ccf-deadlines
python scripts/sync_base_info.py

# 2. 抓 workshop proposal deadline（核心），官网直连优先，猜不中用 Brave Search 兜底
python scripts/check_workshop_deadlines.py
python scripts/check_workshop_deadlines.py --only PPoPP HPCA   # 只跑指定会议，按"刊物简称"，调试用
```

**行为规则**：`sync_base_info.py` 只根据"刊物简称"去 ccfddl 匹配对应会议，刷新它的"会议地点"/"会议时间"/"会议URL"；匹配不上（不在 CCF 名单里、你手动加的会议，如 ISSCC/IEDM/VLSI/WINE）原样保留，不碰。`check_workshop_deadlines.py` 则对 xlsx 里**当前所有行**一视同仁地去抓 workshop deadline，不管这行是不是在 ccfddl 里能找到。

> **不要直接在 Excel 里改"会议地点"/"会议时间"/"会议URL"这三列**——下次 `sync_base_info.py` 一跑就会被 ccfddl 的数据覆盖掉（除非该会议不在 ccfddl 里）。想手动订正错误数据，用下面的覆盖表。

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

没有任何一列能保证 100% 准，越关键的信息越应该点开 `会议URL` / `workshopURL` 这两个超链接亲自确认：
- 备注写 `extracted from official site` —— 直接从官网正则抠出来的，相对可信但也可能抓错
- 备注写 `...搜索兜底...需要人工审核` —— 走的是 Brave 搜索，没验证过页面是否确属该会议
- `TBD` / `需要人工核实官网` —— 脚本自己都没把握，别当真

## GitHub Actions 自动运行

[.github/workflows/check_deadlines.yml](.github/workflows/check_deadlines.yml)：

- **定时**：每周一 06:00 UTC（对应多伦多时间凌晨 2 点左右，夏令时/冬令时切换时有约 1 小时漂移，不影响周任务）
- 跑完依次执行 `sync_base_info.py` → `check_workshop_deadlines.py`，`data/Conferences.xlsx` 有变化就自动 commit + push

**需要的仓库设置**（一次性）：
1. `Settings → Secrets and variables → Actions → Secrets` 里配置 `BRAVE_SEARCH_API_KEY`（不配也能跑，只是猜不中官网的会议会停在 `TBD`，备注里提示需要人工核实）
2. `Settings → Actions → General → Workflow permissions` 选 **"Read and write permissions"**（否则最后一步 `git push` 会失败）

**手动触发**：GitHub 仓库页面 → `Actions` 标签页 → 左侧选 "Update conference deadlines" → 右侧 "Run workflow"

## 已知局限

- 只覆盖 xlsx 里你手动确定的这份清单，严格只展示 `TARGET_YEAR`（明年要用记得改 [scripts/sync_base_info.py](scripts/sync_base_info.py) 顶部的 `TARGET_YEAR` 常量）
- ccfddl 数据偶尔滞后/出错，靠 `LOCATION_OVERRIDES`/`DATE_OVERRIDES` 订正已知问题
- 正则抽取 Workshop Deadline 无法保证 100% 准确，官网页面格式差异很大，已知会把不相关的日期（如 artifact 提交日期）误判成 workshop deadline 的情况发生过，已修复但不保证以后不会有新的类似情况
- Brave Search 兜底依赖第三方 API 的搜索结果质量，且现在（2026年年中）大部分 2027 届会议本身还没公布 workshop 信息，搜不到不代表脚本有问题
