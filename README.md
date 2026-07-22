# Conference Tracker

追踪 CCF-A 类学术会议每年的举办时间、地点，核心是 **workshop proposal 提交截止时间**（申请举办 workshop 的 deadline，不是论文投稿 deadline）。

设计背景和取舍见 [PROJECT_BRIEF.md](PROJECT_BRIEF.md)。

## 数据在哪

[data/conferences.xlsx](data/conferences.xlsx)，两个 sheet：

- **会议数据**：58 个 CCF-A 会议，每行一条记录，字段见下表
- **变更日志**：每次自动运行如果有字段变化，会在这里追加一行（会议、字段名、旧值、新值、变更时间），方便回溯

| 列 | 说明 |
|---|---|
| conference_id | 唯一标识，小写（如 `ppopp`），配置增删/排序时用这个 |
| name / acronym | 全称 / 缩写 |
| year | 统一是 `TARGET_YEAR`（当前 2027） |
| category | 对应 CCF 分类（AI/CG/CT/DB/DS/HI/MX/NW/SC/SE） |
| location_city / location_country | 地点；该会议还没发布这一届信息则为空 |
| location_region | North America / Latin America / Europe / Asia / Other，**只记录不过滤** |
| date | `mm.dd-mm.dd`；官网自己都没定具体日期就保留原始描述或 `TBD` |
| official_url | 官网链接 |
| main_cfp_deadline | 主会论文投稿 deadline |
| workshop_proposal_deadline | **核心字段**，`yyyy-mm-dd` |
| workshop_proposal_status | `confirmed` / `not_yet_announced` / `not_applicable` |
| workshop_page_url / source_url | 抓到 workshop 信息的具体页面 |
| last_checked_at / last_changed_at | 最近一次检查 / 最近一次实际变化的时间 |
| notes | 备注，抓取异常（如疑似过期页面）会写在这里 |

## 两个脚本

```bash
pip install -r scripts/requirements.txt

# 1. 同步基础字段（地点/日期/主会deadline），数据来自 ccfddl/ccf-deadlines
python scripts/sync_base_info.py

# 2. 抓 workshop proposal deadline（核心），官网直连优先，猜不中用 Brave Search 兜底
python scripts/check_workshop_deadlines.py
python scripts/check_workshop_deadlines.py --only ppopp hpca   # 只跑指定会议，调试用
```

两个脚本分工明确：`sync_base_info.py` 只重建"基础信息"列，不会碰 `workshop_proposal_*` 相关列；`check_workshop_deadlines.py` 只更新这几列。每次 `sync_base_info.py` 跑完都会**整个重建**"会议数据" sheet，所以：

> **不要直接在 Excel 里删行、拖动顺序去调整会议列表**，下次同步会把这类结构性改动冲掉。`notes` 这种内容型单元格的手改是安全的，会被保留。

## 手动增删 / 调整会议顺序

编辑 [data/conference_overrides.json](data/conference_overrides.json)：

```json
{
  "excluded_ids": ["stoc", "focs"],
  "custom_order": ["ppopp", "hpca", "chi"]
}
```

- `excluded_ids`：填 `conference_id`，这些会议会从表里删掉，且不会被自动同步加回来
- `custom_order`：填想置顶的顺序，没写到的会议按字母序排在后面

改完跑一次 `sync_base_info.py`（或等下次 Actions 自动跑）生效。

## GitHub Actions 自动运行

[.github/workflows/check_deadlines.yml](.github/workflows/check_deadlines.yml)：

- **定时**：每周一 06:00 UTC（对应多伦多时间凌晨 2 点左右，夏令时/冬令时切换时有约 1 小时漂移，不影响周任务）
- 跑完依次执行 `sync_base_info.py` → `check_workshop_deadlines.py`，`data/conferences.xlsx` 有变化就自动 commit + push

**需要的仓库设置**（一次性）：
1. `Settings → Secrets and variables → Actions → Secrets` 里配置 `BRAVE_SEARCH_API_KEY`（不配也能跑，只是猜不中官网的会议会停在 `not_yet_announced`，notes 里提示需要人工核实）
2. `Settings → Actions → General → Workflow permissions` 选 **"Read and write permissions"**（否则最后一步 `git push` 会失败）

**手动触发**：GitHub 仓库页面 → `Actions` 标签页 → 左侧选 "Update conference deadlines" → 右侧 "Run workflow"

## 已知局限

- 只覆盖 CCF-A 类会议，且严格只展示 `TARGET_YEAR`（明年要用记得改 [scripts/sync_base_info.py](scripts/sync_base_info.py) 顶部的 `TARGET_YEAR` 常量）；官网还没发布这一届的会议，日期/地点/官网/deadline 都会是空或 `TBD`，这是预期行为不是 bug
- 正则抽取 deadline 无法保证 100% 准确，官网页面格式差异很大；已知一种没法完全自动规避的情况：官网页面标题年份和页面内实际日期对不上（如缓存/命名习惯问题），且刚好只差一年时抓取脚本无法识别，需要偶尔人工抽查"变更日志"
- Brave Search 兜底依赖第三方 API 的搜索结果质量，搜不到、搜到不相关页面都可能发生
