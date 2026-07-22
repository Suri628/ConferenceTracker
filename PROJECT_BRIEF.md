# 会议 Workshop Deadline 追踪器 — 项目说明

## 项目目标
定期追踪 **CCF-A 类**学术会议每年的举办时间、地点，以及（核心）**workshop proposal 提交截止时间**（即"申请举办 workshop"的 deadline，不是 workshop 论文投稿 deadline，也不是主会投稿 deadline）。

不做地区过滤：`location_region` 字段保留、记录，但不作为筛选条件（未来想筛北美/拉美等地区时，数据已经现成）。

## 数据来源与分工

发现 [ccfddl/ccf-deadlines](https://github.com/ccfddl/ccf-deadlines) 是一个持续维护（每天都有更新）的开源项目，按 CCF 官方目录结构化存储了全部会议及 rank。据此拉出完整 **58 个 CCF-A 类会议**清单（已核实排除了常见误解的会议，如 IJCAI 实际是 CCF-B）。

这个仓库能提供的字段，和我们还需自己做的部分：

| 数据 | 来源 |
|---|---|
| 会议名称/缩写/分类(sub)/CCF rank/dblp | ccfddl 仓库现成数据，同步转换即可 |
| 地点、日期、主会论文投稿 deadline | ccfddl 仓库现成数据，同步即可 |
| **workshop proposal deadline**（本项目核心目标） | **ccfddl 完全不追踪**，需要我们自己抓官网 |

结论：不用重复造轮子做基础信息爬虫，把精力集中在 workshop proposal deadline 抓取这一独有价值上。

## CCF-A 完整会议清单（58个，全部收录）

| 分类 | 会议 |
|---|---|
| 人工智能 (AI) | AAAI, ACL, CVPR, ICCV, ICLR, ICML, NeurIPS |
| 图形学/多媒体 (CG) | ACM MM, ACM SIGGRAPH, IEEE VIS, IEEE VR |
| 理论 (CT) | CAV, FOCS, LICS, SODA, STOC |
| 数据库/数据挖掘 (DB) | ICDE, SIGIR, SIGKDD, SIGMOD, VLDB |
| 体系结构/存储/并行 (DS) | ASPLOS, DAC, EuroSys, FAST, HPCA, HPDC, ISCA, MICRO, PPoPP, SC, SIGOPS ATC |
| 人机交互 (HI) | CHI, CSCW, UIST, UbiComp/ISWC |
| 交叉/综合 (MX) | RTSS, WWW |
| 网络 (NW) | INFOCOM, MobiCom, NSDI, SIGCOMM |
| 安全 (SC) | CCS, CRYPTO, EUROCRYPT, NDSS, S&P, USENIX Security |
| 软工/系统/程序语言 (SE) | ASE, FM, FSE, ICSE, ISSTA, OOPSLA, OSDI, PLDI, POPL, SOSP |

> 很多理论/密码类会议（STOC/FOCS/CRYPTO/EUROCRYPT等）大概率没有 workshop proposal 机制，抓取后预期会是 `not_applicable`，这是正常结果，不代表抓取失败。

## 数据模型（每个会议一条记录）
```
conference_id                  # 唯一标识，如 "ppopp"
name                            # 全称
acronym                         # PPoPP
year                             # 届次年份
category                         # 对应 ccfddl 的 sub 分类（AI/CG/CT/DB/DS/HI/MX/NW/SC/SE）
location_city
location_country
location_region                  # North America / Latin America / Europe / Asia / Other（记录但不过滤）
date_start
date_end
official_url
main_cfp_deadline                 # 主会论文投稿 deadline（同步自 ccfddl）
workshop_proposal_deadline          # 核心字段：举办者申请办 workshop 的 deadline（自行抓取）
workshop_proposal_status            # confirmed / not_yet_announced / not_applicable
workshop_page_url
last_checked_at
last_changed_at
source_url                          # 抓到这条信息的具体页面链接
notes                                # 备注（原"状态"列改名，保留括号内容，如"只有Forum/Tutorial"）
```

## 存储格式：xlsx

`data/conferences.xlsx`，两个 sheet：

- **`会议数据`**：主表，每行一个会议，字段同上。
- **`变更日志`**：每次跑脚本时若字段有变化则追加一行（会议、字段名、旧值、新值、变更时间），不污染主表，方便回溯。

## 抓取逻辑（伪代码）

```python
# 1. 基础字段同步（每次都做，成本低）
sync_base_fields_from_ccfddl()   # location/date/main_cfp_deadline/rank

# 2. workshop proposal deadline 抓取（核心，需要网络请求）
for conf in ccf_a_list:
    guess_url = pattern_from_last_year(conf, this_year)
    page = fetch(guess_url) or fallback_search(conf, "workshop proposal deadline")
    extracted = parse_deadline(page)   # 优先正则，页面不规整再考虑 LLM 辅助抽取
    if extracted != db[conf].workshop_proposal_deadline:
        db[conf].update(extracted)
        log_change(conf, old_value, new_value)
save(db)
```

## URL 模式识别（猜测当年 workshop 页面地址）
- `https://{conf}{yy}.sigplan.org/track/{conf}-{year}-workshops`
- `https://{conf}{year}.kdd.org/call-for-workshop-proposals/`
- `https://{conf}.acm.org/authors/workshops/`
- 策略：优先尝试"去年 URL 替换年份"，猜不中再退化到搜索引擎兜底

## 数据源分层（workshop deadline 抓取用）
**一级源（官网，唯一可信）**
- ACM 系：`conf.researchr.org`、`{acronym}.sigplan.org`
- ACM CCS：`www.sigsac.org/ccs/CCS{year}`
- USENIX 系：`usenix.org/conference/{acronym}{yy}`
- IEEE 系：各会议独立域名
- AAAI：`aaai.org/conference/aaai/aaai-{yy}`

**二级源（交叉验证/提前预警用，非唯一依据）**
- wikicfp.com
- ai-deadlines（HuggingFace 项目）

## 运行方式
- GitHub Actions 定时任务（每周一次），免费额度足够
- 结果存 xlsx，变更记录单独在"变更日志" sheet，方便回溯
- （可选）GitHub Pages 生成简单静态展示页

## 后续可扩展方向（暂不实现，先预留字段/接口）
- 主会论文投稿 deadline 展示（已同步自 ccfddl，可直接展示）
- workshop 论文本身投稿 deadline
- 通知渠道：邮件 / 日历订阅（.ics）/ Slack
- 个人关注清单过滤（例如只看某几个分类的会议）
- 未来3个月 deadline 一览周报
- CORE 分级、录用率等更多元数据（ccfddl 部分会议已有 core 字段，可顺带同步）

## 给 Claude Code 的起始任务
1. 目录结构：`data/`、`scripts/`
2. 基础信息同步脚本 `scripts/sync_base_info.py`：拉取 ccfddl/ccf-deadlines 中 58 个 CCF-A 会议数据，生成初始 `data/conferences.xlsx`
3. workshop deadline 抓取脚本 `scripts/check_workshop_deadlines.py`：先支持手动运行，验证能跑通 1-2 个会议（建议先从 PPoPP/HPCA 这种"未公布"的开始试）
4. 本地跑通后再配置 GitHub Actions 定时任务（`.github/workflows/check_deadlines.yml`，每周一次）
5. （可选）生成一个简单的静态展示页（HTML 表格即可，不需要花哨）
