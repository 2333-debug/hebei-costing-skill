# quota.py 命令与计价口径速查

运行目录：本 skill 的 `scripts/`（quota.py 与 定额库.db 必须同级）。quota.py 为 Python 3.10+ 纯标准库，零安装；report.py 另需 openpyxl。

## 常用命令

```bash
python scripts/quota.py stats                          # 库概况（册/子目/消耗行/缺价）
python scripts/quota.py search 挖基坑 --book 建筑      # 定额子目检索（FTS+别名；--book 限定册）
python scripts/quota.py show 5-52 --book 建筑          # 子目详情（五要素+消耗量+原文页码）
python scripts/quota.py price 5-52 --book 建筑         # 综合单价（除税），带缺价/废止提示
python scripts/quota.py total 5-52 --book 建筑 --qty 10        # 合价（含安全文明/增值税）
python scripts/quota.py bill 清单.json [--out 结果.json]      # 清单级组价（全流程）
python scripts/quota.py measure --profession 建筑      # 措施项目费率（05 附件1）
python scripts/quota.py bill-items 挖 --profession 建筑       # 清单项目库检索（国标+河北补充）
python scripts/quota.py bill-items --code 010102001    # 清单项目五要素精查
python scripts/quota.py bill-items --rules 无限风险    # 计价规则检索（17号+50500 共 523 条）
python scripts/quota.py suggest 010102001              # 组价建议：候选子目+建议含量
python scripts/quota.py suggest "现浇混凝土矩形柱" --profession 建筑
python scripts/quota.py map-add 010102001 --items "建筑:1-27:0.01" --feature-key "三类土|深2m"
python scripts/quota.py map-list [编码]                # 查已沉淀的清单→子目映射
python scripts/quota.py gaps --book 建筑 --top 10      # 缺价清单（材料/机械）
python scripts/quota.py alias-add <名称> --ref <id> [--kind machine]   # 别名映射（人工确认后）
python scripts/quota.py fees --profession 建筑         # 取费标准（管理费/利润）
py -3 scripts/report.py 清单示例_组价表.json -o 输出.xlsx   # ★默认输出格式：组价成果 Excel 三表
```

> 组价成果（组价表/综合单价分析表/编制说明）的**默认出数格式**见 `report-format.md`——
> 所有组价输出与聊天答复均须遵循；report.py 输入结构照抄 `scripts/清单示例_组价表.json`。

## bill 清单 JSON 结构

```json
{"工程": {"名称","专业","类别","税制","地区","期次","建筑面积","安全文明档位"},
 "价款调整": 0,
 "措施项目": [{"自动": true, "排除": ["A007"], "系数": {"A001": 0.5}},
              {"编号": "A001"}, {"名称": "冬施", "金额": 1234}],
 "其他项目": [{"名称","金额","含税": false}],
 "清单": [{"编号","清单编码":"010102001","名称","单位","工程量","特征键":"三类土|深2m",
           "子目": [{"册":"建筑","编号":"1-27","含量":0.01}]}]}
```

- 清单行**不写「子目」**时自动查 bill_map 映射（按「清单编码」+「特征键」，回落「默认」）。
- 措施项目四种形态：手填{名称,金额} / {编号[,系数,基数]} / {名称} / {自动:true[,排除,系数]}。
- 详细示例见 `scripts/清单示例*.json`（手填 / 措施自动 / 映射自动）。

## 计价口径（易错点，展开前先记）

- **综合单价不含增值税**；增值税基数 = 分部分项+措施+其他（**专业工程暂估价除外**，含税）。
- **措施项目费 =（人工费+机械费）× 费率%**（05 附件1），基数按「各专业工程」分列；
  人/材/机% 是**费率的组成拆分**（三项之和=费率），不是占比。
- **安全文明施工费**：基数 = 直接费（不含安全文明、设备费）+企管+利润+价款调整；市区/县城外 ×0.97；
  简易计税 ×0.94。园林单一档 **4.5%**（06 册附件2），其余各专业按档位（建筑面积/造价）。
- **管理费/利润基数 = 人工+机械**（不是人材机）。"其他材料/其他机械 %" = 本子目小计 × %。
- 子目**已废止**（如 16 号降水调整前的井点子目）price/total 会带 ⚠ 提示，勿用。
- **条件措施不自动计**（施工与生产同时/有害环境/超高/垂直运输/操作高度等），只提示供人选用。
- 册名≠取费专业：`安装补充` 按 `安装` 取费（BOOK_PROF 映射）；新册缺映射会报错，不会静默套错。
- 单位换算：`_unit_factor(价格单位,消耗单位)` 语义是 元/price×factor=元/cons；块→百块=×100。

## 数据边界

- 11 册定额（建筑/装饰/安装/市政/园林/修缮安装/市政养护/修缮建筑/仿古建筑/古建修缮/安装补充），
  21,688 子目、154,252 消耗行、100% 可复核；变更台账 320 条（勘误/降水/补充）。
- 价格：材料 12,288（雄安 2026-08 + 河北 2023 基期）+ 台班 1,476。缺价以 gaps 为准（古建油饰类真无价）。
- 清单库：bill_item 1,378（国标 50854/50857/50858 + 河北补充），bill_rule 523（17 号 + 50500）。
- 维护/重建流程见 `references/handover-TODO.md`（含全部踩坑记录）。
