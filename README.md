# hebei-costing — 河北定额库与清单组价 Skill（WorkBuddy）

河北省建设工程计价的完整数据与出数链路：定额子目查询 → 综合单价 → 清单级组价
（分部分项 + 措施自动计取 + 安全文明 + 增值税）→ **默认格式报表**（组价表/综合单价分析表/编制说明三表）。

- 数据：河北省消耗量标准 **11 册、21,688 子目、154,252 消耗行**（含 14 号勘误、16 号降水调整、13 号安装补充子目）；
  清单项目库 1,378 条（GB/T 50854/50857/50858 + 河北补充）、计价规则 523 条（GB/T 50500-2024 + 冀建建市函〔2025〕154 号）。
- 适用：造价咨询单位编制**招标工程量清单 / 招标控制价**（河北省级口径）。

## 仓库结构

```
hebei-costing-skill/
├── SKILL.md                  # 技能入口（WorkBuddy 按此识别）
├── scripts/
│   ├── quota.py              # 查询/组价 CLI（纯 Python 标准库，3.10+，零安装）
│   ├── report.py             # 默认报表生成器（Excel 三表，需 openpyxl）
│   ├── 定额库.db              # SQLite 数据（约 37MB，须与脚本同级）
│   └── 清单示例*.json          # bill / 报表输入示例 4 份
└── references/
    ├── report-format.md      # ★默认输出格式规范 v1（组价成果出数格式）
    ├── usage.md              # 命令速查、清单 JSON 结构、计价口径
    └── handover-TODO.md      # 维护交接单（数据重建/踩坑，仅维护时读）
```

## 下载

方式一（git）：

```bash
git clone https://github.com/<你的用户名>/hebei-costing-skill.git
```

方式二（ZIP）：打开本仓库页面 → 绿色 **Code** 按钮 → **Download ZIP** → 解压。

## 安装（Windows / macOS / Linux 通用）

1. 将整个文件夹复制到 WorkBuddy 用户级技能目录，并命名为 `hebei-costing`：

   ```
   Windows:  C:\Users\<你>\.workbuddy\skills\hebei-costing\
   macOS:    ~/.workbuddy/skills/hebei-costing/
   Linux:    ~/.workbuddy/skills/hebei-costing/
   ```

   快捷命令（在解压/克隆目录的上一级执行）：

   ```bash
   # Windows (PowerShell)
   Copy-Item -Recurse hebei-costing-skill "$env:USERPROFILE\.workbuddy\skills\hebei-costing"
   # macOS / Linux
   cp -r hebei-costing-skill ~/.workbuddy/skills/hebei-costing
   ```

   目录校验：`~/.workbuddy/skills/hebei-costing/` 下应直接看到 `SKILL.md`、`scripts/`、`references/`。

2. （仅报表功能需要）安装 openpyxl：

   ```bash
   py -3 -m pip install openpyxl     # 或 python -m pip install openpyxl
   ```

3. 验证安装：

   ```bash
   cd ~/.workbuddy/skills/hebei-costing/scripts
   python quota.py stats             # 应输出 11 册 / 21,688 子目概况
   py -3 report.py 清单示例_组价表.json -o test.xlsx
   # 应输出：清单 7 条｜分部分项合计 39,968.75 元（格式验证锚点）
   ```

## 使用方法

### 在 WorkBuddy 对话中（推荐）

安装后无需任何配置，对话中直接提出计价类需求即自动触发技能：

- 「查定额：挖基坑土方，建筑册」「5-52 综合单价是多少」
- 「这 7 条清单帮我组价」「编一份招标工程量清单」
- 「生成组价表和综合单价分析表」（默认按 `references/report-format.md` 三表格式出 Excel）
- 「010102001 套哪些子目」「安全文明施工费怎么计」「风险条款能写无限风险吗」

### 命令行（直接调用）

```bash
cd ~/.workbuddy/skills/hebei-costing/scripts

python quota.py search 挖基坑 --book 建筑     # 子目检索
python quota.py show 5-52 --book 建筑         # 子目详情（消耗量+原书页码）
python quota.py price 5-52 --book 建筑        # 综合单价（除税）
python quota.py total 5-52 --book 建筑 --qty 10   # 合价（含安全文明/增值税）
python quota.py bill 清单.json                # 清单级组价全流程
python quota.py suggest 010102001             # 组价建议（候选子目+含量）
python quota.py bill-items --rules 无限风险    # 计价规则检索
python quota.py measure --profession 建筑      # 措施项目费率

py -3 report.py 清单示例_组价表.json -o 输出.xlsx   # 默认格式报表（Excel 三表）
```

完整命令与清单 JSON 结构见 `references/usage.md`。

## 输出格式（默认规范）

组价成果默认输出 Excel 三表（详见 `references/report-format.md`）：

1. **组价表** — 清单→定额→资源 三级层次 12 列，含编制说明块；
2. **综合单价分析表** — 表1-19，人工/材料/机械/管理费/利润组成；
3. **编制说明与组价思路** — 七节依据 + 逐条透明计算链。

聊天内答复组价结果同源适用（三级层次 + 计算链 + 出处标注）。

## 计价口径要点

- 综合单价**不含增值税**；管理费/利润基数 = **人工费+机械费**；措施费 =（人工+机械）× 费率%。
- 安全文明施工费按档位（园林单一档 4.5%）；「其他材料%」= 子目小计 × %。
- 定额子目**无基价**（2023 版量价分离），价格独立取自《河北建设工程造价信息》+ 2023 基期价。

## 数据边界与声明

- 仅覆盖**河北省**省级口径；价格默认雄安新区 2026-08 期 + 河北 2023 基期，可按期次/地区换。
- 数据由官方发布标准文本解析校核而成，**仅供学习与工程计价参考**，正式成果以官方发布文本为准；
  使用者应按现行计价依据复核后用于实际工程。
- 缺价材料（如古建油饰类）如实报缺，**工具不编造单价**；正式清单/控制价的编制与审核签章
  由注册造价工程师完成（GB/T 50500-2024 第 1.0.4 条）。

## 许可

数据与脚本仅供工程计价学习参考使用；转载或二次分发请保留出处并遵守相关标准的版权规定。
