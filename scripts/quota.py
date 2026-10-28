# -*- coding: utf-8 -*-
"""定额库客户端：检索 / 查看 / 组价 / 导出。

> 这是给"后期调用"准备的唯一入口：脚本、Agent、Excel 插件都走这里。

用法（命令行）：
  python quota.py stats
  python quota.py fees [--profession 市政]                 # 费用标准一览
  python quota.py search 钢筋 [--book 建筑] [--limit 20]
  python quota.py show 5-52 [--book 建筑]
  python quota.py price 5-52 [--book 建筑] [--region 雄安新区] [--period 2026-08]
                          [--profession 建筑] [--fee 一般-二类] [--tax 一般计税]
  python quota.py total 5-52 [--book 建筑] [--qty 10] [--safety-cat 档位] [--build-area 12000]
  python quota.py bill 清单.json [--out 结果.json]        # 清单级：多条子目 → 汇总到含税造价
  python quota.py export out.jsonl [--book 建筑] [--only-ok]

费用标准按「专业 × 工程类别 × 税制」取，河北计费基数是 直接费中人工费+机械费。
市政必须指出 4 类之一：--fee 土石方及大型机械 / 一般市政 / 道路工程 / 路灯工程。
"""
import argparse, json, os, sqlite3, re, sys

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "定额库.db")

GRADE_OF = {"综合用工一类": "一类", "综合用工二类": "二类", "综合用工三类": "三类"}
# 各专业最常用的默认取费类别（命令行不指定 --fee 时用）
#   园林绿化在 06 册自带费用标准，分「园林工程 / 绿化工程」两类，与 05 册的四个专业不是一套
DEFAULT_FEE = {"建筑": "一般-二类", "装饰": "一般", "安装": "二类", "市政": "一般市政",
               "园林": "园林工程", "修缮安装": "一般", "市政养护": "道路桥梁排水泵站河道",
               "修缮建筑": "一般", "仿古建筑": "一般", "古建修缮": "一般"}
# 册名 ≠ 取费专业时的映射（13 号补充子目是安装工程的补充，无独立费用标准，按安装取费）
BOOK_PROF = {"安装补充": "安装"}


class QuotaDB:
    def __init__(self, path=DB):
        self.con = sqlite3.connect(path)
        self.con.row_factory = sqlite3.Row
        self._mp_cache = {}

    # ---------- 检索 ----------
    def stats(self):
        q = self.con.execute
        return {
            "册": dict(q("SELECT book_id,count(*) FROM item GROUP BY book_id").fetchall()),
            "子目": q("SELECT count(*) FROM item").fetchone()[0],
            "可复核(ok)": q("SELECT count(*) FROM item WHERE review_status='ok'").fetchone()[0],
            "存疑(suspect)": q("SELECT count(*) FROM item WHERE review_status='suspect'").fetchone()[0],
            "未复核(unreviewed)": q("SELECT count(*) FROM item WHERE review_status='unreviewed'").fetchone()[0],
            "缺计量单位": q("SELECT count(*) FROM item WHERE unit IS NULL OR trim(unit)=''").fetchone()[0],
            "名称未解析": q("SELECT count(*) FROM item WHERE name LIKE '(未解析)%'").fetchone()[0],
            "消耗量行": q("SELECT count(*) FROM consumption").fetchone()[0],
            "未计价主材行": q("SELECT count(*) FROM consumption WHERE is_unpriced=1").fetchone()[0],
            "材料价(期次)": q("SELECT group_concat(DISTINCT region||'-'||period) FROM material_price").fetchone()[0],
            "材料价条数": q("SELECT count(*) FROM material_price").fetchone()[0],
            "机械台班价": q("SELECT count(*) FROM machine_price").fetchone()[0],
            "费用标准组": q("SELECT count(*) FROM fee_rate").fetchone()[0],
            "安全文明档位": q("SELECT count(*) FROM safety_fee").fetchone()[0],
            "变更台账": q("SELECT count(*) FROM change_log").fetchone()[0],
            "别名": q("SELECT count(*) FROM alias").fetchone()[0],
        }

    def search(self, keyword, book=None, limit=20):
        kw = (keyword or "").strip()
        if not kw:
            return []
        rows = []
        # 1) 编码 / 编码前缀
        if re.fullmatch(r"\d{1,2}-[0-9A-Za-z\-]{0,6}", kw):
            sql = "SELECT item_id,book_id,code,name,unit,page_no,review_status FROM item WHERE code LIKE ?"
            args = [kw + "%" if kw[-1] != "%" else kw]
            if book:
                sql += " AND book_id=?"
                args.append(book)
            rows = self.con.execute(sql + " ORDER BY book_id,code LIMIT ?", args + [limit]).fetchall()
        # 2) 全文（trigram 需 >=3 字）+ 材料反查
        if not rows:
            if len(kw) >= 3:
                try:
                    sql = """SELECT i.item_id,i.book_id,i.code,i.name,i.unit,i.page_no,i.review_status
                             FROM item_fts f JOIN item i ON i.item_id=f.item_id
                             WHERE item_fts MATCH ?"""
                    args = [f'"{kw}"']
                    if book:
                        sql += " AND i.book_id=?"
                        args.append(book)
                    rows = self.con.execute(sql + " ORDER BY rank LIMIT ?",
                                            args + [limit]).fetchall()
                except sqlite3.OperationalError:
                    rows = []
            if not rows:
                like = f"%{kw}%"
                # 括号必须加：原来写成 name LIKE ? OR c.name LIKE ? AND book_id=?
                # 优先级是 AND > OR，--book 只约束到了后半句，名称命中的行会跨册漏进来
                sql = """SELECT DISTINCT i.item_id,i.book_id,i.code,i.name,i.unit,i.page_no,i.review_status
                         FROM item i LEFT JOIN consumption c ON c.item_id=i.item_id
                         WHERE (i.name LIKE ? OR c.name LIKE ?)"""
                args = [like, like]
                if book:
                    sql += " AND i.book_id=?"
                    args.append(book)
                rows = self.con.execute(sql + " ORDER BY i.book_id,i.code LIMIT ?", args + [limit]).fetchall()
        return [dict(r) for r in rows]

    # ---------- 查看 ----------
    def show(self, code, book=None):
        sql = "SELECT * FROM item WHERE code=?"
        args = [code]
        if book:
            sql += " AND book_id=?"
            args.append(book)
        it = self.con.execute(sql + " ORDER BY book_id LIMIT 1", args).fetchone()
        if not it:
            return None
        out = dict(it)
        out["消耗量"] = {"人工": [], "材料": [], "机械": []}
        for r in self.con.execute("""SELECT kind,name,unit,qty,qty_raw,is_unpriced
                                     FROM consumption WHERE item_id=? ORDER BY kind,sort_no""", (it["item_id"],)):
            out["消耗量"][{"labor": "人工", "material": "材料", "machine": "机械"}[r["kind"]]].append(
                {"名称": r["name"], "单位": r["unit"], "消耗量": r["qty"], "原文": r["qty_raw"],
                 "未计价": bool(r["is_unpriced"])})
        return out

    # ---------- 价格匹配 ----------
    def _price_row(self, name, region, period):
        # 1) 别名（人工确认过的映射）
        r = self.con.execute("""SELECT p.* FROM alias a JOIN material_price p ON p.id=a.ref_id
                                WHERE a.scope='material' AND a.term=? AND p.region=? AND p.period=?
                                LIMIT 1""", (name, region, period)).fetchone()
        if r:
            return r
        # 2) 精确名称
        rows = self.con.execute("SELECT * FROM material_price WHERE name=? AND region=? AND period=?",
                                (name, region, period)).fetchall()
        if len(rows) == 1:
            return rows[0]
        # 3) 名称包含（唯一命中才算）
        rows = self.con.execute("""SELECT * FROM material_price WHERE region=? AND period=?
                                   AND (name LIKE ? OR ? LIKE '%'||name||'%')""",
                                (region, period, f"%{name}%", name)).fetchall()
        return rows[0] if len(rows) == 1 else None

    # ---------- 单位换算（价格单位 → 定额消耗单位） ----------
    # 语义：价格单位 → 消耗单位（元/price_unit × factor = 元/cons_unit）。
    # 例：元/kg → 元/t 是 ×1000（1t=1000kg）；元/m² → 元/100m² 是 ×100。
    # ⚠ 曾经 8 对方向写反（元/km→元/m 写成 1000，应 0.001），2026-10-27 修正。
    #   修正时实测影响面为 0（命中这些单位的材料当时都缺价，换算路径休眠），回归一字不差。
    #   不要再手写这些对——通用「倍数+基础单位」路径（split）方向是对的，只留跨量纲的 kg/t、km/m。
    CONV = {("kg", "t"): 1000.0, ("t", "kg"): 0.001,
            ("km", "m"): 0.001, ("m", "km"): 1000.0}

    @staticmethod
    def _unit_factor(price_unit, cons_unit):
        """价格单位 → 消耗单位的换算系数；无法换算返回 None。

        先查 CONV，再按「倍数 + 基础单位」通用处理（10套↔套、100个↔个、100kg↔kg…），
        中文倍数（百块/千块/万块）也认；m²/m2 这类写法差异先归一化再比。
        """
        pu = QuotaDB.norm_key(price_unit)
        cu = QuotaDB.norm_key(cons_unit)
        if pu == cu:
            return 1.0
        conv = {(QuotaDB.norm_key(a), QuotaDB.norm_key(b)): v for (a, b), v in QuotaDB.CONV.items()}
        if (pu, cu) in conv:
            return conv[(pu, cu)]

        CN_MULT = {"百": 100.0, "千": 1000.0, "万": 10000.0}

        def split(u):
            m = re.match(r"^(\d+(?:\.\d+)?)?\s*(.*)$", u)
            f = float(m.group(1) or 1)
            b = (m.group(2) or "").strip()
            if len(b) > 1 and b[0] in CN_MULT:      # 百块 / 千块 / 万块
                f *= CN_MULT[b[0]]
                b = b[1:]
            return f, b

        f1, b1 = split(pu)
        f2, b2 = split(cu)
        if b1 and b1 == b2:
            return f2 / f1        # 价格是「每 f1 个基础单位」，消耗按「每 f2 个」计
        return None

    # 材料名归一化：定额写 2.5mm2、信息价写 2.5mm²；直径符号有希腊/西里尔两种；
    # 还有全角括号、全角ｍ。只做写法归一，不做语义猜测。
    NORM_MAP = {"²": "2", "³": "3", "㎡": "m2", "㎥": "m3", "ｍ": "m", "Ｍ": "M",
                "／": "/", "（": "(", "）": ")", "－": "-", "　": " ", "．": ".",
                "Φ": "φ", "ϕ": "φ", "Ф": "φ", "ф": "φ"}

    @classmethod
    def norm_key(cls, s):
        s = "".join(cls.NORM_MAP.get(ch, ch) for ch in (s or ""))
        return re.sub(r"\s+", "", s).upper()

    def _material_index(self, reg, per):
        """归一化名称 → 价格行 的索引，一次建好（gaps 会扫几千个名字，逐条 LIKE 太慢）。"""
        key = ("mpidx", reg, per)
        if key not in self._mp_cache:
            idx = {}
            for r in self._price_rows("material_price", reg, per):
                idx.setdefault(self.norm_key(r["name"]), []).append(r)
            self._mp_cache[key] = idx
        return self._mp_cache[key]

    def _material_price(self, name, unit, region, period):
        """返回 (单价, 出处)。同单位优先；跨单位按 CONV 换算；本地期次缺则回落到河北 2023 基期。"""
        for reg, per in ((region, period), ("河北", "2023基期")):
            rows = self.con.execute("""SELECT p.* FROM alias a JOIN material_price p ON p.id=a.ref_id
                                       WHERE a.scope='material' AND a.term=? AND p.region=? AND p.period=?""",
                                    (name, reg, per)).fetchall()
            if rows:
                # 别名是人工确认过的映射，直接采用（单位差异按 _unit_factor 换算），
                # 不再受下面「同单位才用」那道门的限制——那道门是防模糊匹配配错的。
                cand = rows[0]
                price = cand["price_notax"] if cand["price_notax"] is not None else cand["price_tax"]
                if price is None:
                    continue
                factor = self._unit_factor(cand["unit"], unit)
                if factor is None:
                    factor = 1.0
                    tip = f"（别名映射；单位 {cand['unit'] or '—'}→{unit} 未换算，按 1:1 计，请核对）"
                elif (cand["unit"] or "") != (unit or ""):
                    tip = "（别名映射，已按倍数换算）"
                else:
                    tip = "（别名映射）"
                return price * factor, f"{reg}·{per}·{cand['unit'] or ''}{tip}"
            rows = self.con.execute("""SELECT * FROM material_price
                                           WHERE region=? AND period=? AND (name=? OR name LIKE ? OR ? LIKE '%'||name||'%')""",
                                        (reg, per, name, f"%{name}%", name)).fetchall()
            if not rows:
                # 写法差异（2.5mm2 / 2.5mm²、希腊 Φ / 西里尔 Ф）：归一化后唯一命中即可用；
                # 信息价常把单位写进名称（铈钨棒g、自粘橡胶带20mm×5m卷），所以名称+单位也试一次。
                # 归一化全等本身就是「同名」证据，不再走后面的同单位/同名门。
                idx = self._material_index(reg, per)
                for k in (self.norm_key(name), self.norm_key((name or "") + (unit or ""))):
                    cands = idx.get(k) or []
                    if len(cands) != 1:
                        continue
                    hit = cands[0]
                    price = hit["price_notax"] if hit["price_notax"] is not None else hit["price_tax"]
                    if price is None:
                        continue
                    factor = self._unit_factor(hit["unit"], unit)
                    tip = ""
                    if factor is None:
                        factor, tip = 1.0, f"（单位 {hit['unit'] or '—'}→{unit} 无法换算，按 1:1 计，请核对）"
                    elif (hit["unit"] or "") != (unit or ""):
                        tip = "（已按倍数换算）"
                    return price * factor, f"{reg}·{per}·{hit['unit'] or ''}{tip}"
            if not rows:
                continue
            # 模糊命中可能配错（「其他材料」曾命中「张拉锚具及其他材料」）：
            # 只用「同单位」的候选；单位不同时只有名称完全相等才允许换算，否则按缺价处理。
            same = [r for r in rows if (r["unit"] or "") == (unit or "")]
            cand = None
            if same:
                cand = same[0]
            else:
                exact = [r for r in rows if r["name"] == name]
                cand = exact[0] if exact else None
            if cand is None:
                # 走到这里说明名称/单位都只是「像」而不是「是」（比如价格名把单位粘进去了：
                # 铈钨棒g、自粘橡胶带20mm×5m卷）。归一化全等才当同名用。
                idx = self._material_index(reg, per)
                for k in (self.norm_key(name), self.norm_key((name or "") + (unit or ""))):
                    hit_list = idx.get(k) or []
                    if len(hit_list) != 1:
                        continue
                    hit = hit_list[0]
                    p2 = hit["price_notax"] if hit["price_notax"] is not None else hit["price_tax"]
                    if p2 is None:
                        continue
                    f2 = self._unit_factor(hit["unit"], unit)
                    tip = ""
                    if f2 is None:
                        f2, tip = 1.0, f"（单位 {hit['unit'] or '—'}→{unit} 无法换算，按 1:1 计，请核对）"
                    elif (hit["unit"] or "") != (unit or ""):
                        tip = "（已按倍数换算）"
                    return p2 * f2, f"{reg}·{per}·{hit['unit'] or ''}{tip}"
                continue
            price = cand["price_notax"] if cand["price_notax"] is not None else cand["price_tax"]
            if price is None:
                continue
            factor = self._unit_factor(cand["unit"], unit)
            tip = ""
            if factor is None:
                # 名称完全相同但单位不可换算（水 t↔m3、电 kW·h↔度）：定额惯例按 1:1 计，
                # 但必须在「价格出处」里标出来，别让人以为换算过了。
                if cand["name"] != name:
                    continue
                factor, tip = 1.0, f"（单位 {cand['unit']}→{unit} 未换算，按 1:1 计，请核对）"
            return price * factor, f"{reg}·{per}·{cand['unit'] or ''}{tip}"
        return None, None

    # 机械名称俗称：附件6 台班价表写「交流弧焊机」，定额子目写「交流电焊机」，同一种机械
    MACHINE_SYNONYM = (("电焊机", "弧焊机"), ("弧焊机", "电焊机"))

    def _machine_name_variants(self, name):
        out = [name]
        for a, b in self.MACHINE_SYNONYM:
            if a in name:
                out.append(name.replace(a, b))
        return out

    def _machine_price(self, name, unit, region, period):
        for reg, per in ((region, period), ("河北", "2023基期")):
            # 1) 别名（人工确认过的映射；scope='machine' 指向 machine_price.id）
            r = self.con.execute(
                """SELECT p.* FROM alias a JOIN machine_price p ON p.id=a.ref_id
                   WHERE a.scope='machine' AND a.term=? AND p.region=? AND p.period=?""",
                (name, reg, per)).fetchone()
            if r and r["price"] is not None:
                return r["price"], self._machine_src(r, name, name)
            for cand_name in self._machine_name_variants(name):
                r = self.con.execute(
                    "SELECT * FROM machine_price WHERE name=? AND region=? AND period=?",
                    (cand_name, reg, per)).fetchone()
                if not r:
                    # 定额里机械名带规格（「钢筋切断机直径(mm)40」），台班价表是「钢筋切断机[直径(mm)|40]」
                    cjk = "".join(ch for ch in cand_name if "\u4e00" <= ch <= "\u9fa5")
                    prefix = cjk[:5]
                    if prefix:
                        cands = self.con.execute(
                            """SELECT * FROM machine_price WHERE region=? AND period=? AND name LIKE ?
                               ORDER BY length(name)""", (reg, per, prefix + "%")).fetchall()
                        if cands:
                            m = re.search(r"(\d+(?:\.\d+)?)\D*$", cand_name)
                            hit = [c for c in cands if m and (f"|{m.group(1)}]" in c["name"])]
                            r = hit[0] if hit else cands[0]
                if r and r["price"] is not None:
                    return r["price"], self._machine_src(r, name, cand_name)
            r = self._machine_match(reg, per, name)
            if r is not None:
                return r[0], self._machine_src(r[1], name, name)
        return None, None

    @staticmethod
    def _split_machine_name(n):
        """「载重汽车[装载质量(t)|8]」→ ('载重汽车', '8')；无规格段就原样返回。"""
        m = re.match(r"^(.+?)\s*\[([^\]]*)\|\s*([^\]]*)\]?\s*$", n or "")
        if m:
            return m.group(1).strip(), m.group(3).strip()
        return (n or "").strip(), ""

    def _machine_match(self, reg, per, name):
        """按「机械名 + 规格值」匹配台班价表。

        定额侧写「载重汽车装载质量8t」「台式钻床钻孔直径16mm」，附件6 写
        「载重汽车[装载质量(t)|8]」——纯前缀 LIKE 永远差一个「[」，会整类漏价。
        这里把表按 (基准名, 规格值) 拆开缓存，用「基准名是定额名的中文前缀 + 规格数字相等」配对。
        """
        key = (reg, per)
        if key not in self._mp_cache:
            rows = []
            for r in self.con.execute(
                    "SELECT * FROM machine_price WHERE region=? AND period=?", (reg, per)):
                base, val = self._split_machine_name(r["name"])
                rows.append((r, "".join(ch for ch in base if "\u4e00" <= ch <= "\u9fa5"), val))
            self._mp_cache[key] = rows
        qcjk = "".join(ch for ch in name if "\u4e00" <= ch <= "\u9fa5")
        q_nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", name)]
        if not qcjk or not q_nums:
            return None

        def nums(s):
            return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", s or "")]

        for r, base_cjk, val in self._mp_cache[key]:
            if r["price"] is None or not base_cjk or not val:
                continue
            if len(base_cjk) < 2 or not qcjk.startswith(base_cjk):
                continue
            p_nums = nums(val)
            # 首个规格数字必须相等，且价表的规格数字都要在定额名里出现
            # （「工件直径630×工件长度2000」对「630*2000」；「排气量6m3/min」里的 3 不算规格）
            if p_nums and p_nums[0] == q_nums[0] and set(p_nums) <= set(q_nums):
                return r["price"], r
        return None

    @staticmethod
    def _machine_src(r, name, cand_name):
        """价格出处；只有「换了名字/换了机械」才标注替代，纯规格括号差异不标。"""
        src = f"{r['region']}·{r['period']}"
        cjk3 = "".join(ch for ch in name if "\u4e00" <= ch <= "\u9fa5")[:3]
        base = re.sub(r"\s*\[.*$", "", r["name"] or "")
        if cand_name != name or not base.startswith(cjk3):
            src += f"（替代：{r['name']}）"
        return src

    def fees(self, profession=None):
        """费用标准一览（专业 / 工程类别 / 税制 / 管理费率 / 利润率 / 税率）。"""
        sql = "SELECT * FROM fee_rate"
        args = []
        if profession:
            sql += " WHERE profession=?"
            args.append(profession)
        return [dict(r) for r in self.con.execute(sql + " ORDER BY profession,category,tax_mode", args)]

    def _fee(self, profession, category, tax_mode):
        """取费率行；税制不匹配时回落到同类别，找不到则给出可用清单。"""
        row = self.con.execute(
            """SELECT * FROM fee_rate WHERE profession=? AND category=? AND tax_mode=?
               ORDER BY effective_from DESC LIMIT 1""", (profession, category, tax_mode)).fetchone()
        if row:
            return row, None
        row = self.con.execute(
            """SELECT * FROM fee_rate WHERE profession=? AND category=?
               ORDER BY effective_from DESC LIMIT 1""", (profession, category)).fetchone()
        if row:
            return row, f"未找到「{tax_mode}」口径，已回落到「{row['tax_mode']}」"
        avail = [f"{r['category']}（{r['tax_mode']}）" for r in
                 self.con.execute("SELECT category,tax_mode FROM fee_rate WHERE profession=? "
                                  "ORDER BY category,tax_mode", (profession,))]
        if not avail:
            return None, (f"「{profession}」没有费用标准：05 附件3 只给了 建筑/装饰/安装/市政 四个专业，"
                          f"该册需先定取费口径才能组价（见 TODO P0-1）")
        return None, f"「{profession}」可用类别：" + "、".join(avail)

    def safety_rows(self, profession):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM safety_fee WHERE profession=? ORDER BY category", (profession,))]

    @staticmethod
    def _pick_by_threshold(rows, value):
        """按档位名里的阈值挑档：「建筑面积10000m2以上/以下」「工程造价5000万元以上/以下」。

        阈值从档位名里取（带「万」乘 10000），不写死数字，换省市/换算口径不用改代码。
        """
        for r in rows:
            m = re.search(r"(\d+(?:\.\d+)?)\s*(万)?", r["category"] or "")
            if not m:
                continue
            th = float(m.group(1)) * (10000.0 if m.group(2) else 1.0)
            if ("以上" in r["category"]) == (value > th):
                return r
        return None

    @staticmethod
    def _safety_caveat(row):
        """档位出处带「待确认」的（园林是参照建筑室外工程推出来的）在结果里明说，别当原文。"""
        if row and "待确认" in (row.get("source") or ""):
            return (f"安全文明费率「{row['category']}」{row['rate_pct']}% 是参照值："
                    f"05 附件2 未列该专业档位，正式出报告前须确认")
        return None

    def _safety(self, profession, category=None, build_area=None, cost=None):
        """安全文明施工费费率（05 附件2）。返回 (类别, 费率%, 提示)。

        category   ：直接指定档位（最可靠，正式出报告建议用它）
        build_area ：建筑面积 —— 建筑/安装按 10000m2 分档，给了它就不用猜
        cost       ：工程造价 —— 市政按 5000 万元分档，由 total 试算回代
        """
        rows = self.safety_rows(profession)
        if not rows:
            return None, None, f"「{profession}」没有安全文明施工费费率"
        if category:
            hit = [r for r in rows if r["category"] == category]
            if hit:
                return hit[0]["category"], hit[0]["rate_pct"], self._safety_caveat(hit[0])
            return None, None, f"「{profession}」可用安全文明类别：" + "、".join(r["category"] for r in rows)
        if len(rows) == 1:
            return rows[0]["category"], rows[0]["rate_pct"], self._safety_caveat(rows[0])
        value = build_area if build_area is not None else cost
        if value is not None:
            r = self._pick_by_threshold(rows, float(value))
            if r:
                why = (f"按建筑面积 {float(build_area):,.0f} m2" if build_area is not None
                       else f"按工程造价 {float(cost):,.2f} 元")
                return r["category"], r["rate_pct"], f"{why} 自动定档「{r['category']}」"
        r = rows[0]
        return r["category"], r["rate_pct"], (
            f"未指定档位，暂按「{r['category']}」{r['rate_pct']}%，★ 正式出报告前须按工程规模确认"
            f"（用 --safety-cat 指定，或 --build-area 让程序定档）；可选："
            + "、".join(x["category"] for x in rows))

    # ---------- 组价 ----------
    def price(self, code, book=None, region="雄安新区", period="2026-08",
              profession="建筑", fee_category="一般-二类", tax_mode="一般计税"):
        it = self.show(code, book)
        if not it:
            return {"错误": f"未找到子目 {code}"}
        fee, note = self._fee(profession, fee_category, tax_mode)
        if not fee:
            return {"错误": f"未找到费率 {profession}/{fee_category}；{note}"}
        labor_rate = {r["grade"]: r["price"] for r in
                      self.con.execute("SELECT grade,price FROM labor_rate WHERE region='河北'")}

        labor = material = machine = 0.0
        detail, missing, unpriced, pct_items = [], [], [], []
        for kind, items in it["消耗量"].items():
            for c in items:
                name, qty, unit = c["名称"], c["消耗量"] or 0.0, c["单位"]
                if c["未计价"]:
                    unpriced.append({"名称": name, "消耗量": qty, "单位": unit})
                    continue
                # 「其他材料 / 其他机械 3%」是按小计的百分比，不是「单价×数量」。
                # 当普通资源找价必然缺价，等于整份清单少算材料费（全库 46% 的子目都带这一项）。
                if (unit or "").strip() == "%" and name.startswith(("其他材料", "其他机械")):
                    pct_items.append((kind, name, qty))
                    continue
                price, src = None, None
                if kind == "人工":
                    grade = GRADE_OF.get(name)
                    price = labor_rate.get(grade) if grade else None
                    src = f"河北·2023基期·{grade}" if price else None
                elif kind == "材料":
                    price, src = self._material_price(name, unit, region, period)
                else:
                    price, src = self._machine_price(name, unit, region, period)
                amount = round(qty * price, 2) if price is not None else None
                if price is None:
                    missing.append({"类别": kind, "名称": name, "单位": unit, "消耗量": qty})
                if kind == "人工" and price is not None:
                    labor += amount
                elif kind == "机械" and price is not None:
                    machine += amount
                elif kind == "材料" and price is not None:
                    material += amount
                detail.append({"类别": kind, "名称": name, "单位": unit, "消耗量": qty,
                               "单价": price, "金额": amount, "价格出处": src})

        # 百分比项按小计计取：其他材料 = 材料费小计 × %，其他机械 = 机械费小计 × %
        # （不把自身算进基数，否则循环；多条同类百分比取合计）
        for kind, base_val in (("材料", "material"), ("机械", "machine")):
            sub = sum(q for k, n, q in pct_items if k == kind)
            if not sub:
                continue
            pot = material if base_val == "material" else machine
            amt = round(pot * sub / 100.0, 2)
            if base_val == "material":
                material += amt
            else:
                machine += amt
            detail.append({"类别": kind, "名称": f"其他{kind}（按{'材料' if kind == '材料' else '机械'}费 {sub:g}%）",
                           "单位": "%", "消耗量": sub, "单价": round(pot, 2), "金额": amt,
                           "价格出处": f"本子目{'材料' if kind == '材料' else '机械'}费小计 × {sub:g}%"})

        base = labor + machine
        overhead = base * fee["overhead_pct"] / 100.0
        profit = base * fee["profit_pct"] / 100.0
        unit_price = labor + material + machine + overhead + profit
        out = {
            "子目": {"册": it["book_id"], "编号": it["code"], "名称": it["name"], "单位": it["unit"],
                     "原页": it["page_no"], "解析置信度": it["confidence"], "复核状态": it["review_status"]},
            "取价": {"地区": region, "期次": period, "口径": "除税"},
            "明细": detail,
            "未计价主材(需另计)": unpriced,
            "缺价项": missing,
            "汇总": {"人工费": round(labor, 2), "材料费": round(material, 2), "机械费": round(machine, 2),
                     "直接费": round(labor + material + machine, 2),
                     "管理费": round(overhead, 2), "利润": round(profit, 2),
                     "综合单价(除税)": round(unit_price, 2),
                     "费率": f"{profession}/{fee_category}/{fee['tax_mode']} "
                             f"管理费{fee['overhead_pct']}% 利润{fee['profit_pct']}% 税率{fee['tax_pct']}%"},
        }
        if note:
            out["提示"] = note
        if it["review_status"] != "ok":
            warn = (f"该子目复核状态为「{it['review_status']}」，"
                    f"库内消耗量可能与原页不一致（原页 p{it['page_no']}），出报告前请核对")
            out["提示"] = f"{out['提示']}；{warn}" if out.get("提示") else warn
        if it["valid_to"]:
            warn = (f"⚠ 该子目已废止：{it['valid_to']} 起停止执行（见 change_log 同编号记录），"
                    f"不得用于新建项目，仅限存量工程核对")
            out["提示"] = f"{out['提示']}；{warn}" if out.get("提示") else warn
        elif it["valid_from"]:
            warn = f"现行子目，{it['valid_from']} 起执行（调整/补充件）"
            out["提示"] = f"{out['提示']}；{warn}" if out.get("提示") else warn
        return out

    # ---------- 总价（05 p44 计价程序） ----------
    def total(self, code, book=None, region="雄安新区", period="2026-08",
              profession="建筑", fee_category="一般-二类", tax_mode="一般计税",
              qty=1.0, adjust=0.0, other=0.0, safety_cat=None, build_area=None):
        """按 05 附件3 计价程序算到含税造价。

        1 直接费 + 2 企业管理费 + 3 利润 = 分部分项
        5 安全文明施工费 = (1+2+3+4 价款调整) × 费率    ← 基数见 safety_fee.base
        6 税金 = (1+2+3+4+5+其他项目) × 税率
        工程造价 = 1+2+3+4+5+6

        安全文明费档位：给了 --safety-cat 就用它；给了 --build-area（建筑/安装）按其定档；
        市政按「工程造价」分档，阈值只能由造价本身判定 → 先按一档试算再回代。
        """
        p = self.price(code, book, region, period, profession, fee_category, tax_mode)
        if "错误" in p:
            return p
        s = p["汇总"]
        cat, rate, s_note = self._safety(profession, safety_cat, build_area)
        if rate is None:
            return {"错误": s_note}

        qty = float(qty)
        adjust = float(adjust)
        other = float(other)
        tax_pct = self._fee(profession, fee_category, tax_mode)[0]["tax_pct"] or 0.0

        def calc(rate_pct):
            fbfx_ = s["综合单价(除税)"] * qty        # 1+2+3
            base_ = fbfx_ + adjust                   # 安全文明基数 = 1+2+3+4
            safety_ = base_ * rate_pct / 100.0
            before_ = base_ + safety_ + other
            return fbfx_, safety_, before_, before_ * tax_pct / 100.0

        fbfx, safety, before_tax, tax = calc(rate)
        # 市政这类「按工程造价分档」的：拿试算出的造价回代一次，避免 5000 万以上的工程
        # 落到「以下」档（安全文明费少算，属实质性错误）
        if safety_cat is None and build_area is None:
            rows = self.safety_rows(profession)
            if len(rows) > 1:
                alt = self._pick_by_threshold(rows, before_tax + tax)
                if alt:
                    if alt["rate_pct"] != rate:
                        cat, rate = alt["category"], alt["rate_pct"]
                        fbfx, safety, before_tax, tax = calc(rate)
                    s_note = f"按工程造价 {before_tax + tax:,.2f} 元自动定档「{cat}」{rate}%"

        out = {
            "子目": p["子目"],
            "费率": s["费率"],
            "工程量": {"数量": qty, "单位": p["子目"]["单位"]},
            "汇总": {
                "分部分项(1+2+3)": round(fbfx, 2),
                "价款调整(4)": round(adjust, 2),
                "安全文明施工费(5)": round(safety, 2),
                "其他项目": round(other, 2),
                "税前造价": round(before_tax, 2),
                "税金(6)": round(tax, 2),
                "含税工程造价": round(before_tax + tax, 2),
            },
            "依据": {"安全文明费率": f"{profession}/{cat} {rate}%",
                    "安全文明基数": "分部分项+价款调整（05 附件2）",
                    "税率": f"{tax_pct}%（{tax_mode}）"},
            "缺价项数量": len(p["缺价项"]),
            "未计价主材数量": len(p["未计价主材(需另计)"]),
        }
        notes = [n for n in (p.get("提示"), s_note) if n]
        if notes:
            out["提示"] = "；".join(notes)
        return out

    # ---------- 缺价缺口（给 alias 人工确认当靶子） ----------
    def gaps(self, book=None, region="雄安新区", period="2026-08", top=30, kind=None,
             with_candidates=False):
        """列出组价时「匹配不到价」的材料/机械，按使用频次排序。

        with_candidates=True 时给每条缺口补 3 个价格行候选（含 ref_id），
        人工确认后 `alias-add` 一条就能把这类缺价消掉——别名是解决缺价的正路。
        """
        sql = """SELECT c.kind, c.name, c.unit, count(DISTINCT c.item_id) AS items,
                        count(*) AS rows
                 FROM consumption c JOIN item i ON i.item_id=c.item_id
                 WHERE c.kind IN ('material','machine') AND c.is_unpriced=0"""
        args = []
        if book:
            sql += " AND i.book_id=?"
            args.append(book)
        if kind:
            sql += " AND c.kind=?"
            args.append(kind)
        sql += " GROUP BY c.kind, c.name, c.unit"
        rows = self.con.execute(sql, args).fetchall()

        out = []
        for r in rows:
            # 「其他材料/其他机械 %」是按小计的百分比，由 price() 单独计取，不算缺价
            if (r["unit"] or "").strip() == "%" and r["name"].startswith(("其他材料", "其他机械")):
                continue
            if r["kind"] == "material":
                price, src = self._material_price(r["name"], r["unit"], region, period)
            else:
                price, src = self._machine_price(r["name"], r["unit"], region, period)
            if price is not None:
                continue
            item = {"类别": "材料" if r["kind"] == "material" else "机械",
                    "名称": r["name"], "单位": r["unit"],
                    "涉及子目": r["items"], "消耗量行": r["rows"]}
            if with_candidates:
                item["候选"] = self.alias_candidates(r["kind"], r["name"], r["unit"], region, period)
            out.append(item)
        out.sort(key=lambda x: (-x["涉及子目"], -x["消耗量行"]))
        return {"取价": f"{region}·{period}", "缺口条数": len(out),
                "明细": out if top in (0, None) else out[:top]}

    def alias_add(self, term, ref_id, note="", kind="material"):
        """写一条资源别名（term=定额里的叫法，ref_id=价格行 id）。人工确认后调用。

        kind 决定 ref_id 指向哪张价格表：material → material_price，machine → machine_price。
        两张表的 id 各自独立，**搞混会把机械配到同号的材料行上**（踩过：干混砂浆储料罐 → 碳钢法兰）。
        """
        table = "machine_price" if kind == "machine" else "material_price"
        row = self.con.execute(f"SELECT id,name,spec,unit FROM {table} WHERE id=?",
                               (ref_id,)).fetchone()
        if not row:
            return {"错误": f"{table} 里没有 id={ref_id}"}
        old = self.con.execute("SELECT id FROM alias WHERE scope=? AND term=?",
                               (kind, term)).fetchone()
        note = note or f"{row['name']} {row['spec'] or ''}".strip()
        if old:
            self.con.execute("UPDATE alias SET ref_id=?, note=? WHERE id=?", (ref_id, note, old["id"]))
        else:
            self.con.execute("INSERT INTO alias(scope,ref_id,term,kind,note) "
                             "VALUES (?,?,?,'规格映射',?)", (kind, ref_id, term, note))
        self.con.commit()
        return {"已写入": term, "ref_id": ref_id, "表": table, "对应价格行": dict(row)}

    def _near_names(self, table, name, region, period, n=3):
        """库内近似名称：按中文前缀（3 个字）粗筛，够人工判断就行。"""
        cjk = "".join(ch for ch in name if "\u4e00" <= ch <= "\u9fa5")
        hits = []
        for k in range(min(4, len(cjk)), 1, -1):
            hits = [r[0] for r in self.con.execute(
                f"""SELECT DISTINCT name FROM {table}
                    WHERE region=? AND period=? AND name LIKE ? LIMIT ?""",
                (region, period, f"%{cjk[:k]}%", n))]
            if hits:
                break
        return hits

    # ---------- 别名候选（gaps 的建议表用） ----------
    def _price_rows(self, table, region, period):
        """整张价表缓存：缺价候选要全表扫，逐条查库会慢到不可用。"""
        key = (table, region, period)
        if key not in self._mp_cache:
            cols = ("price_notax,price_tax" if table == "material_price" else "price AS price_notax,NULL AS price_tax")
            rows = self.con.execute(
                f"SELECT id,name,spec,unit,{cols} FROM {table} "
                f"WHERE region=? AND period=?", (region, period)).fetchall()
            self._mp_cache[key] = [dict(r) for r in rows]
        return self._mp_cache[key]

    @staticmethod
    def _cjk(s):
        return "".join(ch for ch in (s or "") if "\u4e00" <= ch <= "\u9fa5")

    @staticmethod
    def _nums(s):
        return {x.rstrip("0").rstrip(".") for x in re.findall(r"\d+(?:\.\d+)?", s or "")}

    def alias_candidates(self, kind, name, unit, region, period, n=3):
        """给一个组价缺价的资源名，挑几条最像的价格行。

        门槛（试过更松的规则，全是噪音：单纯按数字/单位匹配会把「圆钢φ5.5～9」配到「LED筒灯 9W」，
        按前 2 字配会把「镀锌精制带帽螺栓」配到「镀锌金属穿线管」）：
          - 名称完全相同 → 6 分；
          - 中文前缀相同：长名（>4 字）要求 ≥4 字相同（5 分），短名（如「圆钢」「角钢」）放宽到 2 字（3 分，
            且不再给规格加分，避免乱配）；
          - 中文前缀都不满足就直接不进候选。
        再叠加：规格数字命中（仅高门槛的候选，每个 +2、最多 +4）、单位一致 +2 / 可换算 +1。
        本地期次没有就顺带扫河北 2023 基期（组价本身也是这么回落的）。
        """
        table = "material_price" if kind == "material" else "machine_price"
        qcjk, qnums = self._cjk(name), self._nums(name)
        need = 2 if len(qcjk) <= 4 else 4
        out = []
        for reg, per in ((region, period), ("河北", "2023基期")):
            for r in self._price_rows(table, reg, per):
                pcjk = self._cjk(r["name"])
                common = 0
                for a, b in zip(qcjk, pcjk):
                    if a != b:
                        break
                    common += 1
                why = []
                if r["name"] == name:
                    score = 6; why.append("同名")
                elif common >= need:
                    score = 5 if common >= 4 else 3
                    why.append(f"前{common}字")
                else:
                    continue
                if score >= 5:
                    hit = qnums & self._nums(r["spec"] or "")
                    if hit:
                        score += min(2 * len(hit), 4); why.append("规格" + "/".join(sorted(hit)))
                if (r["unit"] or "") == (unit or ""):
                    score += 2; why.append("单位同")
                elif self._unit_factor(r["unit"], unit) is not None:
                    score += 1; why.append("可换算")
                out.append({"ref_id": r["id"], "表": ("machine" if kind == "machine" else "material"),
                            "名称": r["name"], "规格": r["spec"],
                            "单位": r["unit"], "单价": r["price_notax"] or r["price_tax"],
                            "出处": f"{reg}·{per}", "score": score, "理由": "+".join(why)})
        out.sort(key=lambda x: (-x["score"], x["出处"] != f"{region}·{period}"))
        return out[:n]

    # ---------- 清单项目库（17 号附录B 河北补充项目 + 河北口径规则） ----------
    BILL_PROF = {"建筑": "房屋建筑与装饰工程", "装饰": "房屋建筑与装饰工程",
                 "房建": "房屋建筑与装饰工程", "房屋建筑与装饰工程": "房屋建筑与装饰工程",
                 "安装": "通用安装工程", "通用安装工程": "通用安装工程",
                 "市政": "市政工程", "园林": "园林绿化工程", "绿化": "园林绿化工程",
                 "仿古": "仿古建筑工程", "构筑物": "构筑物工程",
                 "轨交": "城市轨道交通工程", "城市轨道交通工程": "城市轨道交通工程"}

    def bill_items(self, keyword=None, profession=None, kind=None, code=None, limit=40):
        """查清单项目（bill_item）。keyword 对名称/特征/工作内容模糊匹配。"""
        sql = "SELECT * FROM bill_item WHERE 1=1"
        args = []
        if code:
            sql += " AND code=?"
            args.append(code)
        if profession:
            sql += " AND profession=?"
            args.append(self.BILL_PROF.get(profession, profession))
        if kind:
            sql += " AND item_kind=?"
            args.append(kind)
        if keyword:
            sql += " AND (name LIKE ? OR features LIKE ? OR work_content LIKE ?)"
            args += [f"%{keyword}%"] * 3
        return [dict(r) for r in self.con.execute(
            sql + " ORDER BY profession, code LIMIT ?", args + [limit])]

    def bill_rules(self, keyword=None, profession=None, kind=None, limit=40):
        """查河北清单计价规则（bill_rule）：总则口径 / 计算规则修正 / 编制规定。"""
        sql = "SELECT * FROM bill_rule WHERE 1=1"
        args = []
        if profession:
            sql += " AND (profession=? OR profession='全部专业')"
            args.append(self.BILL_PROF.get(profession, profession))
        if kind:
            sql += " AND kind=?"
            args.append(kind)
        if keyword:
            sql += " AND (content LIKE ? OR subject LIKE ? OR rule_no LIKE ?)"
            args += [f"%{keyword}%"] * 3
        return [dict(r) for r in self.con.execute(
            sql + " ORDER BY profession, page_no, id LIMIT ?", args + [limit])]

    # ---------- 组价建议（智能组价第一步） ----------
    # 清单行（名称/编码/特征）→ 候选定额子目 + 建议含量。
    # 打分只用可复核的确定性信号（名称字重合、章节、单位可换算），不猜配比、不猜规格——
    # 规格/强度/土类别这类「特征值」由人按项目特征在候选里挑（候选本身就是按规格分档的兄弟行）。
    SUGGEST_BOOKS = {
        "房屋建筑与装饰工程": ["建筑", "装饰"],
        "通用安装工程": ["安装", "安装补充"],
        "市政工程": ["市政"],
        "园林绿化工程": ["园林"],
        "仿古建筑工程": ["仿古建筑"],
        "构筑物工程": ["建筑", "市政"],
        "城市轨道交通工程": ["市政", "安装"],
    }
    # 工作内容 → 常见「另计」提示（只提示，不自动加子目；组合关系靠 bill_map 人工沉淀）
    WORK_HINTS = [
        ("装车", "工作内容含「装车」——挖/运常与装车分列，检查是否需另套装车子目"),
        ("场内运输", "工作内容含「场内运输」——可能需另套运土/运距子目（超出基本运距）"),
        ("模板", "工作内容含模板——模板通常另按模板项目列，勿重复计"),
        ("钢筋", "工作内容含钢筋——钢筋通常另按钢筋项目列，勿重复计"),
        ("运输", "工作内容含运输——注意定额基本运距，超距需另计"),
        ("五金", "工作内容含五金——小五金可能未含在子目价内，核对消耗量"),
    ]

    @staticmethod
    def _grams(s, ns=(2, 3)):
        """中文字符 n-gram 集（跳过数字/标点边界的 gram）。名称重合度的基本度量。"""
        t = re.sub(r"[^0-9A-Za-z一-龥]", "", s or "")
        out = set()
        for n in ns:
            for i in range(len(t) - n + 1):
                g = t[i:i + n]
                if re.search(r"[一-龥]", g):
                    out.add(g)
        return out

    def suggest(self, query, profession=None, books=None, limit=8):
        """组价建议：query 是清单编码（9位）或文字描述 → 候选子目 + 建议含量 + 理由。

        返回 {清单, 映射(已沉淀), 建议[], 提示[]}。含量 = 每 1 清单单位折多少个定额单位
        （清单 m³ / 定额 100m³ → 0.01），按 _unit_factor 换算，算不出就不给含量（只给候选）。
        """
        q = (query or "").strip()
        bill = None
        if re.fullmatch(r"\d{9}", q):
            bill = self.con.execute("SELECT * FROM bill_item WHERE code=? ORDER BY is_supplement DESC LIMIT 1",
                                    (q,)).fetchone()
            if bill:
                bill = dict(bill)
                name = bill["name"] or ""
                prof = bill.get("profession")
            else:
                name, prof = q, None
        else:
            name, prof = q, None
            hit = self.con.execute(
                "SELECT * FROM bill_item WHERE name=? ORDER BY is_supplement DESC LIMIT 1", (q,)).fetchone()
            if hit:                                    # 精确同名的清单项目 → 借它的专业/章节/工作内容
                bill = dict(hit)
                prof = bill.get("profession") or prof

        prof_full = self.BILL_PROF.get(profession or prof, profession or prof)
        if books:
            bl = [b.strip() for b in re.split(r"[,，/|]", books) if b.strip()]
        else:
            bl = self.SUGGEST_BOOKS.get(prof_full) or list(self.SUGGEST_BOOKS["房屋建筑与装饰工程"])

        # ---- 查询词：名称整词片段 + 3 字滑窗（「挖基坑土方」整词命不中「人工挖基坑…」，靠 3-gram 命中） ----
        # 通用词（混凝土/工程/现浇…）在海量名称里到处出现，检索要靠专名 3-gram（矩形柱），
        # 打分时通用 gram 降权——否则「聚苯板现浇混凝 土保温」会压过「矩形柱」。
        GENERIC_WORDS = ("混凝土", "现浇", "预制", "工程", "项目", "制作", "安装", "运输",
                         "一般", "其他", "材料", "基础", "地面", "墙面", "厚度", "强度",
                         "等级", "以内", "部分", "类型", "做法")

        def is_generic(g):
            """3-gram 是否通用：与通用词整段重叠，或共享 2 字窗口（「浇混凝」是「混凝土」碎片）。"""
            for wd in GENERIC_WORDS:
                if g in wd or wd in g:
                    return True
                if any(g[i:i + 2] in wd for i in range(len(g) - 1)):
                    return True
            return False

        grams = self._grams(name)
        segs = [s for s in re.findall(r"[一-龥]{2,}", name) if len(s) >= 2]
        segs.sort(key=len, reverse=True)
        segs = segs[:2]
        name_c = re.sub(r"[^一-龥]", "", name or "")
        grams3 = [name_c[i:i + 3] for i in range(len(name_c) - 2)] if len(name_c) >= 3 else []
        spec3 = [g for g in grams3 if not is_generic(g)]             # 专名 3-gram
        grams2 = [name_c[i:i + 2] for i in range(len(name_c) - 1)] if len(name_c) >= 2 else []
        spec2 = [g for g in grams2 if not is_generic(g)]
        # 检索词去重保序：整词 → 专名 3/2-gram → 其余（专名先检索，防通用词挤占 LIMIT）。
        # 2-gram 必须有：「栽植乔木」的 3-gram 与「栽植带土球落叶乔木」全不相交，只有 栽植/乔木 两块 2-gram 能搭上。
        terms = []
        for t in segs + spec3 + spec2 + grams3 + grams2:
            if t and t not in terms and len(t) >= 2:
                terms.append(t)
        terms = terms[:10]
        # 包含加分用的专名片段（整词片段 + 专名 3-gram）
        seg_hit_pool = list(dict.fromkeys(segs + spec3))[:6]
        # 特征名（"1.土类别" → 土类别）用于变体提示，不参与打分（特征值才决定套哪一档）
        feat_names = []
        for ln in ((bill or {}).get("features") or "").split("\n"):
            ln = re.sub(r"^\s*\d+[.．、]\s*", "", ln).strip()
            if 2 <= len(ln) <= 10 and re.fullmatch(r"[一-龥、，,（）()和/]{2,10}", ln):
                feat_names.append(ln)

        # ---- 候选检索：名称词 LIKE（限制册）并集，再全量打分 ----
        cand = {}
        for t in terms:
            for r in self.con.execute(
                    """SELECT i.item_id,i.book_id,i.code,i.name,i.unit,i.valid_to,i.review_status,
                              c.title AS chapter
                       FROM item i LEFT JOIN chapter c ON c.chapter_id=i.chapter_id
                       WHERE i.name LIKE ? AND i.book_id IN ("""
                    + ",".join("?" * len(bl)) + ") LIMIT 300",
                    [f"%{t}%"] + bl):
                cand[r["item_id"]] = dict(r)

        # 短名简称通道：定额里现浇柱就叫「柱」（5-8），名称 1-3 字且是查询的子串才收；
        # 单字名（土/柱）还要求章节与查询相关，防「土」∈「挖基坑土方」这类噪音
        if name_c:
            for r in self.con.execute(
                    """SELECT i.item_id,i.book_id,i.code,i.name,i.unit,i.valid_to,i.review_status,
                              c.title AS chapter
                       FROM item i LEFT JOIN chapter c ON c.chapter_id=i.chapter_id
                       WHERE length(i.name) BETWEEN 1 AND 3 AND i.name<>''
                         AND instr(?, i.name) > 0 AND i.book_id IN ("""
                    + ",".join("?" * len(bl)) + ") LIMIT 120", [name_c] + bl):
                d = dict(r)
                ch_ok = bool(grams & self._grams(d.get("chapter") or ""))
                if len(d["name"]) >= 2 or ch_ok:
                    d["_short"] = ch_ok
                    cand.setdefault(r["item_id"], d)

        # 材料反查通道：定额名常只写「钢管」「楼地面」，材质写在消耗量行里
        for tt in spec3[:3] or terms[:3]:
            for r in self.con.execute(
                    """SELECT DISTINCT i.item_id,i.book_id,i.code,i.name,i.unit,i.valid_to,
                                       i.review_status,c.title AS chapter
                       FROM consumption cc JOIN item i ON i.item_id=cc.item_id
                       LEFT JOIN chapter c ON c.chapter_id=i.chapter_id
                       WHERE cc.name LIKE ? AND i.book_id IN ("""
                    + ",".join("?" * len(bl)) + ") LIMIT 150", [f"%{tt}%"] + bl):
                d = dict(r)
                d["_mat"] = tt
                cand.setdefault(r["item_id"], d)

        bill_unit = (bill or {}).get("unit") or ""
        bill_ch = ((bill or {}).get("chapter") or "")
        out = []
        for it in cand.values():
            why = []
            score = 0.0
            ng = self._grams(it["name"])
            inter = grams & ng
            if inter:
                # 通用 gram 降权（混凝土/现浇到处都有），专名 gram 全权重
                w = sum(0.35 if is_generic(g) else 1.0 for g in inter)
                score += 3.0 * w / max(1.0, sum(0.35 if is_generic(g) else 1.0
                                               for g in grams))
                why.append(f"名称重合{len(inter)}个字块")
            seg_hit = [s for s in seg_hit_pool if s in it["name"]]
            if seg_hit:
                score += 2.0 * min(len(seg_hit), 3)
                why.append(f"含「{seg_hit[0]}」")
            elif it.get("_short"):
                score += 1.6
                why.append(f"简称命中（定额名「{it['name']}」是查询的简称）")
            elif it.get("_mat"):
                score += 1.4
                why.append(f"消耗材料含「{it['_mat']}」")
            ch = it.get("chapter") or ""
            if bill_ch and (bill_ch[:2] in ch or any(t in ch for t in re.findall(r"[一-龥]{2,}", bill_ch))):
                score += 1.2
                why.append("章节相近")
            elif ch:
                # 章节名与查询词重合（查询「现浇混凝土矩形柱」→ 混凝土章节，而非模板章节）
                ch_inter = grams & self._grams(ch)
                if ch_inter:
                    score += 0.8 * min(len(ch_inter), 3)
                    why.append(f"章节「{ch[:10]}」相关")
            coef = None
            if bill_unit and it.get("unit"):
                f = self._unit_factor(bill_unit, it["unit"])
                if f:
                    coef = round(1.0 / f, 6)
                    score += 1.0
                    why.append(f"单位可换算（含量≈{coef:g}）")
                elif self.norm_key(bill_unit) == self.norm_key(it["unit"]):
                    coef = 1.0
                    score += 1.0
                    why.append("单位相同")
            if it.get("valid_to"):
                score -= 5.0
                why.append(f"⚠已废止({it['valid_to']})")
            if it.get("review_status") != "ok":
                score -= 0.5
                why.append(f"复核状态{it['review_status']}")
            if score <= 0:
                continue
            out.append({"册": it["book_id"], "编号": it["code"], "名称": it["name"],
                        "单位": it["unit"] or "", "章节": ch, "含量": coef,
                        "评分": round(score, 2), "理由": "；".join(why),
                        "废止": it["valid_to"]})
        out.sort(key=lambda x: (-x["评分"], x["册"], x["编号"]))

        # 跨册参考：限定册内没有强候选时，全库找同类（不混进主列表，只作参考）
        hints = []
        if out and out[0]["评分"] < 1.5:
            glob = []
            for tt in (spec3[:2] + segs[:1]) or terms[:2]:
                for r in self.con.execute(
                        """SELECT DISTINCT i.book_id,i.code,i.name,i.unit
                           FROM item i WHERE i.name LIKE ? LIMIT 6""", [f"%{tt}%"]):
                    key = (r["book_id"], r["code"])
                    if key not in [(g["册"], g["编号"]) for g in glob] and r["book_id"] not in bl:
                        glob.append({"册": r["book_id"], "编号": r["code"], "名称": r["name"]})
            if glob:
                hints.append("限定册内候选较弱，跨册参考：" +
                             "、".join(f"{g['册']}{g['编号']}{g['名称'][:10]}" for g in glob[:4]) +
                             "（加 --books 可放开册过滤）")
        wc = (bill or {}).get("work_content") or ""
        for kw, h in self.WORK_HINTS:
            if kw in wc and h not in hints:
                hints.append(h)
        if feat_names:
            hints.append("候选多为同名不同档（土类别/深度/规格…），按项目特征选档：" + "、".join(feat_names[:5]))

        # 已沉淀的特征→子目组合（bill_map，人工确认的映射，优先展示）
        maps = []
        if bill:
            self._ensure_map_table()
            maps = []
            for r in self.con.execute(
                    "SELECT * FROM bill_map WHERE bill_code=? ORDER BY hits DESC", (bill["code"],)):
                d = dict(r)
                d["items"] = json.loads(d["items_json"])
                maps.append(d)

        return {"清单": bill, "查询": q, "检索册": bl,
                "映射": maps, "建议": out[:limit], "提示": hints}

    def _ensure_map_table(self):
        self.con.execute("""CREATE TABLE IF NOT EXISTS bill_map(
            id INTEGER PRIMARY KEY, bill_code TEXT NOT NULL, bill_name TEXT,
            feature_key TEXT DEFAULT '默认', items_json TEXT NOT NULL,
            hits INTEGER DEFAULT 0, note TEXT, created_at TEXT DEFAULT (datetime('now','localtime')),
            UNIQUE(bill_code, feature_key))""")

    def _map_lookup(self, bill_code, feature_key=None):
        """查 bill_map：先按特征键精确匹配，再回落『默认』。"""
        self._ensure_map_table()
        for key in ([str(feature_key)] if feature_key else []) + ["默认"]:
            r = self.con.execute("SELECT * FROM bill_map WHERE bill_code=? AND feature_key=?",
                                 (str(bill_code), key)).fetchone()
            if r:
                d = dict(r)
                d["items"] = json.loads(d["items_json"])
                return d
        return None

    def map_add(self, bill_code, items, note="", feature_key="默认"):
        """沉淀「清单项目 → 子目组合」映射（智能组价的核心资产，人工确认后写入）。
        items: [{"册","编号","含量","说明"}]；重跑同 key 覆盖。"""
        if not re.fullmatch(r"\d{9}", str(bill_code or "")):
            return {"错误": f"清单编码须为 9 位数字，收到 {bill_code!r}"}
        bill = self.con.execute("SELECT code,name FROM bill_item WHERE code=? LIMIT 1",
                                (bill_code,)).fetchone()
        for it in items:
            if not it.get("册") or not it.get("编号"):
                return {"错误": f"子目项缺 册/编号：{it}"}
            r = self.con.execute("SELECT name FROM item WHERE book_id=? AND code=?",
                                 (it["册"], it["编号"])).fetchone()
            if not r:
                return {"错误": f"库里没有 {it['册']} {it['编号']}（item 表）"}
            it.setdefault("名称", r["name"])
        self._ensure_map_table()
        self.con.execute("""INSERT INTO bill_map(bill_code,bill_name,feature_key,items_json,note)
                            VALUES (?,?,?,?,?)
                            ON CONFLICT(bill_code,feature_key) DO UPDATE SET
                              items_json=excluded.items_json, note=excluded.note""",
                         (bill_code, bill["name"] if bill else "", feature_key,
                          json.dumps(items, ensure_ascii=False), note))
        self.con.commit()
        return {"已写入": f"{bill_code} {bill['name'] if bill else ''}｜{feature_key}",
                "子目数": len(items),
                "组合": [{"册": i["册"], "编号": i["编号"], "含量": i.get("含量")} for i in items]}

    def map_list(self, bill_code=None, limit=40):
        self._ensure_map_table()
        sql = "SELECT * FROM bill_map"
        args = []
        if bill_code:
            sql += " WHERE bill_code=?"
            args.append(bill_code)
        out = []
        for r in self.con.execute(sql + " ORDER BY bill_code, feature_key LIMIT ?", args + [limit]):
            d = dict(r)
            d["items"] = json.loads(d["items_json"])
            out.append(d)
        return out

    # ---------- 措施项目费（05 附件1 其它措施项目） ----------
    # 计取口径（05 附件1 说明二）：措施项目费 =（人工费+机械费）× 费率%
    # 基数按「各专业工程」分列（说明二），范围映射见 _measure_scope。
    MEASURE_BOOKS = {"建筑": "建筑工程", "装饰": "装饰装修工程", "安装": "安装工程",
                     "市政": "市政工程", "安装补充": "安装工程"}
    MEASURE_PROFESSION = {"建筑": "建筑工程", "装饰": "装饰装修工程",
                          "安装": "安装工程", "市政": "市政工程"}
    CITY_SCOPE_DEFAULT = "土石方工程、拆除工程、护坡及挡土墙工程、大型机械一次安拆及场外运输费"
    MEASURE_COND_RE = re.compile(r"施工与生产同时|有害身体|不断交")
    MEASURE_SPECIAL_BASE_RE = re.compile(r"操作高度")   # 基数=高度增加部分的人工+机械，推不出来

    def measure_rows(self, code=None, profession=None, name=None, scope=None):
        """查 measure_fee。name 比对忽略空格（表内多行单元格会带出空格）。"""
        sql = "SELECT * FROM measure_fee WHERE 1=1"
        args = []
        if code:
            sql += " AND code=?"
            args.append(str(code).strip())
        if profession:
            sql += " AND profession=?"
            args.append(self.MEASURE_PROFESSION.get(profession, profession))
        if name:
            sql += " AND replace(name,' ','')=replace(?,' ','')"
            args.append(str(name).strip())
        if scope:
            sql += " AND scope=?"
            args.append(scope)
        return [dict(r) for r in self.con.execute(sql + " ORDER BY code", args)]

    def _chapter_path(self, chapter_id):
        titles, cid, seen = [], chapter_id, set()
        while cid and cid not in seen:
            seen.add(cid)
            r = self.con.execute("SELECT title,parent_id FROM chapter WHERE chapter_id=?",
                                 (cid,)).fetchone()
            if not r:
                break
            titles.append(r["title"] or "")
            cid = r["parent_id"]
        return " ".join(reversed(titles))

    def _measure_scope(self, book, chapter_id, page_no):
        """子目 → 05 附件1 的「专业工程」计费范围（与 measure_fee.scope 对齐）。

        安装第八章（刷油、防腐蚀、绝热）三类费率不同，按子目页码分段：
        p<899 除锈/刷油，899-950 防腐蚀，951-998 绝热，≥999 补口补伤（防腐类）。
        """
        prof = self.MEASURE_BOOKS.get(book)
        if not prof:
            return None
        path = self._chapter_path(chapter_id)
        if prof == "装饰装修工程":
            return "ALL"
        if prof == "建筑工程":
            return ("地基处理与边坡支护、桩基础工程"
                    if re.search(r"地基处理|边坡|桩基础", path) else "一般土建工程")
        if prof == "市政工程":
            for kw, sc in (("道路", "道路工程"), ("桥涵", "桥涵工程"),
                           ("管网", "管网工程"), ("路灯", "路灯工程")):
                if kw in path:
                    return sc
            return self.CITY_SCOPE_DEFAULT
        # 安装（含安装补充）
        if "机械设备" in path:
            return "机械设备安装工程"          # 05 附件1 没有对应费率行，只提示不计取
        for kw, sc in (("电气设备", "电气设备工程"), ("智能化", "智能化系统工程"),
                       ("消防", "消防设备工程"), ("给排水", "给排水、采暖、燃气工程"),
                       ("泵间", "泵间管道工程"), ("通风空调", "通风空调工程")):
            if kw in path:
                return sc
        if any(k in path for k in ("刷油", "防腐", "绝热")):
            if page_no is None:
                return "刷油工程|防腐工程|绝热工程"
            if page_no < 899:
                return "刷油工程"
            if page_no < 951:
                return "防腐工程"
            if page_no < 999:
                return "绝热工程"
            return "防腐工程"
        return "ALL"

    def _measure_auto_default(self, row):
        """自动计取的默认集：只含「一般都要计」的行。

        操作高度/超高/垂直运输/系统调试费与施工与生产、有害环境、不断交都是条件项，
        自动模式不计，列进提示由人按项目实际补（宁可不给数，不可给错数）。
        """
        if self.MEASURE_COND_RE.search(row["name"] or ""):
            return False
        code = row["code"] or ""
        if not code.startswith("C"):
            return True                      # A/B/D：其它措施项目，一般计取
        n = code[1:]
        if n.isdigit():
            i = int(n)
            return 1 <= i <= 9 or 132 <= i <= 139    # 脚手架搭拆费 / 安装通用其它措施
        return False

    def _bill_measures(self, entries, base_by):
        """措施项目计取。返回 (行, 合计, 提示)。

        条目形态：
          {"名称","金额"}                  手填
          {"编号":"A001"[,"系数","基数"]}  按 measure_fee 自动计取
          {"名称":"冬季施工增加费"[,...]}  按名称匹配（忽略空格），多义时报候选
          {"自动": true}                    各专业范围的默认全计（条件项除外）
        系数：冬雨季施工期不足规定天数 50% 时按 0.5（05 附件1 说明一.3）
        """
        rows, notes, warns = [], [], []
        used, skipped_cond = set(), []
        total = 0.0

        def base_of(row, override):
            if override is not None:
                return float(override), "指定基数"
            prof, scope = row["profession"], row["scope"] or "ALL"
            if scope == "ALL":
                return (sum(v for (p, s), v in base_by.items() if p == prof),
                        f"{prof}·全部人工+机械")
            if "|" in scope:
                subs = scope.split("|")
                return (sum(base_by.get((prof, s), 0.0) for s in subs), scope)
            return base_by.get((prof, scope), 0.0), scope

        def pick(row, coef, override):
            nonlocal total
            if self.MEASURE_SPECIAL_BASE_RE.search(row["name"] or "") and override is None:
                warns.append(f"{row['code']} {row['name']}：计费基数是「高度增加部分的人工费+"
                             f"机械费之和」（05 附件1 说明三.2），推不出来——"
                             f"请给『基数』或改填『金额』")
                return
            base, bnote = base_of(row, override)
            if override is None and base <= 0:
                warns.append(f"{row['code']} {row['name']}：范围「{row['scope'] or '全部'}」"
                             f"没有清单子目（基数 0），未计取")
                return
            rate = row["rate_pct"] or 0.0
            amt = base * rate / 100.0 * coef

            def part(p):
                return round(base * (p or 0.0) / 100.0 * coef, 2)
            total += amt
            used.add(row["code"])
            rows.append({
                "编号": row["code"], "名称": row["name"], "专业": row["profession"],
                "范围": row["scope"], "费率%": rate,
                "基数": round(base, 2), "基数来源": bnote, "系数": coef,
                "人工": part(row["labor_pct"]), "材料": part(row["material_pct"]),
                "机械": part(row["machine_pct"]), "金额": round(amt, 2),
                "计算式": f"({bnote}){base:,.2f} × {rate}%"
                          + ("" if coef == 1 else f" × 系数{coef}"),
                "来源": f"05 附件1 p{row['page_no']}", "方式": "自动计取",
            })

        entries = entries if isinstance(entries, list) else (
            [{"自动": True}] if entries == "自动" else [])
        for m in entries:
            if isinstance(m, str):
                m = {"名称": m}
            raw_coef = m.get("系数")
            coef = (1.0 if isinstance(raw_coef, dict)
                    else float(raw_coef if raw_coef is not None else 1.0))
            override = m.get("基数")
            if m.get("自动"):
                excl = {str(x) for x in (m.get("排除") or [])}
                coef_map = m.get("系数") if isinstance(m.get("系数"), dict) else {}
                profs = {p for (p, s) in base_by}
                pool = self.measure_rows()
                for row in pool:
                    if row["profession"] not in profs or row["code"] in used \
                            or row["code"] in excl:
                        continue
                    if not self._measure_auto_default(row):
                        if self.MEASURE_COND_RE.search(row["name"] or "") or \
                                (row["code"] or "").startswith("C"):
                            skipped_cond.append(f"{row['code']} {row['name']}")
                        continue
                    scope = row["scope"] or "ALL"
                    if scope != "ALL" and not any(
                            base_by.get((row["profession"], s), 0) > 0
                            for s in scope.split("|")):
                        continue
                    pick(row, float(coef_map.get(row["code"], coef)), override)
                continue
            row = None
            if m.get("编号"):
                pool = self.measure_rows(code=str(m["编号"]))
                if not pool:
                    warns.append(f"措施编号 {m['编号']} 不存在（quota.py measure 可查全部）")
                    continue
                row = pool[0]
            elif m.get("名称") and m.get("金额") is None:
                pool = self.measure_rows(name=str(m["名称"]),
                                         profession=m.get("专业") or None)
                if len(pool) == 1:
                    row = pool[0]
                elif not pool:
                    warns.append(f"措施名称「{m['名称']}」在 measure_fee 里没有对应行；"
                                 f"要手填请带上『金额』")
                    continue
                else:
                    warns.append(f"措施名称「{m['名称']}」有多条："
                                 + "、".join(f"{r['code']}({r['profession']}/{r['scope'] or '全部'})"
                                            for r in pool) + "，请改用『编号』指定")
                    continue
            if row is not None:
                if row["code"] in used:
                    warns.append(f"{row['code']} {row['name']} 已计取过，跳过"
                                 f"（要调整系数请用自动条目的『系数』映射）")
                    continue
                pick(row, coef, override)
            else:
                amt = float(m.get("金额") or 0)
                total += amt
                rows.append({"名称": m.get("名称"), "金额": round(amt, 2), "方式": "手填"})

        if any("冬季施工增加费" in (r.get("名称") or "") for r in rows):
            notes.append("冬（雨）季施工增加费：施工期不足规定天数 50% 按 50% 计取，"
                         "超过按全部计取（05 附件1 说明一.3）——不足时给条目加系数 0.5"
                         "（自动条目可写 \"系数\":{\"A001\":0.5}）")
        if skipped_cond:
            notes.append("条件项未自动计取，按项目实际用 {\"编号\":\"xxx\"} 选用：" +
                         "；".join(skipped_cond[:8]) +
                         ("…" if len(skipped_cond) > 8 else ""))
        mech = [v for (p, s), v in base_by.items() if s == "机械设备安装工程"]
        if mech:
            notes.append(f"机械设备安装工程的子目人工+机械 {mech[0]:,.2f} 元无对应措施费率行"
                         f"（05 附件1 C001-C009 不含机械），未纳入措施基数")
        return rows, round(total, 2), notes + warns

    # ---------- 清单级汇总（bill） ----------
    def bill(self, spec, region=None, period=None):
        """清单行批量组价：清单 → 子目 → 分部分项 → 措施/其他 → 安全文明 → 增值税。

        spec（JSON，见 方案.md「清单级组价」）：
          {"工程": {"名称","专业","类别","税制","地区","期次","建筑面积","安全文明档位"},
           "价款调整": 0,
           "措施项目": [{"名称","金额"}                       # 手填
                       ,{"编号":"A001","系数":1}             # 按 05 附件1 自动计取
                       ,{"名称":"冬季施工增加费"}            # 按名称匹配后计取
                       ,{"自动": true}],                     # 默认全计（条件项除外）
           "其他项目": [{"名称","金额","含税": false}],     # 含税=true 的（专业工程暂估价）不进税金基数
           "清单": [{"编号","名称","单位","工程量",
                     "子目": [{"册","编号","含量","专业","类别","措施类别"}]}]}

        含量 = 每 1 个清单单位消耗多少个该定额单位（不靠程序猜单位换算，宁可让录入的人说清楚）。
        措施项目自动计取口径（05 附件1 说明二）：（人工费+机械费）× 费率%，基数按「各专业工程」
        分列（子目按册/章节自动归范围，`措施类别` 可覆盖）。
        安全文明费基数 = 分部分项 + 价款调整（05 附件2）；增值税基数 = 分部分项+措施+其他(不含含税项)+价款调整+安全文明。
        """
        head = spec.get("工程") or {}
        region = region or head.get("地区") or "雄安新区"
        period = period or head.get("期次") or "2026-08"
        tax_mode = head.get("税制") or "一般计税"
        profession = head.get("专业") or "建筑"
        fee_cat = head.get("类别") or DEFAULT_FEE.get(profession, "一般-二类")
        build_area = head.get("建筑面积")

        fee, fee_note = self._fee(profession, fee_cat, tax_mode)
        if not fee:
            return {"错误": f"未找到费率 {profession}/{fee_cat}；{fee_note}"}
        tax_pct = fee["tax_pct"] or 0.0

        rows_out, errs, unit_warns = [], [], []
        map_used = []
        fbfx, miss_total, unp_total = 0.0, 0, 0
        base_by = {}          # {(专业全称, 措施计费范围): 人工+机械}，措施项目计取基数
        for ln in spec.get("清单") or []:
            lcode = str(ln.get("编号") or "").strip()
            qty = float(ln.get("工程量") or 0)
            subs, unit_price, miss, unp = [], 0.0, 0, 0
            subs_in = ln.get("子目") or []
            if not subs_in:                       # 没给子目 → 查 bill_map 沉淀映射（智能组价闭环）
                m = self._map_lookup(ln.get("清单编码") or lcode, ln.get("特征键"))
                if m:
                    subs_in = m["items"]
                    map_used.append(f"{ln.get('清单编码') or lcode}→{m['feature_key']}×{len(m['items'])}子目")
                    self.con.execute("UPDATE bill_map SET hits=hits+1 WHERE id=?", (m["id"],))
            for s in subs_in:
                book = s.get("册") or profession
                scode = str(s.get("编号") or "").strip()
                coef = float(s.get("含量", 1.0) if s.get("含量") is not None else 1.0)
                prof = s.get("专业") or BOOK_PROF.get(book) or (book if book in DEFAULT_FEE else profession)
                cat = s.get("类别") or DEFAULT_FEE.get(prof, fee_cat)
                pr = self.price(scode, book, region, period, prof, cat, s.get("税制") or tax_mode)
                if "错误" in pr:
                    errs.append({"清单编号": lcode, "子目": scode, "错误": pr["错误"]})
                    continue
                up = pr["汇总"]["综合单价(除税)"] * coef
                unit_price += up
                # 措施费基数：按 05 附件1「各专业工程的人工费+机械费之和」归范围
                mp = self.MEASURE_BOOKS.get(book)
                if mp:
                    it = self.con.execute(
                        "SELECT chapter_id,page_no FROM item WHERE book_id=? AND code=?",
                        (book, scode)).fetchone()
                    scope = (s.get("措施类别")
                             or self._measure_scope(book, it["chapter_id"] if it else None,
                                                    it["page_no"] if it else None))
                    lm = (pr["汇总"]["人工费"] + pr["汇总"]["机械费"]) * coef * qty
                    base_by[(mp, scope)] = base_by.get((mp, scope), 0.0) + lm
                miss += len(pr["缺价项"])
                unp += len(pr["未计价主材(需另计)"])
                su = pr["子目"]["单位"]
                if ln.get("单位") and su and coef == 1 and su != ln["单位"]:
                    unit_warns.append(f"{lcode} 清单单位「{ln['单位']}」≠ 子目 {scode} 单位「{su}」，请确认含量")
                subs.append({"册": book, "编号": scode, "名称": pr["子目"]["名称"], "单位": su,
                             "含量": coef, "综合单价(除税)": round(up, 2),
                             "缺价项": len(pr["缺价项"]),
                             "未计价主材": len(pr["未计价主材(需另计)"])})
            amount = unit_price * qty
            fbfx += amount
            miss_total += miss
            unp_total += unp
            rows_out.append({"编号": lcode, "名称": ln.get("名称"), "单位": ln.get("单位"),
                             "工程量": qty, "综合单价(除税)": round(unit_price, 2),
                             "合价": round(amount, 2), "子目": subs,
                             "缺价项": miss, "未计价主材": unp})

        measures, m_sum, m_notes = self._bill_measures(spec.get("措施项目"), base_by)
        others = spec.get("其他项目") or []
        o_sum = sum(float(o.get("金额") or 0) for o in others)
        o_taxed = sum(float(o.get("金额") or 0) for o in others if not o.get("含税"))
        adjust = float(spec.get("价款调整") or spec.get("工程", {}).get("价款调整") or 0)

        cost_hint = fbfx + m_sum + o_sum + adjust
        cat, rate, s_note = self._safety(profession, head.get("安全文明档位"), build_area,
                                        cost=(cost_hint if cost_hint else None))
        if rate is None:
            return {"错误": s_note}
        safety = (fbfx + adjust) * rate / 100.0
        before_tax = fbfx + m_sum + o_sum + adjust + safety
        tax = (fbfx + m_sum + o_taxed + adjust + safety) * tax_pct / 100.0

        out = {
            "工程": {"名称": head.get("名称"), "专业": profession, "类别": fee_cat, "税制": tax_mode,
                    "取价": f"{region}·{period}", "建筑面积": build_area},
            "清单": rows_out,
            "措施项目": measures,
            "其他项目": others,
            "汇总": {
                "分部分项": round(fbfx, 2),
                "措施项目": round(m_sum, 2),
                "其他项目": round(o_sum, 2),
                "其中含税项(不进税金基数)": round(o_sum - o_taxed, 2),
                "价款调整": round(adjust, 2),
                "安全文明施工费": round(safety, 2),
                "税前造价": round(before_tax, 2),
                "增值税": round(tax, 2),
                "含税工程造价": round(before_tax + tax, 2),
            },
            "依据": {"安全文明费率": f"{profession}/{cat} {rate}%",
                    "安全文明基数": "分部分项+价款调整（05 附件2）",
                    "增值税率": f"{tax_pct}%（{tax_mode}）",
                    "计费基数": fee["base_expr"],
                    "措施项目": "05 附件1：（人工费+机械费）×费率%，基数按各专业工程分列"},
            "措施基数": {f"{p}/{s}": round(v, 2) for (p, s), v in sorted(base_by.items())},
            "校验": {"缺价项合计": miss_total, "未计价主材合计": unp_total,
                    "错误": errs, "单位提示": unit_warns},
        }
        notes = [n for n in (fee_note, s_note) if n] + list(m_notes or [])
        if map_used:
            self.con.commit()                     # hits 累计
            notes.append("自动套用 bill_map 沉淀映射：" + "；".join(map_used))
        if notes:
            out["提示"] = "；".join(notes)
        if build_area:
            out["汇总"]["单方造价(元/m2)"] = round((before_tax + tax) / float(build_area), 2)
        return out

    # ---------- 导出（给 LLM / 其它程序） ----------
    def export_jsonl(self, out, book=None, only_ok=False):
        sql = "SELECT * FROM item"
        args, cond = [], []
        if book:
            cond.append("book_id=?")
            args.append(book)
        if only_ok:
            cond.append("review_status='ok'")
        if cond:
            sql += " WHERE " + " AND ".join(cond)
        n = 0
        with open(out, "w", encoding="utf-8") as f:
            for r in self.con.execute(sql + " ORDER BY book_id,code", args):
                f.write(json.dumps(self.show(r["code"], r["book_id"]), ensure_ascii=False) + "\n")
                n += 1
        return n


def main():
    ap = argparse.ArgumentParser(description="河北定额库客户端")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("stats")
    s = sub.add_parser("fees"); s.add_argument("--profession")
    s = sub.add_parser("measure")                     # 05 附件1 措施项目费率
    s.add_argument("--profession", help="建筑/装饰/安装/市政 或全称")
    s.add_argument("--code", help="措施编号，如 A001")
    s.add_argument("--scope", help="计费范围，如 电气设备工程 / 一般土建工程")
    s = sub.add_parser("bill-items")              # 清单项目库（17号附录B 河北补充项目）
    s.add_argument("keyword", nargs="?", help="名称/特征/工作内容 模糊检索")
    s.add_argument("--profession", help="建筑/安装/市政/园林/仿古… 或全称")
    s.add_argument("--kind", choices=["分部分项", "措施项目"])
    s.add_argument("--code", help="项目编码精确查询（打印五要素详情）")
    s.add_argument("--rules", action="store_true", help="查河北计价规则（bill_rule）")
    s.add_argument("--limit", type=int, default=40)
    s = sub.add_parser("suggest")                 # 组价建议：清单行 → 候选子目+含量
    s.add_argument("query", help="清单编码（9位）或文字描述，如 010102001 / 挖基坑土方")
    s.add_argument("--profession", help="建筑/安装/市政/园林… 或全称（缺省按清单项目推断）")
    s.add_argument("--books", help="限定册，逗号分隔，如 建筑,装饰（缺省按专业推断）")
    s.add_argument("--limit", type=int, default=8)
    s.add_argument("--json", action="store_true", help="输出 JSON")
    s = sub.add_parser("map-add")                 # 沉淀 清单→子目组合 映射（人工确认后）
    s.add_argument("bill_code", help="清单项目编码（9位）")
    s.add_argument("--items", required=True,
                   help='子目组合，如 "建筑:1-24:0.01,建筑:1-40:0.01"（册:编号:含量）')
    s.add_argument("--feature-key", default="默认", help="特征组合键，如 C30|矩形；默认『默认』")
    s.add_argument("--note")
    s = sub.add_parser("map-list")
    s.add_argument("bill_code", nargs="?")
    s.add_argument("--limit", type=int, default=40)
    s = sub.add_parser("search"); s.add_argument("keyword"); s.add_argument("--book"); s.add_argument("--limit", type=int, default=20)
    s = sub.add_parser("show"); s.add_argument("code"); s.add_argument("--book")
    s = sub.add_parser("price"); s.add_argument("code"); s.add_argument("--book")
    s.add_argument("--region", default="雄安新区"); s.add_argument("--period", default="2026-08")
    s.add_argument("--profession", default=None, help="默认按 --book 推断")
    s.add_argument("--fee", default=None, help="工程类别；缺省按专业取常用类别")
    s.add_argument("--tax", default="一般计税", choices=["一般计税", "简易计税"])
    s = sub.add_parser("total"); s.add_argument("code"); s.add_argument("--book")
    s.add_argument("--region", default="雄安新区"); s.add_argument("--period", default="2026-08")
    s.add_argument("--profession", default=None); s.add_argument("--fee", default=None)
    s.add_argument("--tax", default="一般计税", choices=["一般计税", "简易计税"])
    s.add_argument("--qty", type=float, default=1.0, help="清单工程量（按子目单位）")
    s.add_argument("--adjust", type=float, default=0.0, help="价款调整（元）")
    s.add_argument("--other", type=float, default=0.0, help="其他项目费（元）")
    s.add_argument("--safety-cat", default=None, help="安全文明施工费类别，见 safety_fee 表")
    s.add_argument("--build-area", type=float, default=None,
                   help="建筑面积(m2)：建筑/安装按 10000m2 自动定安全文明档；市政按造价回代定档")
    s = sub.add_parser("bill"); s.add_argument("spec", help="清单 JSON（结构见 方案.md）")
    s.add_argument("--out", help="把完整结果写成 JSON")
    s.add_argument("--region"); s.add_argument("--period")
    s = sub.add_parser("gaps"); s.add_argument("--book"); s.add_argument("--region", default="雄安新区")
    s.add_argument("--period", default="2026-08"); s.add_argument("--top", type=int, default=30)
    s.add_argument("--kind", choices=["material", "machine"])
    s.add_argument("--csv", help="把候选表写成 CSV（含 ref_id，人工勾选后 alias-add）")
    s = sub.add_parser("alias-add"); s.add_argument("term")
    s.add_argument("--ref", type=int, required=True,
                   help="价格行 id：材料=material_price.id，机械=machine_price.id（gaps --csv 里可查）")
    s.add_argument("--kind", default="material", choices=["material", "machine"],
                   help="机械类必须写 machine，否则会配到同号的材料行上")
    s.add_argument("--note")
    s = sub.add_parser("export"); s.add_argument("out"); s.add_argument("--book"); s.add_argument("--only-ok", action="store_true")
    a = ap.parse_args()

    q = QuotaDB()
    if a.cmd == "stats":
        print(json.dumps(q.stats(), ensure_ascii=False, indent=1))
    elif a.cmd == "fees":
        print(f"{'专业':<4}{'工程类别':<24}{'税制':<8}{'管理费%':>7}{'利润%':>7}{'税率%':>7}  基数/出处")
        for r in q.fees(a.profession):
            print(f"{r['profession']:<4}{r['category']:<24}{r['tax_mode']:<8}"
                  f"{r['overhead_pct']:>7.1f}{r['profit_pct']:>7.1f}{(r['tax_pct'] or 0):>7.1f}  "
                  f"{r['base_expr']} / {r['source']}")
    elif a.cmd == "measure":
        rows = q.measure_rows(code=a.code, profession=a.profession, scope=a.scope)
        if not rows:
            print("没有匹配的措施费率行（quota.py measure 可看全部）")
        print(f"{'编号':<6}{'专业':<8}{'范围':<28}{'名称':<32}{'费率%':>7}{'人%':>6}{'材%':>6}{'机%':>6}  出处")
        for r in rows:
            print(f"{r['code']:<6}{r['profession'][:4]:<8}{(r['scope'] or 'ALL')[:26]:<28}"
                  f"{(r['name'] or '')[:30]:<32}{(r['rate_pct'] or 0):>7.2f}"
                  f"{(r['labor_pct'] or 0):>6.2f}{(r['material_pct'] or 0):>6.2f}"
                  f"{(r['machine_pct'] or 0):>6.2f}  p{r['page_no']}")
        print("\n计取口径：措施项目费 =（人工费+机械费）× 费率%（05 附件1 说明二），"
              "基数按各专业工程分列；人/材/机% 是费率的组成拆分（三项之和=费率）。\n"
              "条件项（操作高度/超高/垂直运输/系统调试/施工与生产/有害环境/不断交）"
              "不进自动计取，按项目实际选用。")
    elif a.cmd == "bill-items":
        if a.rules:
            rows = q.bill_rules(a.keyword, a.profession, None, a.limit)
            if not rows:
                print("没有匹配的计价规则")
            for r in rows:
                print(f"[{r['profession']} {r['rule_no'] or ''}] {r['kind']} "
                      f"p{r['page_no']}"
                      + (f" 涉及{r['subject']}" if r["subject"] else ""))
                print(f"   {r['content'][:150]}")
        else:
            if a.code:
                rows = q.bill_items(code=a.code)
            else:
                rows = q.bill_items(a.keyword, a.profession, a.kind, limit=a.limit)
            if not rows:
                print("没有匹配的清单项目（bill-items --rules 可查计价规则）")
            for r in rows:
                ch = (r.get("chapter") or "")[:10]
                print(f"{r['code']}  {r['name'][:22]:24}{(r['unit'] or ''):5}"
                      f"{r['item_kind']:5}{r['profession'][:8]:10}{ch:12}p{r['page_no']}"
                      + (f"  ←继承{r['inherited']}" if r["inherited"] else ""))
                if a.code:      # 精查打印五要素
                    print(f"  分部: {r.get('chapter') or ''}｜来源: {r.get('source') or ''}"
                          f"（{'河北补充' if r.get('is_supplement') else '国标'}）")
                    print(f"  项目特征: {(r['features'] or '').replace(chr(10), '；')}")
                    print(f"  工作内容: {(r['work_content'] or '').replace(chr(10), '；')}")
                    print(f"  计算规则: {r['calc_rule'] or ''}")
                    if r["note"]:
                        print(f"  ⚠ {r['note']}")
            print(f"\n共 {len(rows)} 条"
                  f"（河北补充项目编码第7位 B / 国标项目；可加 --code 精查五要素）")
    elif a.cmd == "suggest":
        r = q.suggest(a.query, a.profession, a.books, a.limit)
        if a.json:
            print(json.dumps(r, ensure_ascii=False, indent=1))
        else:
            b = r["清单"]
            if b:
                print(f"清单 {b['code']} {b['name']}（{b.get('unit') or '—'}｜{b['profession']}）")
                if b.get("chapter"):
                    print(f"  分部: {b['chapter']}")
                if b.get("features"):
                    print(f"  项目特征: {b['features'].replace(chr(10), '；')}")
            else:
                print(f"查询「{r['查询']}」（未匹配到清单项目，直接按文字检索子目）")
            print(f"  检索册: {'、'.join(r['检索册'])}")
            if r["映射"]:
                print("\n★ 已沉淀映射（bill_map，人工确认过）:")
                for m in r["映射"]:
                    combo = " + ".join(f"{i['册']} {i['编号']}×{i.get('含量', 1)}"
                                       for i in m["items"])
                    print(f"  [{m['feature_key']}] {combo}"
                          + (f"｜{m['note']}" if m.get("note") else "")
                          + f"｜用过{m['hits']}次")
            print(f"\n候选子目（前 {len(r['建议'])}，评分只用名称/章节/单位的确定性信号）:")
            for i, s2 in enumerate(r["建议"], 1):
                coef = "—" if s2["含量"] is None else f"{s2['含量']:g}"
                print(f"  {i}. {s2['册']} {s2['编号']:10}{s2['名称'][:26]:28}"
                      f"[{s2['单位'] or '—':6}] 含量建议{coef:>8}  评分{s2['评分']}")
                print(f"     {s2['理由']}")
            if r["提示"]:
                print("\n提示:")
                for h in r["提示"]:
                    print(f"  · {h}")
            if r["建议"]:
                print("\n可直接粘进 bill 清单行的『子目』段:")
                pick = [s2 for s2 in r["建议"] if not s2["废止"]][:2]
                snippet = [{"册": s2["册"], "编号": s2["编号"],
                            "含量": s2["含量"] if s2["含量"] is not None else 1}
                           for s2 in pick]
                print("  " + json.dumps(snippet, ensure_ascii=False))
                print("  （含量按单位换算给的默认值；多子目组合/规格档位须人工确认后 map-add 沉淀）")
    elif a.cmd == "map-add":
        items = []
        for part in re.split(r"[,，]", a.items):
            part = part.strip()
            if not part:
                continue
            bits = part.split(":")
            if len(bits) < 2:
                print(f"跳过无法解析的段: {part}（格式 册:编号[:含量]）")
                continue
            items.append({"册": bits[0].strip(), "编号": bits[1].strip(),
                          "含量": float(bits[2]) if len(bits) > 2 and bits[2].strip() else 1.0})
        print(json.dumps(q.map_add(a.bill_code, items, a.note or "", a.feature_key),
                         ensure_ascii=False, indent=1))
    elif a.cmd == "map-list":
        rows = q.map_list(a.bill_code, a.limit)
        if not rows:
            print("bill_map 还没有映射（map-add 沉淀后此处可查）")
        for m in rows:
            print(f"{m['bill_code']} {m['bill_name']}｜{m['feature_key']}｜用过{m['hits']}次")
            for i in m["items"]:
                print(f"   {i['册']} {i['编号']}×{i.get('含量', 1)}  {i.get('名称', '')[:24]}"
                      + (f"｜{i['说明']}" if i.get("说明") else ""))
            if m.get("note"):
                print(f"   备注: {m['note']}")
    elif a.cmd == "search":
        for r in q.search(a.keyword, a.book, a.limit):
            print(f"{r['book_id']:<3} {r['code']:<8} {r['name'][:40]:<40} {r['unit'] or '':<6} p{r['page_no']} [{r['review_status']}]")
    elif a.cmd == "show":
        print(json.dumps(q.show(a.code, a.book), ensure_ascii=False, indent=1))
    elif a.cmd == "price":
        # 册名即专业名：园林这类没有费率的册要如实报错，不能悄悄回落到建筑费率
        prof = a.profession or BOOK_PROF.get(a.book, a.book) or "建筑"
        fee_cat = a.fee or DEFAULT_FEE.get(prof, "一般-二类")
        print(json.dumps(q.price(a.code, a.book, a.region, a.period, prof, fee_cat, a.tax),
                         ensure_ascii=False, indent=1))
    elif a.cmd == "total":
        # 册名即专业名：园林这类没有费率的册要如实报错，不能悄悄回落到建筑费率
        prof = a.profession or BOOK_PROF.get(a.book, a.book) or "建筑"
        fee_cat = a.fee or DEFAULT_FEE.get(prof, "一般-二类")
        print(json.dumps(q.total(a.code, a.book, a.region, a.period, prof, fee_cat, a.tax,
                                 a.qty, a.adjust, a.other, a.safety_cat, a.build_area),
                         ensure_ascii=False, indent=1))
    elif a.cmd == "bill":
        spec = json.load(open(a.spec, encoding="utf-8"))
        r = q.bill(spec, a.region, a.period)
        if "错误" in r:
            print(json.dumps(r, ensure_ascii=False, indent=1))
        else:
            print(f"{'编号':<14}{'名称':<26}{'单位':<6}{'工程量':>10}{'综合单价':>12}{'合价':>14}")
            for ln in r["清单"]:
                print(f"{ln['编号']:<14}{(ln['名称'] or '')[:24]:<26}{ln['单位'] or '':<6}"
                      f"{ln['工程量']:>10.3f}{ln['综合单价(除税)']:>12,.2f}{ln['合价']:>14,.2f}")
            print("\n汇总（元）")
            for k, v in r["汇总"].items():
                print(f"  {k:<24}{v:>16,.2f}")
            if r.get("措施项目"):
                print("\n措施项目（05 附件1）")
                for m in r["措施项目"]:
                    if m.get("方式") == "手填":
                        print(f"  {'—':<6}{(m.get('名称') or '')[:26]:<28}"
                              f"{'手填':>8}{m['金额']:>14,.2f}")
                    else:
                        print(f"  {m['编号']:<6}{(m['名称'] or '')[:26]:<28}"
                              f"{m['费率%']:>7.2f}%{m['金额']:>13,.2f}"
                              f"  基数{m['基数']:,.2f}"
                              + (f" × 系数{m['系数']}" if m['系数'] != 1 else ""))
            if r.get("措施基数"):
                print("\n措施基数（人工费+机械费，按专业工程分列）")
                for k, v in r["措施基数"].items():
                    print(f"  {k:<44}{v:>14,.2f}")
            v = r["校验"]
            print(f"\n校验：缺价项 {v['缺价项合计']}｜未计价主材 {v['未计价主材合计']}"
                  f"｜错误 {len(v['错误'])}｜单位提示 {len(v['单位提示'])}")
            for w in v["单位提示"]:
                print("  ⚠", w)
            if r.get("提示"):
                print("提示：", r["提示"])
            if a.out:
                json.dump(r, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
                print("WROTE", a.out)
    elif a.cmd == "gaps":
        r = q.gaps(a.book, a.region, a.period, a.top, a.kind, with_candidates=bool(a.csv))
        print(f"取价 {r['取价']}｜缺口 {r['缺口条数']} 条（按涉及子目数排序，前 {len(r['明细'])} 条）")
        print(f"{'类别':<4}{'名称':<32}{'单位':<6}{'涉及子目':>6}{'消耗量行':>7}  首选候选（ref_id）")
        for g in r["明细"]:
            cands = g.get("候选") or []
            c0 = cands[0] if cands else None
            tip = (f"#{c0['ref_id']} {c0['名称']} {c0['规格'] or ''} [{c0['理由']}]" if c0 else "")
            print(f"{g['类别']:<4}{g['名称'][:30]:<32}{(g['单位'] or ''):<6}"
                  f"{g['涉及子目']:>6}{g['消耗量行']:>7}  {tip}")
        if a.csv:
            import csv
            with open(a.csv, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["类别", "缺价名称", "单位", "涉及子目", "消耗量行",
                            "候选1_ref_id", "候选1_表", "候选1_名称", "候选1_规格", "候选1_单位", "候选1_单价", "候选1_理由",
                            "候选2_ref_id", "候选2_名称", "候选2_规格", "候选2_理由",
                            "候选3_ref_id", "候选3_名称", "候选3_规格", "候选3_理由"])
                for g in r["明细"]:
                    row = [g["类别"], g["名称"], g["单位"], g["涉及子目"], g["消耗量行"]]
                    cs = g.get("候选") or []
                    for i in range(3):
                        c = cs[i] if i < len(cs) else None
                        if i == 0:
                            row += [c["ref_id"] if c else "", c.get("表", "") if c else "",
                                    c["名称"] if c else "",
                                    c["规格"] if c else "", c["单位"] if c else "",
                                    c["单价"] if c else "", c["理由"] if c else ""]
                        else:
                            row += [c["ref_id"] if c else "", c["名称"] if c else "",
                                    c["规格"] if c else "", c["理由"] if c else ""]
                    w.writerow(row)
            print("WROTE", a.csv)
        print("\n→ 确认后在「人看的那一列」选定 ref_id，执行："
              "quota.py alias-add \"<缺价名称>\" --ref <id> [--kind machine]")
        print("  ⚠ 机械类必须加 --kind machine（material_price / machine_price 的 id 各自独立）；"
              "再跑一次 gaps 复核。")
    elif a.cmd == "alias-add":
        print(json.dumps(q.alias_add(a.term, a.ref, a.note or "", a.kind), ensure_ascii=False, indent=1))
    elif a.cmd == "export":
        print("EXPORTED", q.export_jsonl(a.out, a.book, a.only_ok))


if __name__ == "__main__":
    main()

