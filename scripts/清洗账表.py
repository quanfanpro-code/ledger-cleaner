"""把同类账户组合账表转换成借贷两列。"""
import argparse
from collections import Counter, defaultdict
from datetime import date
import json
from pathlib import Path
import re
import sys
import tempfile
from zipfile import BadZipFile

try:
    from openpyxl import load_workbook
    from openpyxl.utils.cell import column_index_from_string, get_column_letter
    from openpyxl.utils.exceptions import InvalidFileException
    import xlsxwriter
except ImportError as exc:
    raise SystemExit('缺少 Excel 读写组件。请由 Agent 按技能中的安装说明补齐后再运行。') from exc

from 读取账表 import 输入错误, 字段, 规范, 文本, 有值, 读表, 查看结构, 金额, 位置
from 接入鼎信诺 import 辅助表, 核对列, 应用辅助, 标记上级, 查看旧表, 维度资料, 维度来源, 读鼎信诺


ROOT = Path(__file__).resolve().parents[1]
表名 = {'序时账':'序时账-借贷两列.xlsx', '余额表':'余额表-借贷两列.xlsx'}
表头 = {
    '序时账': ['日期','凭证类型','凭证号','摘要','科目编码','科目名称','借方金额','贷方金额','辅助核算编码','辅助核算名称','一级编码','一级科目'],
    '余额表': ['科目编码','科目名称','辅助核算编码','辅助核算名称',
             '期初余额(借方)','期初余额(贷方)','累计发生额(借方)','累计发生额(贷方)','期末余额(借方)','期末余额(贷方)','科目级次'],
}
原金额 = {'序时账': ('借方发生额','贷方发生额'), '余额表': ('期初净额','借方发生额','贷方发生额','期末净额')}
金额列 = {'序时账': (6,7), '余额表': (4,6,7,8)}
可选来源 = {1,3,5,6,7,8,9,'loan','interest','year','literal','currency'}
期间格式 = re.compile(r'\d{4}-(?:0[1-9]|1[0-2])~\d{4}-(?:0[1-9]|1[0-2])')


def 准备参数():
    p = argparse.ArgumentParser(description='自动识别并清洗序时账或科目余额表。')
    p.add_argument('--输入', type=Path, nargs='+', help='一个或多个源工作簿，顺序随意')
    p.add_argument('--序时账', type=Path, help='已确认此文件为序时账时使用')
    p.add_argument('--余额表', type=Path, help='已确认此文件为余额表时使用')
    p.add_argument('--工作表', help='只提供一个文件时，限制到指定工作表')
    p.add_argument('--序时账工作表')
    p.add_argument('--余额表工作表')
    p.add_argument('--公司标识', help='账套未标注时使用已经确认的本批公司标识')
    p.add_argument('--余额期间', help='鼎信诺余额表未标注期间时，使用已核实的起止月份，例如 2025-01~2025-12')
    p.add_argument('--输出目录', type=Path)
    p.add_argument('--预检', action='store_true', help='识别、对账并报告歧义，不生成结果文件')
    p.add_argument('--科目规则', type=Path, help='用户确认后的额外科目规则文件')
    p.add_argument('--表头别名', type=Path, help='用户确认的表头别名文件')
    p.add_argument('--指定列', type=Path, help='用户确认的列位置文件，用于解决重名列')
    p.add_argument('--查看结构', action='store_true', help='只查看工作表、合并区域、列位置和样例行')
    p.add_argument('--样例行数', type=int, default=18, help='查看结构时读取前多少行')
    p.add_argument('--版式方案', type=Path, help='Agent 根据原表证据整理的版式与字段位置')
    p.add_argument('--表头扫描行数', type=int, default=120)
    p.add_argument('--金额小数位', type=int, choices=range(0,7), default=2)
    return p.parse_args()


def 选择文件(choose_destination):
    if sys.platform != 'win32':
        raise 输入错误('请由 Agent 传入选好的工作簿。文件选择窗口面向 Windows。')
    try:
        from tkinter import Tk, filedialog, TclError
    except ImportError as exc:
        raise 输入错误('本机没有可用的 Windows 文件选择组件，请让 Agent 传入选好的文件。') from exc
    try:
        root = Tk()
        root.withdraw()
        try:
            files = filedialog.askopenfilenames(title='选择序时账、余额表或核算项目表', filetypes=[('账表文件','*.xlsx *.xlsm *.xls *.tsv *.txt')])
            if not files:
                raise 输入错误('未选择文件。')
            directory = filedialog.askdirectory(title='选择交付位置') if choose_destination else ''
            if choose_destination and not directory:
                raise 输入错误('未选择输出位置。')
            return [Path(name) for name in files], Path(directory) if directory else None
        finally:
            root.destroy()
    except TclError as exc:
        raise 输入错误('当前无法打开文件选择窗口，请通过 Agent 提供文件。') from exc


def 载入规则(override):
    exact = json.loads((ROOT/'references'/'科目规则.json').read_text(encoding='utf-8-sig'))
    updates = json.loads(override.read_text(encoding='utf-8-sig')) if override else {}
    if not isinstance(updates, dict):
        raise 输入错误('科目规则文件应以科目编码为索引。')
    for code, row in updates.items():
        if (not isinstance(row, list) or len(row) not in (3,4) or
            not all(isinstance(v,str) and v.strip() for v in row[:2]) or
            not isinstance(row[2],(str,int)) or isinstance(row[2],bool) or
            (row[2] not in 可选来源 and not (isinstance(row[2],str) and row[2].startswith('鼎信诺:') and 规范(row[2][4:]))) or
            (row[2] == 'literal') != (len(row) == 4) or
            (len(row) == 4 and (not isinstance(row[3],str) or not row[3].strip()))):
            raise 输入错误(f'科目 {code} 的自定义规则格式不完整，请核对科目名、维度和取值来源。')
    return exact | updates, updates


def 载入别名(file):
    if not file:
        return {}
    data = json.loads(file.read_text(encoding='utf-8-sig'))
    if not isinstance(data,dict):
        raise 输入错误('表头别名应以字段为索引。')
    result = {}
    for k, names in data.items():
        if k not in 字段 or not isinstance(names,list) or not names or not all(isinstance(n,str) and 规范(n) for n in names):
            raise 输入错误(f'无效的表头别名：{k}')
        result[k] = {规范(n) for n in names}
    return result


def 载入指定列(file):
    if not file:
        return {}
    data = json.loads(file.read_text(encoding='utf-8-sig'))
    if not isinstance(data,dict) or not data.keys() <= 表名.keys():
        raise 输入错误('指定列文件需要以「序时账」或「余额表」分类。')
    result = {}
    for kind, mapping in data.items():
        if not isinstance(mapping,dict) or not mapping.keys() <= 字段.keys():
            raise 输入错误(f'{kind} 的指定列文件含不认识的字段。')
        result[kind] = {}
        for key, value in mapping.items():
            if not isinstance(value,str) or not re.fullmatch('[A-Za-z]{1,3}',value.strip()):
                raise 输入错误(f'{kind}「{key}」的指定列应为 Excel 列字母。')
            col = column_index_from_string(value.strip())
            if col > 16384:
                raise 输入错误(f'{kind}「{key}」的指定列超出 Excel 范围。')
            result[kind][key] = col-1
    return result


def 载入版式(file, input_names):
    if not file:
        return {}
    data = json.loads(file.read_text(encoding='utf-8-sig'))
    if not isinstance(data,dict) or not data or len(input_names) != len(set(input_names)) or not data.keys() <= set(input_names):
        raise 输入错误('版式方案需以本次输入中互不重名的文件名为索引。')
    required = {'类型','工作表','表头末行','列'}
    allowed = required | {'公司标识','会计期间'}
    result = {}
    for name,entry in data.items():
        if not isinstance(entry,dict) or not required <= entry.keys() <= allowed:
            raise 输入错误(f'{name} 的版式方案应包含类型、工作表、表头末行和列。')
        kind,title,end,fields = (entry[key] for key in ('类型','工作表','表头末行','列'))
        if not isinstance(kind,str) or kind not in 表名 or not isinstance(title,str) or not title.strip():
            raise 输入错误(f'{name} 的版式方案需要有效的类型和工作表名称。')
        if type(end) is not int or not 1 <= end <= 10000 or not isinstance(fields,dict) or not fields.keys() <= 字段.keys():
            raise 输入错误(f'{name} 的版式方案表头行号或字段不合法。')
        columns = {}
        for key,value in fields.items():
            if not isinstance(value,str) or not re.fullmatch('[A-Za-z]{1,3}',value.strip()):
                raise 输入错误(f'{name} 的「{key}」需要填写 Excel 列字母。')
            col = column_index_from_string(value.strip())
            if col > 16384 or col-1 in columns.values():
                raise 输入错误(f'{name} 的「{key}」指定列超出范围或重复。')
            columns[key] = col-1
        if '公司标识' in entry and (not isinstance(entry['公司标识'],str) or not 有值(entry['公司标识'])):
            raise 输入错误(f'{name} 的公司标识需要填写来自原表的实际标识。')
        if '会计期间' in entry:
            period = entry['会计期间']
            if (kind != '余额表' or not isinstance(period,str) or
                not 期间格式.fullmatch(period) or period[:7] > period[8:]):
                raise 输入错误(f'{name} 的余额表期间需要填写起止月份，例如 2025-01~2025-12。')
        result[name] = {**entry,'列':columns}
    return result


def 补充规则(title):
    leaf = re.split(r'[-－_]', title)[-1].strip()
    account_kind = re.sub(r'^重分类[_-]', '', title)
    if title.endswith(('结转','结转损益')):
        return '结转类别', 'literal', title
    if account_kind.startswith('银行存款'):
        return '银行账户（名称和账号）', 5
    if account_kind.startswith('库存现金'):
        currency = next((c for c in ('人民币','美元','欧元','港币','日元') if c in title), None)
        return ('币种','literal',currency) if currency else None
    if account_kind.startswith('其他货币资金'):
        return '结算账户', 5
    if account_kind.startswith(('应付长期借款利息','应付短期借款利息')) or (account_kind.startswith('应付利息') and '借款' in title):
        return '贷款银行', 'interest'
    if account_kind.startswith('应付利息'):
        return '付息对象', 8
    if account_kind.startswith(('长期借款','短期借款')):
        return '贷款银行', 'loan'
    if account_kind.startswith(('应收应付中转','虚拟银行中转')):
        return '中转用途', 'literal', leaf
    if account_kind.startswith(('应收','应付账款','应付票据','应付股利','其他应付','其他应收','预付','预收','合同负债','合同资产')):
        return '往来单位或个人', 8
    if account_kind.startswith(('长期股权投资','实收资本','股本')):
        return ('被投资单位' if account_kind.startswith('长期股权投资') else '股东'), 8
    if account_kind.startswith(('固定资产','累计折旧','无形资产','累计摊销')) and leaf != title:
        return ('固定资产类别' if account_kind.startswith(('固定资产','累计折旧')) else '无形资产类别'), 'literal', leaf
    if account_kind.startswith(('应付职工薪酬','管理费用')):
        return '归属部门', 3
    if account_kind.startswith(('开发成本','开发产品','投资性房地产','主营业务收入','主营业务成本','销售费用')):
        return '归属项目', 7
    if account_kind.startswith(('应交税费','税金及附加')) and leaf != title:
        if any(x in title for x in ('房产税','土地使用税','土地增值税','预交增值税','简易计税')):
            return '涉税项目', 7
        return ('税率档' if re.fullmatch(r'\d+(?:\.\d+)?%|其他税率', leaf) else '税种或核算项目'), 'literal', leaf
    return None


def 一级科目资料(books):
    explicit, inferred = defaultdict(set), defaultdict(set)
    for book in books:
        for row in book.行:
            code, title = row.编码[3], row.名称.get(4,'')
            if row.一级编码 and row.一级科目:
                explicit[row.一级编码].add(row.一级科目)
            if re.fullmatch(r'\d{4}',code) and title:
                explicit[code].add(title)
            if re.fullmatch(r'\d{5,}',code) and title:
                prefix,*children = re.split(r'[-－_＞>]',title,maxsplit=1)
                if children and prefix.strip():
                    inferred[code[:4]].add(prefix.strip())
    return {key:min(group) for key,group in (inferred | explicit).items() if len({规范(v) for v in group}) == 1}


def 鼎信诺规则(title, root, general, dimensions):
    base = general or 补充规则(root or title)
    kind = root or title
    fixed = base and base[1] in ('loan','interest',5,8,'literal')
    eligible = dimensions
    if fixed:
        if base[1] in ('loan','interest','literal'):
            return base
        eligible = {d for d in dimensions if 维度来源(d) == base[1]}
        if kind.startswith(('应收账款','应收票据','预收','合同负债')):
            eligible = {d for d in eligible if '供应商' not in d}
        if kind.startswith(('应付账款','应付票据','预付')):
            eligible = {d for d in eligible if '客户' not in d}
    elif base:
        preferred = {d for d in dimensions if 维度来源(d) == base[1] and '合同' not in d}
        eligible = preferred or dimensions
    if len({规范(d) for d in eligible}) > 1:
        return None
    if eligible:
        dim = min(eligible)
        return (base[0] if fixed else dim), f'鼎信诺:{dim}'
    if base:
        return base
    if kind.startswith(('库存商品','发出商品','在途物资','委托加工物资','存货','原材料','周转材料','生产成本')):
        return '商品或物料',6
    if kind.startswith(('营业税金','应交税费','税金及附加','所得税费用','递延所得税')):
        return '税种或纳税项目','literal',title
    if kind.startswith(('投资','交易性金融','债权投资','其他权益','其他债券','租赁负债','长期应收','长期应付','坏账准备')):
        return '交易对象',8
    if kind.startswith(('固定资产','累计折旧','无形资产','累计摊销','使用权资产','在建工程','长期待摊','研发')):
        return '资产或项目',7
    if kind.startswith(('财务费用','费用','营业外','其他业务','其他收益','资产减值','信用资产减值','资产处置',
                        '待处理财产','其他流动','一年内到期','递延收益','资本公积','盈余公积','本年利润','利润分配','以前年度损益')):
        return '会计明细类别','literal',title
    if kind == '库存现金':
        return '币种','currency'
    return None


def 账户规则(books, exact, overrides, roots, dimensions):
    account_names = defaultdict(set)
    known = defaultdict(set)
    unresolved_names = {}
    ding, other, detail = set(), set(), set()
    for book in books:
        for row in book.行:
            company = row.编码[0]
            for segment, name in row.名称.items():
                token = row.编码[segment-1]
                if 有值(token) and 有值(name):
                    known[(company,segment,token)].add(name)
            code = row.编码[3]
            is_ding = book.补充核对 is not None
            (ding if is_ding else other).add(code)
            if book.类别 == '序时账' or row.末级:
                detail.add(code)
            raw = row.名称.get(4)
            if 有值(raw):
                account_names[code].add(raw)
            else:
                unresolved_names.setdefault(code,row.位置)

    account_rules, names, questions = {}, {}, []
    for code in sorted(account_names.keys() | unresolved_names.keys()):
        candidates = account_names[code]
        normalized = {规范(x) for x in candidates}
        if len(normalized) > 1:
            raise 输入错误(f'科目 {code} 在输入中有多个名称：{"、".join(sorted(candidates))}。请确认科目对应关系。')
        if not candidates:
            questions.append(f'{unresolved_names[code]}，科目 {code} 缺少可验证的科目名称。')
            continue
        source_title = sorted(candidates)[0]
        entry = exact.get(code)
        if entry and (规范(entry[0]) == 规范(source_title) or code in overrides):
            if 规范(entry[0]) != 规范(source_title):
                questions.append(f'已指定规则的科目 {code} 名称为 {entry[0]}，源表为 {source_title}，请核对。')
                continue
            rule = tuple(entry[1:])
            basis = '用户确认规则' if code in overrides else '内置逐科目规则'
        else:
            rule = 补充规则(source_title)
            basis = '按科目名称匹配标准规则'
            if code in ding and code not in other:
                rule = 鼎信诺规则(source_title,roots.get(code[:4]),rule,dimensions.get(code,set()))
                basis = '鼎信诺科目与源维度'
        if not rule and code in ding and code not in detail:
            rule = ('上级科目','literal',source_title)
            basis = '保留原表科目级次'
        if not rule:
            listed = '、'.join(sorted(dimensions.get(code,())))
            suffix = f' 源表候选：{listed}。' if listed else ''
            questions.append(f'科目 {code} {source_title}：现有规则无法确定该保留哪个维度。{suffix}')
            continue
        account_rules[code] = (*rule, basis)
        names[code] = source_title
    if questions:
        details = '\n'.join(questions[:30])
        suffix = f'\n另有 {len(questions)-30} 项。' if len(questions) > 30 else ''
        raise 输入错误('确认以下科目规则后再清洗：\n' + details + suffix)
    return account_rules,names,known


def 项目名称(row, segment, catalog):
    token = row.编码[segment-1]
    direct = row.名称.get(segment)
    candidates = catalog.get((row.编码[0],segment,token), set())
    distinct = {规范(v) for v in candidates}
    if segment == 7 and (token == '9999' or direct == '重分类-项目段'):
        return '', ''
    if not 有值(token):
        return ('',direct) if row.分列 and 有值(direct) else ('','')
    if len(distinct) > 1:
        return '', ''
    if len(distinct) == 1:
        return token,min(candidates)
    if 有值(direct):
        return token,direct
    return '', ''


def 贷款银行(row, catalog):
    code,name = row.贷款行
    if 有值(name):
        return (code if 有值(code) else ''),name
    token, lender = 项目名称(row,8,catalog)
    if token and token != row.编码[0] and (
        (not row.分列 and 有值(row.编码[8])) or any(term in lender for term in ('银行','信用社','信用合作','财务有限','财务公司'))
    ):
        return token,lender
    return '', ''


def 辅助规则(books, rules, catalog):
    loans = defaultdict(set)
    for book in books:
        for row in book.行:
            if rules[row.编码[3]][1] != 'loan':
                continue
            bank = 贷款银行(row,catalog)
            contract = row.合同 if row.分列 else row.编码[6]
            if bank[1] and 有值(contract):
                loans[(row.编码[0],contract)].add(bank)

    def choose(row):
        if row.辅助指定 is not None:
            return row.辅助指定
        rule = rules[row.编码[3]]
        dimension, source = rule[:2]
        if source == 'literal':
            return row.编码[3], rule[2], ''
        if source == 'year':
            return (row.年份,row.年份,'') if row.年份 else ('','会计年度未记录','会计期间无法确认')
        if source == 'currency':
            return ('',row.币种,'') if row.币种 else ('','币种未记录','原表缺币种')
        if source == 'loan':
            if not 有值(row.贷款行[1]) and row.往来冲突:
                raise 输入错误(row.往来冲突)
            token, name = 贷款银行(row,catalog)
        elif source == 'interest':
            contract = row.合同 if row.分列 else row.编码[6]
            candidates = loans.get((row.编码[0],contract), set())
            candidates = set(candidates)
            if 有值(row.贷款行[1]):
                candidates.add(贷款银行(row,catalog))
            names = {规范(name) for _,name in candidates}
            codes = {code for code,_ in candidates if 有值(code)}
            if len(names) > 1 or len(codes) > 1:
                if 有值(row.贷款行[1]):
                    raise 输入错误(f'{row.位置} 的利息贷款银行与同公司同合同的本金资料冲突。请核对贷款银行和合同编码。')
            if len(names) != 1 or len(codes) > 1:
                return '', '贷款银行待核', '同公司同合同未唯一匹配贷款银行'
            token = next(iter(codes),'')
            name = min(name for _,name in candidates)
        elif isinstance(source,str) and source.startswith('鼎信诺:'):
            token,name = '', ''
        else:
            if source == 8 and row.往来冲突:
                raise 输入错误(row.往来冲突)
            token,name = 项目名称(row,source,catalog)
        if token:
            return token,name,''
        if name:
            return '',name,f'{dimension}只有名称，按原行保留'
        if source == 7 and (row.编码[6] == '9999' or row.名称.get(7) == '重分类-项目段'):
            return '', '重分类调整', '重分类项目占位单独留痕'
        label = '往来' if source == 8 else ('项目' if source == 7 else dimension)
        return '', f'{label}未记录', f'{dimension}缺值或同一编码名称有冲突'
    return choose


def 向量合计(rows, size):
    amounts = [0] * size
    for row in rows:
        for index, value in enumerate(row):
            amounts[index] += value
    return tuple(amounts)


def 银行跨表对应(journal_aux, balance_aux, journal_accounts, balance_accounts):
    links = []
    for identity in sorted(journal_accounts.keys() & balance_accounts.keys()):
        left, right = journal_accounts[identity], balance_accounts[identity]
        if len(left) != 1 or len(right) != 1:
            continue
        j_key,j_numbers = next(iter(left.items()))
        b_key,b_numbers = next(iter(right.items()))
        numbers = j_numbers | b_numbers
        if len(numbers) != 1 or j_key == b_key:
            continue
        # 至少一边有独立账号列，另一边的完整开户行及账号与其一致。
        target, previous, moving = (j_key,b_key,balance_aux) if j_numbers else (b_key,j_key,journal_aux)
        if target in moving:
            continue
        moving[target] = moving.pop(previous)
        links.append((*identity[:2],j_key[2],b_key[2],next(iter(numbers)),target[3]))
    return links


def 准备输出(books, rules, names, choose, roots):
    results, traces, checks, warnings = {}, {}, [], []
    movements, bank_accounts = {}, {}
    for book in books:
        kind = book.类别
        data = []
        trace = []
        amounts = {}
        group_keys = {}
        by_subject = defaultdict(lambda:[0,0])
        by_aux = defaultdict(lambda:[0,0])
        banks = defaultdict(lambda:defaultdict(set))
        vouchers = defaultdict(int)
        balance_layers = defaultdict(list)
        leaf_rows = []
        for position, row in enumerate(book.行):
            company, code = row.编码[0], row.编码[3]
            aux, aux_name, issue = choose(row)
            if issue:
                warnings.append((kind,str(row.位置),code,issue))
            movement = row.金额 if kind == '序时账' else row.金额[1:3]
            aux_key = (company,code,aux,aux_name)
            destinations = []
            if not row.行类型 or (row.行类型 == '科目汇总' and row.末级):
                destinations.append(((company,code),by_subject,movement))
            if row.辅助分项:
                for object_code,object_name,part in row.辅助分项:
                    destinations.append(((company,code,object_code,object_name),by_aux,part))
                allocated = 向量合计((part for _,_,part in row.辅助分项),2)
                remainder = tuple(value-used for value,used in zip(movement,allocated,strict=True))
                if any(remainder):
                    destinations.append(((company,code,'','未列明辅助项目'),by_aux,remainder))
            elif row.行类型 != '科目汇总':
                destinations.append((aux_key,by_aux,movement))
            for key,destination,part in destinations:
                for i,value in enumerate(part):
                    destination[key][i] += value
            if rules[code][0] in ('银行账户（名称和账号）','结算账户') and aux and aux_name:
                label, number = 规范(aux_name),规范(row.银行账号)
                observed = banks[(company,code,label)][aux_key]
                if number and aux == row.银行账号 and label.endswith(number) and len(label) > len(number):
                    observed.add(row.银行账号)
            if kind == '序时账':
                vouchers[(company,row.期间,row.类型,row.凭证)] += movement[0] - movement[1]
                first_code = row.一级编码 or (code[:4] if re.fullmatch(r'\d{4,}',code) else '')
                first_name = row.一级科目 or roots.get(first_code)
                if not first_name:
                    warnings.append((kind,str(row.位置),code,'源表未提供可确认的一级科目名称'))
                data.append((row.日期,row.类型 or None,row.凭证,row.摘要 or None,code,names[code],*row.金额,
                             aux or None,aux_name,first_code or None,first_name))
                output_row = len(data)+1
            elif row.行类型:
                o,d,c,e = row.金额
                level = row.科目级次 if row.行类型 == '科目汇总' else None
                data.append((code,names[code],aux or None,aux_name or None,o,None,d,c,e,None,level))
                if row.行类型 == '科目汇总' and row.末级:
                    leaf_rows.append(data[-1])
                output_row = len(data)+1
                balance_layers[(company,code)].append((row.行类型,row.金额))
            else:
                key = (company,code,aux,aux_name,row.科目级次) if aux else (company,code,position)
                if key not in amounts:
                    amounts[key] = [0,0,0,0]
                    group_keys[key] = (len(amounts)+1,code,names[code],aux or None,aux_name,row.科目级次)
                for i,value in enumerate(row.金额):
                    amounts[key][i] += value
                output_row = group_keys[key][0]
            if kind == '余额表' and row.行类型 != '未列明辅助项目' and row.金额[0]+row.金额[1]-row.金额[2] != row.金额[3]:
                warnings.append((kind,str(row.位置),code,'原表期初加借减贷不等于期末，原金额已保留'))
            trace.append((row.位置.文件,row.位置.页,row.位置.行,row.行号 or None,company,code,row.原始组合 or None,rules[code][0],
                          aux or None,aux_name,output_row))
        if kind == '余额表':
            for (_,code),layer in balance_layers.items():
                if len(layer) > 1 and layer[0][1] != 向量合计((amount for _,amount in layer[1:]),4):
                    raise 输入错误(f'科目 {code} 的辅助明细与未列明项目无法对回科目汇总。')
            for key,(o,d,c,e) in amounts.items():
                _,code,name,aux,aux_name,level = group_keys[key]
                data.append((code,name,aux,aux_name,o,None,d,c,e,None,level))
        else:
            if any(vouchers.values()):
                bad_vouchers = sum(v != 0 for v in vouchers.values())
                warnings.append((kind,book.文件,'',f'{bad_vouchers} 个凭证分组借贷发生额不同，原分录已保留'))
        results[kind], traces[kind] = data, trace
        movements[kind] = (by_subject,by_aux)
        bank_accounts[kind] = banks
        size = len(原金额[kind])
        controlled = book.控制合计 is not None
        source_total = book.控制合计 if controlled else 向量合计((r.金额 for r in book.行),size)
        output_total = 向量合计((tuple(r[i] for i in 金额列[kind]) for r in (leaf_rows if controlled else data)),size)
        if source_total != output_total:
            raise 输入错误(f'{kind} 清洗前后金额不一致，停止交付。')
        for field,a,b in zip(原金额[kind],source_total,output_total):
            checks.append((kind,field,a,b,a-b,'清洗前后一致'))
        for pos,amount in book.合计:
            if amount != source_total:
                warnings.append((kind,str(pos),'','源表合计行和明细合计不同，建议检查源表范围'))
        warnings.extend((kind,str(pos),'',f'{reason}：{text}') for pos,reason,text in book.警示)
    differences, auxiliary, bank_links = [], [], []
    cross_status = '单表运行'
    if len(books) == 2:
        journal = next(b for b in books if b.类别 == '序时账')
        balance = next(b for b in books if b.类别 == '余额表')
        companies = {r.编码[0] for r in journal.行}
        same_company = '本批账套' not in companies and companies == {r.编码[0] for r in balance.行}
        same_period = balance.期间 and journal.期间 == balance.期间
        cross_status = '公司、会计期间可对应' if same_company and same_period else '公司或会计期间未确认一致，跨表差异供参考'
        for index,destination in [(0,differences),(1,auxiliary)]:
            j,b = movements['序时账'][index],movements['余额表'][index]
            if index == 1 and same_company and same_period:
                bank_links = 银行跨表对应(j,b,bank_accounts['序时账'],bank_accounts['余额表'])
            for key in sorted(j.keys()|b.keys()):
                jd,jc = j.get(key,(0,0))
                bd,bc = b.get(key,(0,0))
                if (jd,jc) == (bd,bc):
                    continue
                company, code, *aux = key
                destination.append((company,code,names[code],*aux,jd,jc,bd,bc,jd-bd,jc-bc))
        for field,i in (('跨表借方发生额',0),('跨表贷方发生额',1)):
            j = sum(v[i] for v in movements['序时账'][0].values())
            b = sum(v[i] for v in movements['余额表'][0].values())
            checks.append(('跨表',field,j,b,j-b,cross_status))
    return results,traces,checks,warnings,differences,auxiliary,bank_links,cross_status


def _write_book(path, sheets, precision):
    opts = {'constant_memory':True,'strings_to_formulas':False,'strings_to_urls':False,'strings_to_numbers':False}
    with xlsxwriter.Workbook(path,opts) as book:
        header_fmt = book.add_format({'bold':True,'font_color':'white','bg_color':'#174D6C'})
        amount_fmt = book.add_format({'num_format':'#,##0' + ('.'+'0'*precision if precision else '')})
        date_fmt = book.add_format({'num_format':'yyyy-mm-dd'})
        for name,columns,rows,num_cols,widths in sheets:
            sheet = book.add_worksheet(name)
            sheet.freeze_panes(1,0)
            sheet.hide_gridlines(2)
            for i,width in enumerate(widths):
                sheet.set_column(i,i,width)
            for i,title in enumerate(columns):
                sheet.write(0,i,title,header_fmt)
            for r,row in enumerate(rows,1):
                if r >= 1048576:
                    raise 输入错误(f'{name} 的结果超过 Excel 单张表可容纳的行数。')
                for c,value in enumerate(row):
                    if value is None:
                        continue
                    if c in num_cols:
                        if abs(value) >= 10**15:
                            raise 输入错误(f'{name} 第{r+1}行金额超过 Excel 可安全回读的精度。')
                        sheet.write_number(r,c,value/(10**precision),amount_fmt)
                    elif isinstance(value,date):
                        sheet.write_datetime(r,c,value,date_fmt)
                    else:
                        if sheet.write(r,c,value) != 0:
                            raise 输入错误(f'{name} 第{r+1}行的值超出Excel单元格允许长度。')
            if rows:
                sheet.autofilter(0,0,len(rows),len(columns)-1)


def _verify_output(path, kind, expected, precision):
    from datetime import datetime
    book = load_workbook(path,read_only=True,data_only=True)
    try:
        iterator = book[kind].iter_rows(max_col=len(表头[kind]), values_only=True)
        if list(next(iterator)) != 表头[kind]:
            raise 输入错误(f'{kind} 保存后表头复核失败。')
        count = 0
        for index,(got,wanted) in enumerate(zip(iterator,expected,strict=True),2):
            actual = list(got)
            for col in 金额列[kind]:
                actual[col] = 金额(actual[col],位置(path.name,kind,index),col,precision)
            if kind == '序时账' and isinstance(actual[0],datetime):
                actual[0] = actual[0].date()
            if tuple(actual) != wanted:
                raise 输入错误(f'{kind} 保存后第{index}行回读结果有变化。')
            count += 1
        if count != len(expected):
            raise 输入错误(f'{kind} 保存后条数复核失败。')
    except ValueError as exc:
        raise 输入错误(f'{kind} 保存后回读校验失败：{exc}') from exc
    finally:
        book.close()


def 写结果(args, books, supplements, rules, names, prepared):
    results,traces,checks,warnings,differences,auxiliary,bank_links,_ = prepared
    output_dir = (args.输出目录 or Path.cwd()).resolve()
    output_dir.mkdir(parents=True,exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.处理中-',dir=output_dir))
    final = output_dir / stage.name.replace('.处理中-','账表清洗-',1)
    precision = args.金额小数位
    try:
        for kind,data in results.items():
            widths = [14,14,20,48,16,35,17,17,22,37,17,25] if kind == '序时账' else [16,43,22,40,18,18,19,19,18,18,23]
            _write_book(stage/表名[kind],[(kind,表头[kind],data,金额列[kind],widths)],precision)
            _verify_output(stage/表名[kind],kind,data,precision)
        sources = [*books,*supplements.values()]
        overview = [(b.类别,b.文件,b.页,b.表头,b.期间 or '未识别',
                     b.源行数 if b.源行数 is not None else len(b.行),
                     len(results[b.类别]) if b.类别 in results else None) for b in sources]
        trace_columns = ['源文件','源工作表','源表行号','原分录行号','公司标识','科目编码','原账户组合','选定维度','辅助核算编码','辅助核算名称','结果表行号']
        diff_cols = ['序时账借方','序时账贷方','余额表借方','余额表贷方','借方差额','贷方差额']
        sheets = [
            ('读表结果',['表种类','源文件','工作表','识别的表头末行','会计期间','源表明细行数','输出明细行数'],overview,(),[14,36,24,19,22,17,17]),
            ('字段定位',['表种类','源文件','工作表','字段','原表列','识别方式'],
             [(b.类别,b.文件,b.页,key,get_column_letter(col+1),b.识别方式)
              for b in sources for key,col in sorted(b.字段.items(),key=lambda item:item[1])],(),[14,36,24,19,12,22]),
            ('金额核对',['范围','项目','原表或序时账','结果或余额表','差额','说明'],checks,(2,3,4),[13,21,22,22,18,49]),
            ('待核记录',['表种类','位置','科目编码','原因'],warnings,(),[14,51,18,60]),
            ('科目规则',['科目编码','科目名称','选定维度','规则依据'],
             [(code,names[code],r[0],r[-1]) for code,r in sorted(rules.items())],(),[18,48,29,31]),
        ]
        if len(books) == 2:
            sheets.extend([
                ('跨表科目差异',['公司编码','科目编码','科目名称',*diff_cols],differences,range(3,9),[16,16,44,*([19]*6)]),
                ('跨表辅助差异',['公司编码','科目编码','科目名称','辅助核算编码','辅助核算名称',*diff_cols],auxiliary,range(5,11),[16,16,42,19,35,*([19]*6)]),
            ])
        if bank_links:
            sheets.append(('跨表银行对应',['公司标识','科目编码','序时账辅助编码','余额表辅助编码','实际银行账号','开户行及账号'],
                           bank_links,(),[16,17,23,23,26,48]))
        native_audit = [item for book in books for item in (book.补充核对 or [])]
        if native_audit:
            sheets.append(('辅助匹配',核对列,native_audit,range(8,12),[38,48,48,18,21,19,34,15,*([19]*4)]))
        for kind,trace in traces.items():
            sheets.append((kind+'来源',trace_columns,trace,(),[34,25,15,15,16,16,48,26,20,38,17]))
        report = stage/'清洗核对.xlsx'
        _write_book(report,sheets,precision)
        audit = load_workbook(report,read_only=True,data_only=True)
        try:
            rows = list(audit['金额核对'].values)
            if len(rows) != len(checks)+1 or len(audit.sheetnames) != len(sheets):
                raise 输入错误('核对工作簿保存后结构复核失败。')
        finally:
            audit.close()
        stage.rename(final)
    except Exception as exc:
        raise 输入错误(f'结果尚未通过交付。现场文件保留在：{stage}。原因：{exc}') from exc
    return final


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    args = 准备参数()
    inp = [('',path,None) for path in (args.输入 or [])]
    inp += [(kind,path,sheet) for kind,path,sheet in (
        ('序时账',args.序时账,args.序时账工作表),('余额表',args.余额表,args.余额表工作表)
    ) if path]
    if not inp:
        paths, directory = 选择文件(not args.预检 and not args.查看结构 and not args.输出目录)
        args.输出目录 = args.输出目录 or directory
        inp = [('',path,None) for path in paths]
    if not 1 <= len(inp) <= 4:
        raise 输入错误('每批最多一份序时账、一份科目余额表和各自的一份配套核算项目表。')
    if args.工作表 and len(inp) != 1:
        raise 输入错误('指定单个工作表时，只能提供一份文件。两份文件请分别指定其类型和工作表。')
    if args.表头扫描行数 < 4:
        raise 输入错误('表头扫描行数应至少为4。')
    if not 1 <= args.样例行数 <= 200:
        raise 输入错误('样例行数应在1至200之间。')
    if args.公司标识 is not None and not 有值(args.公司标识):
        raise 输入错误('请使用有效的本批公司标识。')
    period = args.余额期间
    if period is not None and (not 期间格式.fullmatch(period) or period[:7] > period[8:]):
        raise 输入错误('余额期间须为从早到晚的起止月份，例如 2025-01~2025-12。')
    if len({file.resolve() for _,file,_ in inp}) != len(inp):
        raise 输入错误('同一文件重复提交。')
    for _,file,_ in inp:
        if file.suffix.casefold() not in ('.xlsx','.xlsm','.xls','.tsv','.txt') or not file.is_file():
            raise 输入错误(f'需要可读的 Excel 或鼎信诺账表：{file.name}。')
    if args.查看结构:
        print(json.dumps([(查看结构 if file.suffix.casefold() in ('.xlsx','.xlsm') else 查看旧表)(file,args.样例行数)
                          for _,file,_ in inp],ensure_ascii=False,default=str))
        return
    overrides = 载入别名(args.表头别名)
    assigned = 载入指定列(args.指定列)
    layouts = 载入版式(args.版式方案,[file.name for _,file,_ in inp])
    books,supplements = [],{}
    seen = {}
    for kind,file,sheet in inp:
        plan = layouts.get(file.name)
        selected_sheet = sheet or args.工作表
        native = None if plan else 读鼎信诺(file,selected_sheet,args.表头扫描行数,args.金额小数位)
        if native is None and file.suffix.casefold() not in ('.xlsx','.xlsm'):
            raise 输入错误(f'{file.name} 无法自动识别旧版账表。请查看原表结构，或另存为新版 Excel 后用版式方案适配。')
        book = native or 读表(file,kind,selected_sheet,args.表头扫描行数,overrides,assigned,args.金额小数位,plan)
        if kind and book.类别 != kind:
            raise 输入错误(f'{file.name} 实际识别为{book.类别}，与指定的{kind}不同。')
        if book.类别 in seen:
            raise 输入错误(f'{seen[book.类别]} 和 {file.name} 都识别为{book.类别}，请确认使用哪份。')
        seen[book.类别] = file.name
        if native and args.公司标识:
            for row in book.行:
                if row.编码[0] != '本批账套' and 规范(row.编码[0]) != 规范(args.公司标识):
                    raise 输入错误(f'{row.位置} 公司标识与确认的本批账套不一致。')
                row.编码 = (args.公司标识,*row.编码[1:])
        if native and period and book.类别 in ('余额表','核算项目余额表'):
            if book.期间 and book.期间 != period:
                raise 输入错误(f'{book.文件} 的原表期间 {book.期间} 与指定期间 {period} 不一致。')
            if not book.期间:
                book.警示.append((位置(book.文件,book.页,book.表头),'源表未标注期间',f'本次按已核实的 {period} 处理'))
                book.期间 = period
                for row in book.行:
                    row.年份 = period[:4] if period[:4] == period[8:12] else ''
        if native and book.类别 == '余额表' and '末级' not in book.字段:
            标记上级(book.行)
        if book.类别 in 辅助表:
            supplements[book.类别] = book
        else:
            books.append(book)
    if not books:
        raise 输入错误('只识别到了核算项目表，请补充相应的序时账或科目余额表。')
    if period and not any(b.类别 == '余额表' and b.补充核对 is not None for b in books):
        raise 输入错误('余额期间参数适用于鼎信诺科目余额表，其他源表可在版式方案里填写会计期间。')
    exact,user_rules = 载入规则(args.科目规则)
    roots = 一级科目资料(books)
    rules,names,catalog = 账户规则(books,exact,user_rules,roots,维度资料(books,supplements))
    应用辅助(books,supplements,rules,names)
    prepared = 准备输出(books,rules,names,辅助规则(books,rules,catalog),roots)
    for book in supplements.values():
        print(f'{book.文件} → {book.类别}，工作表「{book.页}」，表头至第{book.表头}行，{len(book.行)}条。')
    for book in books:
        print(f'{book.文件} → {book.类别}，工作表「{book.页}」，表头至第{book.表头}行，{len(book.行)}条。')
    print(f'科目 {len(rules)} 项。待核记录 {len(prepared[3])} 条。跨表科目差异 {len(prepared[4])} 项。{prepared[-1]}。')
    if args.预检:
        issues = Counter(reason for _,_,_,reason in prepared[3])
        for reason,count in issues.most_common(6):
            print(f'  {count} 条 · {reason}')
        for key,field,a,b,diff,note in prepared[2]:
            unit = 10**args.金额小数位
            print(f'  {key} {field}：{a/unit:,.{args.金额小数位}f} / {b/unit:,.{args.金额小数位}f}，差额 {diff/unit:,.{args.金额小数位}f}')
        print('预检通过，已知缺值按规则留痕，本次没有生成结果文件。')
    else:
        print(f'成果目录：{写结果(args,books,supplements,rules,names,prepared)}')


if __name__ == '__main__':
    try:
        main()
    except (输入错误, OSError, BadZipFile, InvalidFileException, json.JSONDecodeError) as exc:
        print(f'请确认：{exc}',file=sys.stderr)
        raise SystemExit(2)
