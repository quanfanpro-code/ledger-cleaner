"""识别 Excel 账表并统一账户组合和分列账表的读取结果。"""
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import math
import posixpath
import re
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.datetime import from_excel


class 输入错误(ValueError):
    pass


def 规范(value):
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().replace('帐','账')
    return ''.join(c for c in text if c.isalnum())


def 文本(value):
    return str(value).strip() if value is not None else ''


def 有值(value):
    return 文本(value) not in ('', '0', '-', '--', '—', '缺省', '默认', '#N/A', '#REF!', '#VALUE!')


字段 = {
    '组合': '账户组合|账户组合编码|账户编码组合|会计账户组合|账户编号|账户编码|科目组合|账户组合编号|account combination|code combination|account number',
    '描述': '账户描述|账户说明|账户名称|组合说明|组合描述|账户组合描述|account description|combination description',
    '科目编码': '科目|科目编码|会计科目|会计科目编码|科目编号|科目段编码|natural account',
    '科目级次': '科目级次|科目级别|科目层级|account level',
    '科目名称': '科目段|科目名称|会计科目名称|科目说明|natural account description',
    '一级编码': '一级编码|一级科目编码|一级科目编号',
    '一级科目': '一级科目|一级科目名称|一级科目说明',
    '日期': '有效日期|记账日期|凭证日期|会计日期|业务日期|日期|effective date|accounting date|journal date',
    '凭证号': '凭证号|凭证号码|凭证编号|凭证序号|journal number|voucher number',
    '摘要': '行摘要|分录摘要|摘要|分录说明|凭证摘要|line description|journal line description',
    '凭证类型': '凭证类型|凭证类别|凭证字|journal category|voucher type',
    '行号': '行号|分录行号|凭证行号|line number',
    '期间': '期间|会计期间|记账期间|会计月份|period|accounting period',
    '公司': '公司段|公司名称|记账公司|公司|company description',
    '公司编码': '公司编码|公司编号|公司代码|账套编码|账套编号|ledger code|company code',
    '部门': '部门段|部门名称|部门|成本中心名称|cost center description|department',
    '部门编码': '部门编码|部门编号|部门代码|成本中心编码|department code|cost center code',
    '子科目': '子科目段|子科目名称|子科目|subaccount description',
    '子科目编码': '子科目编码|子科目编号|subaccount code',
    '产品': '产品段|产品名称|product description',
    '产品编码': '产品编码|产品编号|product code',
    '项目': '项目段|项目名称|project description',
    '项目编码': '项目编码|项目编号|项目代码|project code',
    '往来': '往来段|往来单位|往来对象|往来名称|往来段名称|counterparty description',
    '往来编码': '往来编码|往来编号|往来对象编码|counterparty code',
    '客户': '客户名称|客户|customer name',
    '客户编码': '客户编码|客户编号|customer code',
    '供应商': '供应商名称|供应商|supplier name|vendor name',
    '供应商编码': '供应商编码|供应商编号|supplier code|vendor code',
    '个人': '职员姓名|员工姓名|个人姓名|employee name',
    '个人编码': '职员编号|员工编号|个人编码|employee code',
    '银行': '银行段|银行名称|开户行|开户银行|bank name',
    '银行编码': '银行编码|银行编号|bank code',
    '银行账号': '银行账号|银行账户号码|银行账户号|bank account number',
    '贷款银行': '贷款银行|贷款行|借款银行|贷款机构|lender name',
    '贷款银行编码': '贷款银行编码|贷款行编码|贷款机构编码|lender code',
    '合同编码': '借款合同编号|合同编号|合同编码|contract number|contract code',
    '借方': '本位币借方金额|本位币借方|本期借方|本期借方发生额|本期发生额借方|本年发生额借方|累计发生额借方|累计借方发生额|借方发生额|借方金额|借方|accounted debit|debit',
    '贷方': '本位币贷方金额|本位币贷方|本期贷方|本期贷方发生额|本期发生额贷方|本年发生额贷方|累计发生额贷方|累计贷方发生额|贷方发生额|贷方金额|贷方|accounted credit|credit',
    '期初': '期初|期初余额|期初净额|开始余额|opening balance|beginning balance',
    '期末': '期末|期末余额|期末净额|结束余额|ending balance|closing balance',
    '期初借方': '期初借方|期初余额借方|期初余额借方金额|期初借方余额|opening debit',
    '期初贷方': '期初贷方|期初余额贷方|期初余额贷方金额|期初贷方余额|opening credit',
    '期末借方': '期末借方|期末余额借方|期末余额借方金额|期末借方余额|closing debit',
    '期末贷方': '期末贷方|期末余额贷方|期末余额贷方金额|期末贷方余额|closing credit',
}
字段 = {key: {规范(x) for x in names.split('|')} for key, names in 字段.items()}
公共 = {'借方', '贷方'}
分列标识 = {'科目编码', '科目名称'}
名称段位 = {'公司':1, '部门':3, '科目名称':4, '子科目':5, '产品':6, '项目':7, '往来':8, '银行':9}
分列专用字段 = {
    '公司编码','部门编码','子科目编码','产品编码','项目编码','往来编码',
    '客户','客户编码','供应商','供应商编码','个人','个人编码',
    '银行编码','银行账号','贷款银行','贷款银行编码','合同编码',
}
日期级别 = {'有效日期': 3, '会计日期': 2, '记账日期': 2, '凭证日期': 1, '业务日期': -2, '日期': -1}
金额字段 = {'借方', '贷方', '期初', '期末', '期初借方', '期初贷方', '期末借方', '期末贷方'}
单位 = {'元', '本位币', '记账本位币', '人民币', '金额', '人民币元', 'cny', 'rmb', '单位元'}


@dataclass
class 位置:
    文件: str
    页: str
    行: int

    def __str__(self):
        return f'{self.文件} / {self.页} 第{self.行}行'


@dataclass
class 行:
    位置: 位置
    编码: tuple
    名称: dict
    金额: tuple
    日期: object = None
    凭证: str = ''
    摘要: str = ''
    类型: str = ''
    期间: str = ''
    年份: str = ''
    行号: str = ''
    原始组合: str = ''
    分列: bool = False
    贷款行: tuple = ('','')
    合同: str = ''
    往来冲突: str = ''
    银行账号: str = ''
    一级编码: str = ''
    一级科目: str = ''
    辅助指定: tuple | None = None
    行类型: str = ''
    末级: bool = True
    科目级次: int | None = None
    鼎信诺键: tuple = ()
    借贷方向: str = ''
    辅助记录: tuple = ()
    辅助分项: tuple = ()
    币种: str = ''


@dataclass
class 表:
    类别: str
    文件: str
    页: str
    表头: int
    字段: dict
    行: list
    合计: list
    期间: str
    警示: list
    识别方式: str
    补充核对: list | None = None
    源行数: int | None = None
    控制合计: tuple | None = None


def 工作表部件(path):
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        links = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        parts = {el.get('Id'): el.get('Target') for el in links}
    for el in workbook.iter():
        if el.tag.rsplit('}', 1)[-1] == 'sheet':
            key = next((v for k, v in el.attrib.items() if k.endswith('}id')), None)
            target = parts.get(key)
            if target:
                yield el.get('name'), (
                    target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
                )


def 读取布局(path, part):
    merges, empty_formulas = [], set()
    with ZipFile(path) as archive, archive.open(part) as stream:
        parents = []
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                parents.append(node)
                continue
            tag = node.tag.rsplit('}', 1)[-1]
            if tag == 'mergeCell':
                merges.append(range_boundaries(node.get('ref')))
            if tag == 'c' and any(child.tag.endswith('}f') for child in node):
                raw = next((child.text for child in node if child.tag.endswith('}v')), None)
                if raw is None:
                    empty_formulas.add(node.get('r'))
            # 每行读完就释放，避免大账表占用过多内存。
            if len(parents) > 1 and tag in ('row', 'mergeCell'):
                parents[-2].remove(node)
            parents.pop()
    return merges, empty_formulas


def 展开表头(raw, merges):
    height = len(raw)
    width = max((len(r) for r in raw), default=0)
    grid = [list(r) + [None] * (width - len(r)) for r in raw]
    for left, top, right, bottom in merges:
        if top > height or left > width:
            continue
        value = grid[top-1][left-1]
        for r in range(top-1, min(bottom, height)):
            for col in range(left-1, min(right, width)):
                grid[r][col] = value
    return grid


def 查看结构(path, row_limit):
    """列出原始单元格、合并范围及公式位置，供 Agent 建立版式方案。"""
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        parts = dict(工作表部件(path))
        result = []
        for sheet in book.worksheets:
            sheet.reset_dimensions()
            merges, formulas = 读取布局(path, parts[sheet.title])
            data = []
            missing_by_row = defaultdict(list)
            for cell in formulas:
                number = int(re.search(r'\d+$',cell)[0])
                if number <= row_limit:
                    missing_by_row[number].append(cell)
            for number, row in enumerate(sheet.iter_rows(max_row=row_limit, values_only=True), 1):
                cells = {get_column_letter(col):value for col,value in enumerate(row,1) if value is not None}
                missing = sorted(missing_by_row.get(number,()))
                if cells or missing:
                    data.append({'行号':number,'单元格':cells,'未计算公式':missing})
            result.append({
                '名称':sheet.title, '可见性':sheet.sheet_state, '样例行':data,
                '合并区域':[f'{get_column_letter(a)}{b}:{get_column_letter(c)}{d}'
                           for a,b,c,d in merges if b <= row_limit],
            })
        return {'文件':path.name,'工作表':result}
    finally:
        book.close()


def 列候选(kind, grid, extra):
    hits = defaultdict(list)
    for col, levels in enumerate(zip(*grid)):
        names = list(dict.fromkeys(v for v in levels if v))
        names = [val for val in names if val not in 单位 or val in ('本位币', '人民币')]
        if not names:
            continue
        meaningful = [n for n in names if n not in 单位]
        if not meaningful:
            continue
        leaf = meaningful[-1]
        joined = ''.join(meaningful)
        path = ''.join(names)
        if any(term in path for term in ('外币', '原币', '交易币', 'entered', 'foreigncurrency')):
            foreign = True
        else:
            foreign = False
        end = any(x in leaf or x in joined for x in ('期初', '期末', 'opening', 'closing', 'beginning', 'ending'))
        for key, builtins in 字段.items():
            targets = builtins | extra.get(key, set())
            score = 6 if leaf in targets else 0
            if joined in targets and joined != leaf:
                score = max(score, 7)
            if leaf.removesuffix('元') in targets:
                score = max(score, 5)
            if not score:
                continue
            if key in 金额字段:
                if foreign or (key in ('借方', '贷方') and end):
                    continue
                if key in ('期初', '期末') and leaf in 字段['借方'] | 字段['贷方']:
                    continue
                if key in ('借方', '贷方') and kind == '余额表' and '本年' not in path and '累计' not in path and '本月' in path:
                    continue
                if '本位币' in path or 'accounted' in path:
                    score += 4
                elif '人民币' in path:
                    score += 2
                if kind == '余额表' and any(n in path for n in ('本期', '本年', '累计')):
                    score += 1
            if key == '日期':
                score += 日期级别.get(leaf, 0)
            hits[key].append((score, col))
    return hits


def 必需字段(kind, fields):
    identity = {'组合'} if '组合' in fields else 分列标识
    required = 公共 | identity
    if kind == '序时账':
        return required | {'日期', '凭证号'}
    for base in ('期初', '期末'):
        paired = {base+'借方', base+'贷方'}
        required |= paired if base not in fields and paired <= fields else {base}
    return required


def 匹配表头(kind, grid, extra, limits):
    best = None
    identity_aliases = 字段['组合'] | 字段['科目编码'] | extra.get('组合',set()) | extra.get('科目编码',set())
    for end in range(1, min(len(grid), limits) + 1):
        for start in range(max(0, end-4), end):
            block = grid[start:end]
            if not any(v in identity_aliases for row in block for v in row):
                continue
            matches = 列候选(kind, block, extra)
            required = 必需字段(kind, matches.keys())
            if not required <= matches.keys():
                continue
            selected = {}
            for field, choices in matches.items():
                high = max(score for score, _ in choices)
                selected[field] = (high, [col for score, col in choices if score == high])
            relevant = (set(字段) - {'期初','期末','期初借方','期初贷方','期末借方','期末贷方'}) if kind == '序时账' else (
                set(字段) - {'日期','凭证号','期间','行号','摘要','凭证类型'}
            )
            score = (sum(selected[key][0] for key in required), len(relevant & selected.keys()), -end, start)
            if best is None or score > best[0]:
                best = (score, end, selected, required)
    return best


def 按方案定位(path, book, parts, kind, forced, layout):
    selected_kind, title, end = (layout[key] for key in ('类型','工作表','表头末行'))
    if (kind and kind != selected_kind) or (forced and forced != title):
        raise 输入错误(f'{path.name} 的版式方案和指定的表类型或工作表冲突。')
    if title not in parts or title not in {sheet.title for sheet in book.worksheets}:
        raise 输入错误(f'{path.name} 的版式方案指定了找不到的工作表：{title}。')
    sheet = book[title]
    sheet.reset_dimensions()
    columns = layout['列']
    required = 必需字段(selected_kind, columns.keys())
    if not required <= columns.keys():
        raise 输入错误(f'{path.name} 的版式方案缺少字段：{"、".join(sorted(required-columns.keys()))}。')
    raw = list(sheet.iter_rows(max_row=end, values_only=True))
    if end > len(raw) or not any(v is not None for v in raw[end-1]):
        raise 输入错误(f'{path.name} / {title} 的表头末行 {end} 超出有效范围或为空。')
    width = max(map(len, raw))
    if any(col >= width for col in columns.values()):
        raise 输入错误(f'{path.name} / {title} 的版式方案选了超出表头范围的列。')
    merges, formulas = 读取布局(path, parts[title])
    header = [[规范(v) for v in row] for row in 展开表头(raw,merges)[-4:]]
    recognized = 列候选(selected_kind,header,{})
    for field in 金额字段 & columns.keys():
        col = columns[field]
        headings = ''.join(row[col] for row in header)
        if any(token in headings for token in ('原币','外币','交易币','foreigncurrency','entered')):
            raise 输入错误(f'{path.name} / {title} 的{get_column_letter(col+1)}列标注为原币或外币，请核对本位币口径。')
        known = {key for key in 金额字段 if any(index == col for _,index in recognized.get(key,()))}
        if known and field not in known:
            raise 输入错误(f'{path.name} / {title} 的{get_column_letter(col+1)}列标注「{"、".join(sorted(known))}」，与版式方案的「{field}」冲突。')
    last = list(raw[-1]) + [None]*(width-len(raw[-1]))
    id_key = '组合' if '组合' in columns else '科目编码'
    id_value = last[columns[id_key]]
    if 有值(id_value) and 规范(id_value) not in 字段[id_key] and any(
        isinstance(last[columns[key]],(int,float)) for key in 金额字段 & columns.keys()
    ):
        raise 输入错误(f'{path.name} / {title} 第{end}行已有科目标识和金额数据，请核对表头末行，避免漏读。')
    matches = {key:(1,[index]) for key,index in columns.items()}
    return selected_kind, sheet, (None,end,matches,required), merges, formulas, raw


def 定位表(path, book, kind, forced, limit, extra, layout):
    parts = dict(工作表部件(path))
    if layout:
        return 按方案定位(path, book, parts, kind, forced, layout)
    candidates = []
    aliases = 字段['组合'] | 字段['科目编码'] | extra.get('组合',set()) | extra.get('科目编码',set())
    for sheet in book.worksheets:
        if forced and sheet.title != forced:
            continue
        part = parts.get(sheet.title)
        if not part:
            continue
        sheet.reset_dimensions()
        raw = list(sheet.iter_rows(min_row=1, max_row=limit, values_only=True))
        # 先找科目标识，避开封面和纯空表。
        if not any(规范(v) in aliases for row in raw for v in row):
            continue
        merges, formulas = 读取布局(path, part)
        grid = [[规范(v) for v in row] for row in 展开表头(raw, merges)]
        for option in ([kind] if kind else ['序时账', '余额表']):
            hit = 匹配表头(option, grid, extra, limit)
            if hit:
                candidates.append((option, sheet, hit, merges, formulas, raw))
    if not candidates:
        sheets = '、'.join(book.sheetnames)
        raise 输入错误(f'{path.name} 未找到完整的{kind or "序时账或余额表"}表头。'
                  f'已查看前{limit}行，工作表：{sheets}。请先查看结构，核对科目、借贷金额及凭证或期初期末，再提供版式方案。')
    if len(candidates) != 1:
        options = '；'.join(f'{sheet.title} 第{hit[1]}行（{option}）' for option, sheet, hit, *_ in candidates)
        raise 输入错误(f'{path.name} 有多个表或格式候选：{options}。请确认工作表和表类型。')
    return candidates[0]


组合格式 = re.compile(r'[0-9A-Za-z_/-]+(?:[.．][0-9A-Za-z_/-]+){9}\Z')


def 解析描述(description, parts, explicit):
    body = 文本(description)
    if not body:
        return {}
    pieces = re.split(r'(?<!\\)[.．]', body)
    names = {}
    active = [i for i,c in enumerate(parts) if i in (0, 2, 3, 4, 6) or c != '0']
    if len(pieces) == len(active) and len(active) < 10:
        for j, i in enumerate(active):
            end = active[j+1] if j+1 < len(active) else 10
            count = end-i-1
            suffix = '-' * count
            if count and not pieces[j].endswith(suffix):
                break
            value = pieces[j][:-count] if count else pieces[j]
            if parts[i] == '0' and 有值(value):
                break
            names[i+1] = value.strip().replace('\\.', '.')
        else:
            return names
    if len(pieces) > 10:
        # 旧导出在合同利率中也使用普通点号，只有这能还原整十段时才采用。
        decimal_safe = re.split(r'(?<!\\)(?<!\d)[.．]|(?<!\\)[.．](?!\d)', body)
        if len(decimal_safe) == 10 and all(
            有值(value) == (parts[i] != '0') for i,value in enumerate(decimal_safe) if i != 2
        ):
            pieces = decimal_safe
    if len(pieces) == 10:
        return {i+1: p.strip().replace('\\.', '.') for i,p in enumerate(pieces)}
    if not 10 < len(pieces) <= 35:
        return {}
    # 全点号描述可能把利率、小数或公司名称中的点写成普通点号。只采用能唯一分段的结果。
    results = []

    def search(seg, offset, record):
        if len(results) > 1:
            return
        if seg == 10:
            if offset == len(pieces):
                results.append(record)
            return
        end = len(pieces) - (9-seg)
        for stop in range(offset+1, end+1):
            name = '.'.join(pieces[offset:stop]).strip().replace('\\.', '.')
            if (parts[seg] == '0') != (not 有值(name)) and seg != 2:
                continue
            anchor = explicit.get(seg+1)
            if 有值(anchor) and 规范(name) != 规范(anchor):
                continue
            search(seg+1, stop, {**record, seg+1: name})

    search(0, 0, {})
    return results[0] if len(results) == 1 else {}


def 金额(value, position, column, precision):
    label = f'{position} {get_column_letter(column+1)}{position.行}'
    if value is None or (isinstance(value, str) and not value.strip()):
        return 0
    raw = unicodedata.normalize('NFKC', str(value)).strip().replace('−', '-')
    if raw.startswith('(') and raw.endswith(')'):
        raw = '-' + raw[1:-1].strip()
    if raw in ('-', '—', '–'):
        return 0
    if ',' in raw:
        if not re.fullmatch(r'[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?', raw):
            raise 输入错误(f'{label} 金额分组格式无法确认：{value}')
        raw = raw.replace(',', '')
    try:
        result = Decimal(raw)
        scaled = result.scaleb(precision)
        integer = scaled.to_integral_value()
    except InvalidOperation as exc:
        raise 输入错误(f'{label} 无法读取金额：{value}') from exc
    if not result.is_finite():
        raise 输入错误(f'{label} 无法读取金额：{value}')
    tolerance = Decimal(0)
    if isinstance(value, float):
        tolerance = Decimal(str(math.ulp(value) * 2)).scaleb(precision)
    if abs(scaled - integer) > tolerance:
        raise 输入错误(f'{label} 有超过{precision}位的有效小数，请确认金额精度：{value}')
    if abs(integer) >= 10**15:
        raise 输入错误(f'{label} 金额超出本次Excel精度范围。')
    return int(integer)


def 日期(value, epoch, loc):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = 文本(value)
    if isinstance(value, (int, float)) and 0 < value < 100000:
        interpreted = from_excel(value, epoch)
        if isinstance(interpreted, datetime):
            return interpreted.date()
    raw = re.sub(r'[年月./]', '-', raw).replace('日', '')
    try:
        if re.fullmatch(r'\d{8}', raw):
            return datetime.strptime(raw, '%Y%m%d').date()
        if re.fullmatch(r'\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?', raw):
            return date(*map(int, re.split(r'[- T]', raw, maxsplit=3)[:3]))
    except (ValueError, OverflowError):
        pass
    raise 输入错误(f'{loc} 日期无法识别：{value}')


def 标识(cell, loc):
    value = cell.value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if int(value) != value or abs(value) >= 10**15:
            raise 输入错误(f'{loc} 编号不是可精确读取的整数，请从源系统以文本导出。')
        fmt = cell.number_format or ''
        return str(int(value)).zfill(len(fmt)) if re.fullmatch(r'0{2,}', fmt) else str(int(value))
    return 文本(value)


def 解析科目级次(value, loc):
    if not 文本(value):
        return None
    match = re.fullmatch(r'(?:第)?([0-9]+)(?:级)?',unicodedata.normalize('NFKC',文本(value)))
    if not match or len(match[1]) > 6:
        raise 输入错误(f'{loc} 科目级次「{文本(value)}」无法识别，请核对原表。')
    return int(match[1])


def 分列往来(title, val, identifier, position):
    direct = (identifier('往来编码'),文本(val('往来')))
    groups = (('个人',('职工','职员','员工','个人')),
              ('供应商',('供应商','应付账款','预付账款')),
              ('客户',('客户','应收账款','预收账款','合同负债','合同资产')))
    matches = [role for role,words in groups if any(word in title for word in words)]
    evidence = {key:(identifier(key+'编码'),文本(val(key))) for key,_ in groups}
    if any(有值(v) for v in direct):
        if len(matches) == 1:
            other = evidence[matches[0]]
            if any(有值(a) and 有值(b) and 规范(a) != 规范(b) for a,b in zip(direct,other)):
                return '', '', f'{position} 的通用往来与「{matches[0]}」列冲突。请确认本次采用的列。'
            token,name = (a if 有值(a) else b for a,b in zip(direct,other))
            return token,name,''
        return direct[0],direct[1],''
    if len(matches) == 1:
        token,name = evidence[matches[0]]
        return token,name,''
    choices = {pair for pair in evidence.values() if any(有值(v) for v in pair)}
    if len(choices) > 1:
        return '', '', f'{position} 的科目「{title}」同时有多种往来对象。请根据科目和原始资料确认应采用哪列。'
    token,name = next(iter(choices), ('',''))
    return token,name,''


def 表头公司(raw):
    """只采纳标题中有明确标签的公司标识。"""
    codes, names = set(), set()
    for cell in (v for row in raw for v in row if isinstance(v,str)):
        code = re.fullmatch(r'\s*(?:公司|账套)(?:编码|编号|代码)\s*[:：]\s*([A-Za-z0-9_-]+)\s*',cell)
        name = re.fullmatch(r'\s*(?:编制单位|公司名称|核算主体)\s*[:：]\s*([^\n:：]+)\s*',cell)
        if code:
            codes.add(code[1])
        if name:
            names.add(name[1].strip())
    candidates = codes or names
    return next(iter(candidates)) if len(candidates) == 1 else ''


def 表头期间(top_text):
    for pattern in (
        r'((?:19|20)\d{2})年\s*(\d{1,2})月?\s*(?:至|到|[-—~～])\s*(?:((?:19|20)\d{2})年)?\s*(\d{1,2})月',
        r'((?:19|20)\d{2})[-/年](\d{1,2})月?\s*(?:至|到|[~～])\s*((?:19|20)\d{2})[-/年](\d{1,2})',
    ):
        match = re.search(pattern, top_text)
        if match:
            a,b,c,d = match.groups()
            if 1 <= int(b) <= 12 and 1 <= int(d) <= 12:
                return f'{a}-{int(b):02d}~{c or a}-{int(d):02d}'
    return ''


def 分列数据(val, identifier, columns, company, position):
    parts = ['0']*10
    names = {segment:文本(val(field)) for field,segment in 名称段位.items() if 有值(val(field))}
    for field,segment in 名称段位.items():
        code = identifier(field+'编码')
        if 有值(code):
            parts[segment-1] = code
    if parts[0] == '0':
        key = (names.get(1) or company) if '公司编码' not in columns else None
        if not 有值(key):
            raise 输入错误(f'{position} 无法确认公司范围。请核对公司列或报表标题，在版式方案中给出已确认的公司标识。')
        parts[0] = key
    parts[3] = identifier('科目编码')
    account_number = identifier('银行账号')
    if 有值(account_number):
        parts[4] = account_number
        bank = names.get(9) or names.get(5)
        if bank and 规范(bank).endswith(规范(account_number)):
            names[5] = bank
        else:
            names[5] = ' '.join(filter(None,(bank,account_number)))
    token, partner, conflict = 分列往来(names.get(4,''),val,identifier,position)
    parts[7] = token if 有值(token) else '0'
    if 有值(partner):
        names[8] = partner
    lender = (identifier('贷款银行编码'),文本(val('贷款银行')))
    return tuple(parts),names,lender,identifier('合同编码'),conflict


def 读表(path, kind=None, forced=None, limit=120, extra=None, assigned=None, precision=2, layout=None):
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        return _读表(path, book, kind, forced, limit, extra or {}, assigned or {}, precision, layout or {})
    finally:
        book.close()


def _读表(path, book, kind, forced, limit, extra, assigned, precision, layout):
    kind, sheet, hit, merges, formulas, raw = 定位表(path, book, kind, forced, limit, extra, layout)
    _, end, matches, required = hit
    if '组合' in required:
        matches = {key:value for key,value in matches.items() if key not in 分列专用字段}
    columns = {}
    assigned = assigned.get(kind, {})
    if layout and assigned:
        raise 输入错误(f'{path.name} 已提供完整版式方案，请在方案内改列位置。')
    for field,col in assigned.items():
        if field not in matches or not 0 <= col < max(map(len,raw)):
            raise 输入错误(f'{path.name} / {sheet.title}：「{field}」指定列超出表头范围或字段尚未识别。')
    for field, (_, choices) in matches.items():
        if field in assigned:
            columns[field] = assigned[field]
            continue
        if len(choices) != 1:
            if field in required or '组合' not in required or field in ('描述','科目名称','科目编码','公司','部门','子科目','项目','往来','银行','摘要','凭证类型','期间'):
                letters = '、'.join(f'{get_column_letter(c+1)}列' for c in choices)
                raise 输入错误(f'{path.name} / {sheet.title} 表头第{end}行，「{field}」匹配到{letters}，请确认使用哪列。')
            continue
        columns[field] = choices[0]
    if len(set(columns.values())) != len(columns):
        raise 输入错误(f'{path.name} / {sheet.title}：多个字段匹配到了同一列，请核对表头或指定列。')
    for base in ('期初', '期末'):
        if kind == '余额表' and base in columns and {base+'借方',base+'贷方'} <= columns.keys():
            raise 输入错误(f'{path.name} / {sheet.title} 同时有带符号{base}与{base}借贷两列，请确认使用哪组。')
    split_layout = '组合' not in columns
    identity_field = '科目编码' if split_layout else '组合'
    company = (layout.get('公司标识') or 表头公司(raw[:end])) if split_layout else ''
    width = max(max(columns.values())+1, min(12, max(map(len, raw))))
    rows, totals, warnings = [], [], []
    span_start = defaultdict(list)
    active = {}
    for left, top, right, bottom in merges:
        if top <= end:
            continue
        affected = [name for name,col in columns.items() if left <= col+1 <= right]
        if not affected:
            continue
        if left != right or (affected[0] in 金额字段 and top != bottom):
            raise 输入错误(f'{path.name} / {sheet.title} 第{top}行：数据区合并跨越{affected[0]}列，请拆分后确认原值。')
        if top != bottom:
            span_start[top].append((left-1,bottom))
    amount_keys = ['借方','贷方'] if kind == '序时账' else ['期初','借方','贷方','期末']
    used_amounts = [key for key in columns if key in 金额字段 and (kind == '余额表' or key in amount_keys)]
    top_text = ' '.join(文本(v) for row in raw[:end] for v in row if isinstance(v, str))
    amount_headers = ' '.join(文本(row[columns[key]]) for row in raw[:end]
                              for key in used_amounts if columns[key] < len(row))
    if re.search(r'单位\s*[:：]?\s*(?:人民币\s*)?[万千百]元', top_text) or re.search(r'[万千百]元', amount_headers):
        raise 输入错误(f'{path.name} / {sheet.title} 的报表金额单位不是元，请确认如何换算后再处理。')
    period = 表头期间(top_text) if kind == '余额表' else ''
    plan_period = layout.get('会计期间','')
    if period and plan_period and period != plan_period:
        raise 输入错误(f'{path.name} / {sheet.title} 的表头会计期间与版式方案不一致。')
    period = plan_period or period
    formula_watch = columns.keys() if split_layout else required | set(used_amounts) | ({'科目名称','描述','科目编码'} & columns.keys())
    for number, cells in enumerate(sheet.iter_rows(min_row=end+1, max_col=width), end+1):
        pos = 位置(path.name, sheet.title, number)
        values = [cell.value for cell in cells]
        for col,bottom in span_start.get(number, []):
            active[col] = (bottom, cells[col])
        for col,(bottom,cell) in tuple(active.items()):
            if number > bottom:
                del active[col]
            else:
                values[col] = cell.value
        def val(key):
            return values[columns[key]] if key in columns else None

        def identifier(key):
            col = columns.get(key)
            return 标识(active[col][1] if col in active else cells[col], pos) if col is not None else ''

        for key in formula_watch:
            col = columns[key]
            if f'{get_column_letter(col+1)}{number}' in formulas:
                raise 输入错误(f'{pos} {get_column_letter(col+1)}{number} 的{key}公式没有已保存的计算结果。请在Excel重新计算并保存源文件。')
        identity = identifier(identity_field)
        if 规范(identity) in 字段[identity_field] | extra.get(identity_field, set()):
            for key in required | set(used_amounts):
                col = columns[key]
                names = 字段[key] | extra.get(key, set())
                found = [i for i,cell in enumerate(values) if 规范(cell) in names]
                if found and col not in found:
                    raise 输入错误(f'{pos} 重复表头的「{key}」列发生变化，请确认是否分开清洗。')
            continue
        money_values = [val(key) for key in used_amounts]
        if not 有值(identity) and all(v is None or not 文本(v) or 规范(v) in 单位 for v in money_values):
            if val('日期') or val('凭证号'):
                raise 输入错误(f'{pos} 有凭证信息，「{identity_field}」为空。')
            continue

        def amount(key):
            if key in columns:
                return 金额(val(key), pos, columns[key], precision)
            return amount(key+'借方') - amount(key+'贷方')

        total_labels = r'(?:总计|合计|小计|累计|本页合计|本期合计|本年合计|grand total|total)\s*[:：]?'
        if not 有值(identity) or re.fullmatch(total_labels,identity,re.I) or (not split_layout and not 组合格式.fullmatch(identity)):
            rollup = next((文本(x) for x in values if re.fullmatch(total_labels, 文本(x), re.I)), '')
            if rollup and not (val('日期') or val('凭证号')):
                if '小计' not in rollup and '累' not in rollup and '本页' not in rollup:
                    totals.append((pos, tuple(amount(key) for key in amount_keys)))
                continue
            if split_layout:
                raise 输入错误(f'{pos} 有金额但无法确认科目编码：{identity or "空白"}。')
            raise 输入错误(f'{pos} 的账户组合无法识别，预期为10段编码：{identity or "空白"}。请先核对版式。')
        amounts = tuple(amount(key) for key in amount_keys)
        if split_layout:
            parts,names,lender,contract,conflict = 分列数据(val,identifier,columns,company,pos)
            row = 行(pos,parts,names,amounts,分列=True,贷款行=lender,合同=contract,
                    往来冲突=conflict,银行账号=identifier('银行账号'))
        else:
            parts = tuple(re.split('[.．]', identity))
            other_code = identifier('科目编码')
            if other_code and other_code != parts[3]:
                raise 输入错误(f'{pos} 科目列为 {other_code}，账户组合第4段为 {parts[3]}，请确认口径。')
            explicit = {segment:文本(val(field)) for field,segment in 名称段位.items() if 有值(val(field))}
            names = 解析描述(val('描述'), parts, explicit)
            for seg,name in explicit.items():
                alternative = names.get(seg)
                if 有值(alternative) and 规范(alternative) != 规范(name):
                    warnings.append((pos, '分列名称与组合说明不同', f'第{seg}段采用分列：{name}，描述：{alternative}'))
                names[seg] = name
            row = 行(pos,parts,names,amounts,原始组合=identity)
        if kind == '序时账':
            row.日期 = 日期(val('日期'), book.epoch, pos)
            row.凭证 = identifier('凭证号')
            if not 有值(row.凭证):
                raise 输入错误(f'{pos} 凭证号为空或无效。')
            row.摘要 = 文本(val('摘要'))
            row.类型 = 文本(val('凭证类型'))
            row.期间 = 文本(val('期间')) or row.日期.strftime('%Y-%m')
            row.年份 = str(row.日期.year)
            row.行号 = 文本(val('行号'))
            row.一级编码 = identifier('一级编码')
            row.一级科目 = 文本(val('一级科目'))
        else:
            row.年份 = period[:4] if period and period[:4] == period.split('~')[1][:4] else ''
            row.科目级次 = 解析科目级次(identifier('科目级次'),pos)
        rows.append(row)
    if not rows:
        raise 输入错误(f'{path.name} / {sheet.title} 找到表头，但没有可转换的账户明细。')
    if kind == '序时账':
        periods = sorted({row.日期.strftime('%Y-%m') for row in rows})
        period = f'{periods[0]}~{periods[-1]}'
    method = '版式方案' if layout else ('分列自动识别' if split_layout else '账户组合自动识别')
    return 表(kind, path.name, sheet.title, end, columns, rows, totals, period, warnings, method)
