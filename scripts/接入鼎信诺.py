"""把鼎信诺主表、辅助明细和辅助余额转换为统一账表。"""
import csv
import re
from collections import Counter, defaultdict
from contextlib import contextmanager
from dataclasses import replace
from itertools import islice
from types import SimpleNamespace

from openpyxl import load_workbook
from openpyxl.utils.datetime import WINDOWS_EPOCH
from openpyxl.utils.cell import get_column_letter

from 读取账表 import (
    位置, 展开表头, 工作表部件, 日期, 有值, 标识, 规范, 表, 表头公司, 表头期间, 解析科目级次,
    行, 读取布局, 输入错误, 金额, 文本, 字段, 币种表头, 匹配表头, 分列数据, 余额方向, 核实分录日期,
)


名称映射 = {
    '科目编码': '科目编号|科目编码|会计科目编码',
    '科目名称': '科目名称|会计科目名称',
    '日期': '记账时间|记账日期|凭证日期|日期',
    '凭证号': '凭证编号|凭证号',
    '凭证类型': '凭证种类|凭证类型|凭证类别',
    '行号': '编号|分录行号|凭证行号',
    '摘要': '业务说明|摘要|分录摘要',
    '借方': '借方发生额|本位币借方发生额|借方金额',
    '贷方': '贷方发生额|本位币贷方发生额|贷方金额',
    '账面借方': '账面借方发生额|账面本期借方发生额|账面累计借方发生额|账面借方',
    '账面贷方': '账面贷方发生额|账面本期贷方发生额|账面累计贷方发生额|账面贷方',
    '期初': '账面期初数|账面期初余额|账面期初|期初余额|期初数',
    '期末': '账面期末数|账面期末余额|账面期末|期末余额|期末数',
    '方向': '借贷方向|余额方向',
    '末级': '是否明细科目|是否末级|末级科目',
    '级次': '科目级次|科目级别',
    '一级编码': '一级编码|一级科目编码|一级科目编号',
    '一级科目': '一级科目|一级科目名称',
    '公司': '公司编码|公司名称|核算主体|账套编号|账套编码',
    '年': '会计年|会计年度',
    '月': '会计月|会计月份',
    '币种': '货币种类|币种',
    '辅助维度': '核算项目类型名称|核算项目类别|核算项目类型|辅助核算类型名称|辅助类别',
    '辅助名称': '核算项目名称|辅助项目名称|辅助核算名称',
    '辅助编码': '核算项目编号|核算项目编码|辅助项目编码|辅助核算编码',
    '合并辅助': '辅助核算信息',
}
名称映射 = {key: {规范(v) for v in names.split('|')} for key,names in 名称映射.items()}
for key in ('日期','凭证号','凭证类型','行号','摘要','借方','贷方','科目编码','科目名称',
            '银行','银行账号','银行编码','贷款银行','贷款银行编码','合同编码','客户','客户编码',
            '供应商','供应商编码','个人','个人编码','往来','往来编码','部门','部门编码',
            '项目','项目编码','子科目','子科目编码','产品','产品编码'):
    名称映射.setdefault(key,set()).update(字段[key])
名称映射['公司'] = (名称映射['公司'] | 字段['公司']) - 字段['公司编码']
名称映射['公司编码'] = 字段['公司编码']
名称映射['级次'].update(字段['科目级次'])
源标记 = {'记账时间','凭证种类','业务说明','是否明细科目','科目级次','辅助核算信息'}
源标记.update(名称映射['辅助维度'])
金额字段 = {'期初','期末','借方','贷方','账面借方','账面贷方'}
总计标签 = {'合计','总计','小计','本页合计','本期合计','本年合计'}
辅助表 = {'核算项目明细表':'序时账', '核算项目余额表':'余额表'}
未列明 = '未列明辅助项目'
核对列 = ['情况','主表位置','辅助来源','科目编码','源维度','辅助编码','辅助名称',
        '原借贷方向','期初净额','借方发生额','贷方发生额','期末净额']


@contextmanager
def 打开源(path):
    with path.open('rb') as stream:
        signature = stream.read(8)
    if signature.startswith(b'PK'):
        with path.open('rb') as stream:
            book = load_workbook(stream, read_only=True, data_only=True, keep_links=False)
            try:
                parts = dict(工作表部件(path))
                for sheet in book.worksheets:
                    sheet.reset_dimensions()
                yield [
                    (sheet.title, sheet.sheet_state, sheet.iter_rows,
                     lambda name=sheet.title: 读取布局(path, parts[name]), book.epoch)
                    for sheet in book.worksheets
                ]
            finally:
                book.close()
        return
    if signature == bytes.fromhex('d0cf11e0a1b11ae1'):
        try:
            import xlrd
        except ImportError as exc:
            raise 输入错误('读取旧版 Excel 需要技能清单中的旧表读取组件。') from exc
        try:
            book = xlrd.open_workbook(str(path), formatting_info=True, on_demand=True)
        except (xlrd.XLRDError, ValueError) as exc:
            raise 输入错误(f'{path.name} 的旧版 Excel 工作簿无法读取：{exc}') from exc

        def rows(sheet, min_row=1, max_row=None):
            for r in range(min_row-1, min(max_row or sheet.nrows, sheet.nrows)):
                cells = []
                for cell in sheet.row(r):
                    value = cell.value
                    pattern = ''
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        value = None
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        value = xlrd.error_text_from_code.get(value, '#ERROR')
                    elif cell.ctype == xlrd.XL_CELL_NUMBER and cell.xf_index is not None:
                        fmt = book.format_map.get(book.xf_list[cell.xf_index].format_key)
                        pattern = fmt.format_str if fmt else ''
                    cells.append(SimpleNamespace(value=value, number_format=pattern))
                yield cells

        try:
            sheets = book.sheets()
            yield [
                (sheet.name, 'hidden' if sheet.visibility else 'visible', lambda s=sheet, **kw: rows(s, **kw),
                 lambda s=sheet: ([(c+1,a+1,d,b) for a,b,c,d in s.merged_cells], set()), WINDOWS_EPOCH)
                for sheet in sheets
            ]
        finally:
            book.release_resources()
        return
    if path.suffix.casefold() not in ('.xls','.tsv','.txt'):
        raise 输入错误(f'{path.name} 不是可读的 Excel 工作簿。')
    encoding = ('utf-16' if signature[:2] in (b'\xff\xfe', b'\xfe\xff') else 'utf-8-sig')
    for candidate in (encoding, 'gb18030'):
        try:
            with path.open(encoding=candidate, newline='') as stream:
                preview = stream.read(65536)
        except UnicodeDecodeError:
            continue
        if '\t' in preview and '\x00' not in preview:
            encoding = candidate
            break
    else:
        raise 输入错误(f'{path.name} 不是可读的 Excel 或制表符账表。')

    def text_rows(min_row=1, max_row=None):
        try:
            with path.open(encoding=encoding, newline='') as stream:
                reader = csv.reader(stream, delimiter='\t', strict=True)
                for row in islice(reader, min_row-1, max_row):
                    yield [SimpleNamespace(value=value, number_format='') for value in row]
        except (UnicodeError, csv.Error) as exc:
            raise 输入错误(f'{path.name} 的文本账表读到无效编码或引号：{exc}') from exc

    yield [('导出数据','visible',text_rows,lambda: ([],set()),WINDOWS_EPOCH)]


def 原始预览(rows, limit):
    return [[cell.value for cell in row] for row in rows(max_row=limit)]


def 查看旧表(path, limit):
    with 打开源(path) as sheets:
        result = []
        for title, visibility, rows, layout, _ in sheets:
            merges, _ = layout()
            sample = []
            for number, row in enumerate(原始预览(rows, limit), 1):
                cells = {get_column_letter(i): value for i,value in enumerate(row,1) if 文本(value)}
                if cells:
                    sample.append({'行号': number, '单元格': cells, '未计算公式': []})
            result.append({'名称': title, '可见性': visibility, '样例行': sample, '合并区域': [
                f'{get_column_letter(a)}{b}:{get_column_letter(c)}{d}' for a,b,c,d in merges if b <= limit]})
    return {'文件': path.name, '工作表': result}


def 定位专用表头(grid, loc, aliases, assigned):
    headers = []
    for end, bottom in enumerate(grid, 1):
        if not any(bottom):
            continue
        found = defaultdict(list)
        for col, value in enumerate(bottom):
            lower = 规范(value)
            terms = {lower}
            upper = ''
            if end >= 2 and 有值(value):
                upper = 规范(grid[end-2][col])
                if upper != lower:
                    terms.add(upper+lower)
            for field, labels in aliases.items():
                if not terms & labels:
                    continue
                priority = 0
                if field in 金额字段:
                    currency = 币种表头(grid,end,col)
                    if currency == '原币':
                        continue
                    priority = 2*int(currency == '本位币') + int('账面' in upper+lower)
                    if field in ('借方','贷方') and any(s in upper+lower for s in ('期初','期末')):
                        continue
                found[field].append((priority,col))
        fields = found.keys()
        if not {'科目编码','科目名称'} <= fields:
            continue
        has_aux = {'辅助维度','辅助名称'} <= fields
        if {'日期','凭证号','借方','贷方'} <= fields and (
            has_aux or '合并辅助' in fields or {规范(v) for v in bottom} & {'记账时间','凭证种类','业务说明'}
        ):
            kind = '核算项目明细表' if has_aux else '序时账'
            used = fields - {'期初','期末','方向','末级','级次','账面借方','账面贷方'}
        elif {'期初','期末','方向'} <= fields and (
            has_aux or '账面借方' in fields or '合并辅助' in fields or {'级次','末级'} & fields
        ):
            kind = '核算项目余额表' if has_aux else '余额表'
            if '账面借方' in fields and '账面贷方' in fields:
                found['借方'], found['贷方'] = found['账面借方'], found['账面贷方']
            if not {'借方','贷方'} <= found.keys():
                continue
            used = found.keys() - {'账面借方','账面贷方','日期','凭证号','摘要','凭证类型','行号','年','月'}
        else:
            continue
        score = len(used)
        columns = {f:[col for priority,col in found[f] if priority == max(found[f])[0]] for f in used}
        headers.append((score,end,kind,columns))
    if not headers:
        return None
    _,end,kind,found = max(headers, key=lambda item: (item[0], -item[1]))
    for field,col in assigned.get(辅助表.get(kind,kind),{}).items():
        field = '级次' if field == '科目级次' else field
        if col not in found.get(field,()):
            raise 输入错误(f'{loc} 的「{field}」指定列不是该字段可用的同口径候选列。')
        found[field] = [col]
    for field, choices in found.items():
        if len(choices) > 1:
            letters = '、'.join(get_column_letter(col+1) for col in choices)
            raise 输入错误(f'{loc} 表头第{end}行的「{field}」有多个候选列：{letters}。')
    columns = {key: choices[0] for key,choices in found.items()}
    if len(set(columns.values())) != len(columns):
        raise 输入错误(f'{loc} 表头第{end}行多个字段使用同一列。')
    return end,kind,columns


def 读鼎信诺(path, forced, limit, precision, extra, assigned):
    aliases = {k:labels | extra.get('科目级次' if k=='级次' else k,set()) for k,labels in 名称映射.items()}
    with 打开源(path) as sheets:
        candidates = []
        general = []
        for title, visibility, rows, layout, epoch in sheets:
            if (forced and title != forced) or (not forced and visibility != 'visible'):
                continue
            raw = 原始预览(rows, limit)
            tokens = {规范(value) for row in raw for value in row}
            if not tokens & (aliases['科目编码'] | 字段['组合'] | extra.get('组合',set())):
                continue
            merges, formulas = layout()
            grid = 展开表头(raw,merges)
            hit = None
            if (tokens & 源标记 or any(v.startswith('账面') for v in tokens)) and not tokens & 字段['组合']:
                hit = 定位专用表头(grid, f'{path.name} / {title}',aliases,assigned)
            if hit:
                candidates.append((title,rows,raw,formulas,merges,epoch,hit))
            elif len(sheets)>1 and not forced:
                normalized = [[规范(v) for v in row] for row in grid]
                if any(匹配表头(k,normalized,extra,limit) for k in ('序时账','余额表')):
                    general.append(title)
        if not candidates:
            return None
        if len(candidates) + len(general) > 1:
            pages = '、'.join([item[0] for item in candidates]+general)
            raise 输入错误(f'{path.name} 有多个可读取的工作表：{pages}。请指定工作表或分开提交。')
        book = 转换源表(path, *candidates[0], precision,aliases)
        if book.类别 == '余额表' and '级次' not in book.字段 and '合并辅助' in book.字段:
            for title, _, rows, _, _ in sheets:
                if title == '核对明细':
                    读取旧表级次(book,rows)
                    break
        return book


def 读取旧表级次(book, rows):
    iterator = iter(rows())
    header = [规范(cell.value) for cell in next(iterator,())]
    if not {'主表行号','科目级次'} <= set(header):
        return
    row_col,level_col = (header.index(field) for field in ('主表行号','科目级次'))
    controls = {row.位置.行:row for row in book.行 if not row.辅助记录}
    for number,cells in enumerate(iterator,2):
        if level_col >= len(cells) or not 文本(cells[level_col].value):
            continue
        pos = 位置(book.文件,'核对明细',number)
        reference = 标识(cells[row_col],pos) if row_col < len(cells) else ''
        row = controls.get(int(reference)) if reference.isdecimal() else None
        if row is None or row.科目级次 is not None:
            raise 输入错误(f'{pos} 的科目级次无法唯一对应原科目行。')
        row.科目级次 = 解析科目级次(标识(cells[level_col],pos),pos)


def 转换源表(path, title, rows, raw, formulas, merges, epoch, hit, precision,aliases):
    end, kind, columns = hit
    is_journal = kind in ('序时账','核算项目明细表')
    is_aux = kind in 辅助表
    prefix = f'{path.name} / {title}'
    tops = ' '.join(文本(v) for row in raw[:end] for v in row)
    grid = 展开表头(raw[:end],merges)
    amount_headers = ' '.join(文本(row[col]) for row in grid for key,col in columns.items() if key in 金额字段)
    if re.search(r'单位\s*[:：]?\s*(?:人民币\s*)?[万千百]元', tops) or re.search(r'[万千百]元', amount_headers):
        raise 输入错误(f'{prefix} 的金额单位不是元，请确认换算口径。')
    company = 表头公司(raw[:end])
    period = 表头期间(tops)
    result, warnings, totals = [], [], []
    slots = ['借方','贷方'] if is_journal else ['期初','借方','贷方','期末']
    vertical = defaultdict(list)
    active = {}
    for left,top,right,bottom in merges:
        if top <= end or not any(left <= col+1 <= right for col in columns.values()):
            continue
        if left != right or any(col == left-1 and key in 金额字段 for key,col in columns.items()):
            raise 输入错误(f'{prefix} 第{top}行的合并区域跨越字段或金额，请先核对。')
        if bottom > top:
            vertical[top].append((left-1,bottom))
    for number, cells in enumerate(rows(min_row=end+1), end+1):
        pos = 位置(path.name,title,number)
        for col, bottom in vertical.get(number, ()):
            if col < len(cells):
                active[col] = (bottom,cells[col])
        for col,(bottom,_) in tuple(active.items()):
            if number > bottom:
                del active[col]

        def cell_for(key):
            col = columns.get(key,-1)
            if col in active:
                return active[col][1]
            return cells[col] if 0 <= col < len(cells) else None

        def value(key):
            cell = cell_for(key)
            return cell.value if cell is not None else None

        def identifier(key):
            cell = cell_for(key)
            return 标识(cell,pos) if cell is not None else ''

        for key,col in columns.items():
            if f'{get_column_letter(col+1)}{number}' in formulas:
                raise 输入错误(f'{pos}「{key}」公式没有已保存的结果，请在 Excel 重新计算并保存。')
        code = identifier('科目编码')
        if 规范(code) in aliases['科目编码']:
            for key, col in columns.items():
                expected = aliases[key] | aliases.get('账面'+key,set())
                if 规范(value(key)) not in expected:
                    raise 输入错误(f'{pos} 重复表头的「{key}」换了位置，请分开清洗。')
            continue
        if not 有值(code) or 规范(code) in 总计标签:
            labels = {规范(c.value) for c in cells} & 总计标签
            if labels:
                if not is_aux and labels & {'总计','合计'}:
                    total = tuple(金额(value(k),pos,columns[k],precision) for k in slots)
                    if not is_journal and 有值(value('方向')):
                        _,total = 余额方向(value('方向'),total,pos)
                    elif not is_journal:
                        total = (None,total[1],total[2],None)
                    totals.append((pos,total))
                continue
            checked = ['科目名称','凭证号','摘要','日期','辅助名称']
            if any(有值(value(k)) for k in checked) or any(金额(value(k),pos,columns[k],precision) for k in slots):
                raise 输入错误(f'{pos} 有业务内容却缺少科目编码。')
            continue
        name = 文本(value('科目名称'))
        if not 有值(name):
            raise 输入错误(f'{pos} 的科目 {code} 缺少名称。')
        signed = tuple(金额(value(k),pos,columns[k],precision) for k in slots)
        company_cell = identifier('公司编码') or identifier('公司')
        if company and company_cell and 规范(company) != 规范(company_cell):
            raise 输入错误(f'{pos} 的公司标识和表头不一致。')
        company_id = company_cell or company
        parts,names,lender,contract,conflict = 分列数据(value,identifier,columns,company_id or '本批账套',pos)
        row = 行(pos,parts,names,signed,分列=True,贷款行=lender,合同=contract,往来冲突=conflict,银行账号=identifier('银行账号'))
        row.一级编码 = identifier('一级编码')
        row.一级科目 = 文本(value('一级科目'))
        row.币种 = 文本(value('币种'))
        dim, aux_code, aux_name = 文本(value('辅助维度')),identifier('辅助编码'),文本(value('辅助名称'))
        joined = 文本(value('合并辅助'))
        if joined == 未列明 and not is_journal:
            dim,aux_name = 未列明,未列明
        elif joined:
            parsed = re.split('[:：]',joined,maxsplit=1)
            if len(parsed) != 2 or not 有值(parsed[0]):
                raise 输入错误(f'{pos} 的合并辅助信息没有可识别的维度。')
            dim,aux_name = parsed
            if not 有值(aux_name):
                warnings.append((pos,'历史辅助项目为空','沿用源行金额，辅助项目待核'))
                if not is_journal:
                    dim,aux_name = 未列明,未列明
        if is_aux and not (有值(dim) and 有值(aux_name)):
            raise 输入错误(f'{pos} 缺少核算项目类型或名称。')
        if 有值(dim) and 有值(aux_name):
            if 有值(row.银行账号) and 维度来源(dim)==5:
                aux_code = row.银行账号
                bank = 文本(value('银行')) or aux_name
                aux_name = bank if 规范(bank).endswith(规范(aux_code)) else f'{bank} {aux_code}'
            row.辅助记录 = (dim.strip(),aux_code,aux_name.strip())
        if is_journal:
            row.日期 = 日期(value('日期'),epoch,pos)
            row.类型 = 文本(value('凭证类型'))
            row.凭证 = identifier('凭证号')
            row.行号 = identifier('行号')
            row.摘要 = 文本(value('摘要'))
            row.期间 = row.日期.strftime('%Y-%m')
            row.年份 = str(row.日期.year)
            if not 有值(row.凭证):
                raise 输入错误(f'{pos} 的凭证号为空。')
            for field, expected in [('年',row.日期.year),('月',row.日期.month)]:
                observed = identifier(field)
                if observed and (not observed.isdecimal() or int(observed) != expected):
                    raise 输入错误(f'{pos} 记账日期和会计{field}不同，请确认所属期间。')
        else:
            row.借贷方向,row.金额 = 余额方向(value('方向'),signed,pos)
            leaf = 规范(value('末级'))
            if leaf and leaf not in ('是','否','1','0','true','false','yes','no','y','n'):
                raise 输入错误(f'{pos} 无法识别是否明细科目。')
            row.末级 = leaf not in ('否','0','false','no','n')
            if not is_aux and not row.辅助记录:
                row.科目级次 = 解析科目级次(identifier('级次'),pos)
            if row.科目级次 == 1 or (row.科目级次 is None and re.fullmatch(r'\d{4}',code)):
                row.一级编码, row.一级科目 = code, name
            row.行类型 = '辅助明细' if row.辅助记录 else '科目汇总'
            row.年份 = period[:4] if period and period[:4] == period[8:12] else ''
        result.append(row)
    if not result and not is_aux:
        raise 输入错误(f'{prefix} 找到表头，但没有可转换的业务明细。')
    if is_journal:
        核实分录日期(result,period)
    method = f'鼎信诺{kind}'
    return 表(kind,path.name,title,end,columns,result,totals,period,warnings,method,[],len(result))


def 汇总(rows, size):
    totals = [0]*size
    for row in rows:
        for i,value in enumerate(row.金额):
            totals[i] += value
    return tuple(totals)


def 维度来源(dimension):
    label = 规范(dimension)
    if label in ('账户','账号','银行','银行账户','银行账号','存款账户','开户行','结算账户'):
        return 5
    for source, labels in (
        (7,('项目','工程','合同','研发课题')),
        (3,('部门','门店','店铺','成本中心')),
        (6,('商品','产品','存货','物料')),
        (8,('往来','客商','客户','供应商','职员','员工','个人','银行','贷款行','单位','股东')),
    ):
        if any(x in label for x in labels):
            return source
    return None


def 采用维度(row, rule, title):
    dimension,_,name = row.辅助记录
    source = rule[1]
    actual = 维度来源(dimension)
    if source in ('loan','interest'):
        return actual in (5,8) and any(s in name for s in ('银行','信用社','信用合作','财务公司'))
    if isinstance(source,str) and source.startswith('鼎信诺:'):
        return 规范(dimension) == 规范(source.split(':',1)[1])
    if source == 8 and actual == 8:
        label = 规范(dimension)
        if title.startswith(('应收账款','应收票据','预收','合同负债')) and '供应商' in label:
            return False
        if title.startswith(('应付账款','应付票据','预付')) and '客户' in label:
            return False
    return source == actual or (source == 9 and actual == 5)


def 对应公司(reference, extra):
    if not extra.行:
        return
    main = {r.编码[0] for r in reference.行} - {'本批账套'}
    detail = {r.编码[0] for r in extra.行} - {'本批账套'}
    if main and detail and not detail <= main:
        raise 输入错误(f'{reference.文件} 与 {extra.文件} 的公司或账套标识不同。')
    if extra.类别 == '核算项目余额表' and (not reference.期间 or not extra.期间):
        raise 输入错误(f'{reference.文件} 与 {extra.文件} 需要核实余额期间。')
    if reference.期间 and extra.期间 and reference.期间 != extra.期间:
        raise 输入错误(f'{reference.文件} 与 {extra.文件} 的起止期间不同。')
    if any(r.编码[0] == '本批账套' for rows in (reference.行,extra.行) for r in rows):
        raise 输入错误(f'{reference.文件} 与 {extra.文件} 部分行缺少公司标识，请根据来源确认本批账套。')


def 记录(reference, status, base, item):
    dim, code, name = item.辅助记录 or ('','','')
    values = item.金额 if len(item.金额) == 4 else (None,*item.金额,None)
    reference.补充核对.append((status,str(base.位置),str(item.位置),item.编码[3],
                               dim,code,name,item.借贷方向 or None,*values))


def 有据分拆(amount, parts):
    return all(
        all(v == 0 or (expected != 0 and (v > 0) == (expected > 0)) for v in col)
        and abs(sum(col)) <= abs(expected)
        for expected,col in zip(amount,zip(*(r.金额 for r in parts)),strict=True)
    )


def 匹配流水(book, extra):
    attached = defaultdict(list)
    if not extra:
        return attached
    对应公司(book,extra)
    keys = defaultdict(list)
    has_type = all(r.类型 for r in book.行) and all(r.类型 for r in extra.行)
    has_line = all(r.行号 for r in book.行) and all(r.行号 for r in extra.行)

    def identity(row):
        return row.编码[0],row.日期,row.凭证,规范(row.类型) if has_type else '',row.行号 if has_line else '',row.编码[3]

    for row in book.行:
        keys[identity(row)].append(row)
    pending = []
    for item in extra.行:
        choices = [row for row in keys.get(identity(item),[]) if 规范(row.名称[4]) == 规范(item.名称[4])]
        if not has_type and item.类型:
            choices = [row for row in choices if not row.类型 or 规范(row.类型) == 规范(item.类型)]
        if not has_line and item.行号:
            choices = [row for row in choices if not row.行号 or row.行号 == item.行号]
        if item.币种:
            choices = [row for row in choices if not row.币种 or 规范(row.币种) == 规范(item.币种)]
        if len(choices)>1 and item.摘要:
            exact = [r for r in choices if r.摘要.strip() == item.摘要.strip()]
            if exact:
                choices = exact
        if not choices:
            raise 输入错误(f'{item.位置} 找不到同一凭证、科目和分录的序时账来源。')
        if len(choices)==1:
            attached[id(choices[0])].append(item)
        else:
            pending.append((item,choices))
    # 先归入能定位的分录，再用同一维度已分配的金额排除不可能的行。
    while pending:
        rest = []
        for item,choices in pending:
            dim = 规范(item.辅助记录[0])
            possible = []
            for row in choices:
                group = [x for x in attached[id(row)] if 规范(x.辅助记录[0])==dim]
                if 有据分拆(row.金额,[*group,item]):
                    possible.append(row)
            if len(possible)==1:
                attached[id(possible[0])].append(item)
            elif possible:
                rest.append((item,possible))
            else:
                raise 输入错误(f'{item.位置} 辅助金额无法对应同一凭证下的分录。')
        if len(rest)==len(pending):
            unresolved = defaultdict(list)
            for item,choices in rest:
                unresolved[(identity(item),规范(item.辅助记录[0]))].append((item,choices))
            for (_,dim), entries in unresolved.items():
                candidates = {id(r):r for _,choices in entries for r in choices}
                for col in (0,1):
                    for sign in (1,-1):
                        available = sum(max(0,sign*r.金额[col]) for r in candidates.values())
                        assigned = sum(max(0,sign*x.金额[col]) for k in candidates for x in attached[k]
                                       if 规范(x.辅助记录[0])==dim)
                        requested = sum(max(0,sign*x.金额[col]) for x,_ in entries)
                        if assigned+requested>available:
                            raise 输入错误(f'{entries[0][0].位置} 同一凭证的辅助金额超过可对应分录的金额。')
            for item,choices in rest:
                记录(book,'原表未区分分录，保留凭证辅助金额',choices[0],item)
                record = book.补充核对[-1]
                book.补充核对[-1] = (record[0],'；'.join(str(r.位置) for r in choices),*record[2:])
                book.警示.append((item.位置,'原表未区分分录','原分录分别保留；该对象金额已保存在辅助匹配，仅归属到凭证'))
            break
        pending = rest
    return attached


def 分组余额(book, extra):
    details = defaultdict(list)
    controls = {}
    if extra:
        对应公司(book,extra)
    for row in book.行:
        key = row.编码[0],row.编码[3]
        if row.辅助记录:
            details[key].append(row)
        elif key in controls:
            raise 输入错误(f'{row.位置} 的科目 {key[1]} 有重复汇总行，请确认期间及范围。')
        else:
            controls[key] = row
    if extra:
        supplemental = defaultdict(list)
        for row in extra.行:
            supplemental[(row.编码[0],row.编码[3])].append(row)
        for key, additional in supplemental.items():
            same = {规范(row.辅助记录[0]) for row in additional}
            overlap = [row for row in details[key] if 规范(row.辅助记录[0]) in same]

            def signature(row):
                dim,_,name = row.辅助记录
                return 规范(dim),规范(name),row.金额,row.借贷方向

            if overlap:
                if key not in controls:
                    raise 输入错误(f'{additional[0].位置} 重复的核算项目没有对应科目汇总。')
                affected = {规范(row.辅助记录[0]) for row in overlap}
                duplicate = [row for row in additional if 规范(row.辅助记录[0]) in affected]
                if Counter(map(signature,overlap)) != Counter(map(signature,duplicate)):
                    raise 输入错误(f'{additional[0].位置} 与主表已带的辅助余额重叠且金额或对象不同，请确认使用哪份。')
                for row in overlap:
                    记录(book,'重复的旧辅助行，采用配套表',controls[key],row)
                details[key] = [row for row in details[key] if 规范(row.辅助记录[0]) not in affected]
            details[key].extend(additional)
    标记上级(controls.values())
    for key, items in details.items():
        base = controls.get(key)
        if not base or not base.末级:
            raise 输入错误(f'{items[0].位置} 的核算项目没有同公司同科目的末级汇总。')
        for row in items:
            if 规范(row.名称[4]) != 规范(base.名称[4]):
                raise 输入错误(f'{row.位置} 的科目名称与 {base.位置} 不一致。')
    return controls,details


def 标记上级(rows):
    controls = {(row.编码[0],row.编码[3]):row for row in rows if not row.辅助记录 and row.行类型!='辅助明细'}
    for key in 直接下级(controls):
        controls[key].末级 = False


def 直接下级(controls):
    children = defaultdict(list)
    for (company,code),row in controls.items():
        for end in range(len(code)-1,0,-1):
            if not code.isdecimal() and code[end] not in '.-':
                continue
            key = company,code[:end]
            parent = controls.get(key)
            if parent is None:
                continue
            if row.科目级次 is not None and parent.科目级次 is not None and parent.科目级次 >= row.科目级次:
                continue
            children[key].append(row)
            break
    return children


def 核对上级(book, controls):
    by_parent = 直接下级(controls)
    for key,parent in controls.items():
        children = by_parent.get(key,())
        if not children and not parent.末级:
            book.警示.append((parent.位置,'上级科目没有可核实的下级','余额表保留该科目的原金额，请检查源表范围'))
        elif children and 汇总(children,4) != parent.金额:
            book.警示.append((parent.位置,'上级金额与下级汇总不同','两层金额分别留存，请核对源表的科目与期间范围'))


def 维度资料(books, supplements):
    dimensions = defaultdict(set)
    for book in (*books,*supplements.values()):
        for row in book.行:
            if row.辅助记录 and row.辅助记录[0] != 未列明:
                dimensions[row.编码[3]].add(row.辅助记录[0])
    return dimensions


def 选定来源(books,supplements,rules,names):
    options = defaultdict(set)
    for book in (*books,*supplements.values()):
        for row in book.行:
            key = row.编码[0],row.编码[3]
            if row.辅助记录 and 采用维度(row,rules[key[1]],names[key[1]]):
                options[key].add(规范(row.辅助记录[0]))
    chosen = {}
    for key,labels in options.items():
        rule = rules[key[1]]
        if len(labels)>1:
            preferred = {'loan':('贷款银行','贷款行','借款银行','银行账户'),'interest':('贷款银行','贷款行','借款银行','银行账户'),
                         3:('部门','成本中心'),5:('银行账户','银行账号'),8:('往来单位',)}.get(rule[1],())
            match = next((label for label in preferred if label in labels),None)
            if not match:
                raise 输入错误(f'科目 {key[1]} 有多个同类辅助维度：{"、".join(sorted(labels))}。需要确定采用哪一列。')
            chosen[key] = match
        else:
            chosen[key] = next(iter(labels))
    return chosen


def 整理分录(journal, extra, rules, names, chosen):
    selected_parts = defaultdict(list)
    attached = 匹配流水(journal,extra)
    for row in journal.行:
        scope = row.编码[0],row.编码[3]
        rule = rules[scope[1]]
        items = list(attached.get(id(row),()))
        if row.辅助记录:
            same = [item for item in items if 规范(item.辅助记录[0]) == 规范(row.辅助记录[0])]
            if same and any(规范(item.辅助记录[2])!=规范(row.辅助记录[2]) or
                            (有值(item.辅助记录[1]) and 有值(row.辅助记录[1]) and item.辅助记录[1]!=row.辅助记录[1])
                            for item in same):
                raise 输入错误(f'{row.位置} 已带的辅助信息与另外提交的辅助明细冲突。')
            if not same:
                items.append(row)
        by_dim = defaultdict(list)
        for item in items:
            by_dim[规范(item.辅助记录[0])].append(item)
        for group in by_dim.values():
            if 汇总(group,2) != row.金额 and not 有据分拆(row.金额,group):
                raise 输入错误(f'{row.位置} 的「{group[0].辅助记录[0]}」明细借贷额无法与原分录对应，辅助行：{group[0].位置}。')
        selected = []
        audit_rows = []
        for item in items:
            keep = 规范(item.辅助记录[0]) == chosen.get(scope) and 采用维度(item,rule,names[scope[1]])
            记录(journal,'分录辅助已选' if keep else '分录其他维度',row,item)
            if keep:
                selected.append(item)
                audit_rows.append(len(journal.补充核对)-1)
            if rule[1] in ('loan','interest') and '合同' in item.辅助记录[0]:
                contract = item.辅助记录[1] or item.辅助记录[2]
                if row.合同 and row.合同 != contract:
                    raise 输入错误(f'{row.位置} 对应多个借款合同，无法唯一匹配贷款行。')
                row.合同 = contract
        segment = 维度来源(rule[1][4:]) if isinstance(rule[1],str) and rule[1].startswith('鼎信诺:') else rule[1]
        direct = ('','')
        if isinstance(segment,int) and segment in (3,5,6,7,8,9):
            token = row.编码[segment-1]
            direct = (token if 有值(token) else '',row.名称.get(segment,''))
        elif segment in ('loan','interest'):
            direct = row.贷款行
        if not 有值(direct[1]) and row.辅助记录 and 规范(row.辅助记录[0])==chosen.get(scope) and 采用维度(row,rule,names[scope[1]]):
            direct = row.辅助记录[1:]
        if 有值(direct[1]):
            label = chosen.get(scope) or rule[0]
            codes = {v for v in [direct[0],*(x.辅助记录[1] for x in selected)] if 有值(v)}
            if len(codes)>1 or any(规范(x.辅助记录[2])!=规范(direct[1]) for x in selected):
                raise 输入错误(f'{row.位置} 序时账与辅助序时账登记的对象不同。')
            partial = selected and 汇总(selected,2)!=row.金额
            if partial:
                journal.警示.append((row.位置,'辅助序时账仅列该对象的部分金额','原分录已登记同一对象，采用原分录全额'))
                gap = tuple(a-b for a,b in zip(row.金额,汇总(selected,2),strict=True))
                记录(journal,'该对象原分录中未在辅助序时账另列的金额',row,replace(row,辅助记录=(label,next(iter(codes),''),direct[1]),金额=gap))
            for i in audit_rows:
                journal.补充核对[i] = ('已包含在原分录的同一对象金额',*journal.补充核对[i][1:])
            selected = [replace(row,辅助记录=(label,next(iter(codes),''),direct[1]))]
            记录(journal,'原分录直接登记的辅助对象',row,selected[0])
        if not selected:
            continue
        if len({规范(item.辅助记录[0]) for item in selected}) > 1:
            raise 输入错误(f'{row.位置} 有多个符合规则的辅助维度，请确认应使用哪一项。')
        objects = sorted({(item.辅助记录[1],item.辅助记录[2]) for item in selected})
        warning = '辅助明细只覆盖部分分录金额，分项金额见辅助匹配' if 汇总(selected,2) != row.金额 else ''
        if len(objects) == 1:
            token,name = objects[0]
            if rule[1] in ('loan','interest'):
                row.贷款行 = token,name
            else:
                row.辅助指定 = token,name,warning
        elif rule[1] in ('loan','interest'):
            row.辅助指定 = '','贷款银行待核','同一条分录对应多个贷款银行，请核对贷款归属'
        else:
            tokens = '；'.join(token or '未登记编码' for token,_ in objects) if any(token for token,_ in objects) else ''
            row.辅助指定 = tokens,'；'.join(name for _,name in objects),warning
        row.辅助分项 = tuple((item.辅助记录[1],item.辅助记录[2],item.金额) for item in selected)
        selected_parts[scope].extend(selected)
    return selected_parts


def 整理余额(balance, extra, journal, journal_parts, rules, names, chosen):
    controls, details = 分组余额(balance,extra)
    核对上级(balance,controls)
    output = []
    by_account = defaultdict(list)
    period_matches = journal and journal.期间 and journal.期间 == balance.期间
    if period_matches:
        for row in journal.行:
            by_account[(row.编码[0],row.编码[3])].append(row)
    candidate_codes = {code for _,code in journal_parts}
    balance.控制合计 = 汇总((row for row in controls.values() if row.末级),4)
    for key,base in controls.items():
        rule = rules[key[1]]
        base.辅助指定 = '','',''
        output.append(base)
        记录(balance,'科目汇总',base,base)
        if not base.末级:
            continue
        items = details.get(key,[])
        selected = []
        for item in items:
            keep = 规范(item.辅助记录[0]) == chosen.get(key) and 采用维度(item,rule,names[key[1]])
            status = '余额辅助已选' if keep else ('原表未列明余额' if item.辅助记录[0] == 未列明 else '余额其他维度')
            记录(balance,status,base,item)
            if keep:
                selected.append(item)
        if not selected and key[1] in candidate_codes:
            jparts = journal_parts.get(key,())
            opening,debit,credit,closing = base.金额
            if not period_matches or not jparts or key[0] == '本批账套':
                balance.警示.append((base.位置,'序时账辅助待补录','已找到同科目辅助分录，请先核实公司和余额表的起止月份'))
            elif 汇总(by_account[key],2) != (debit,credit) or opening+debit-credit != closing:
                balance.警示.append((base.位置,'序时账辅助待补录','借贷发生额未逐项对上，或科目期初加借减贷与期末不一致'))
            else:
                for item in jparts:
                    d,c = item.金额
                    selected.append(replace(item,金额=(0,d,c,d-c)))
                    记录(balance,'仅分配序时账已知发生额，期初未分配',base,selected[-1])
                balance.警示.append((base.位置,'已补入本期辅助项目','项目期末仅列可确认的净变动，未知期初和未归属发生额另列'))
        if not selected:
            # 科目名已说明的分类仍可直接显示在汇总行，无源辅助维度由共用规则标记。
            base.辅助指定 = None
            if items or extra is not None:
                balance.警示.append((base.位置,'缺少所选维度的余额明细','该科目保留汇总，不按其他维度或未知期初拆分'))
            continue
        for item in selected:
            o,d,c,e = item.金额
            if o+d-c != e:
                balance.警示.append((item.位置,'原辅助余额滚动有差额','各列原金额保留，请检查是期间口径还是源数据差异'))
        groups = defaultdict(list)
        for item in selected:
            _,code,name = item.辅助记录
            groups[(code,name,None if 有值(code) else str(item.位置))].append(item)
        by_code = defaultdict(set)
        for code,name,_ in groups:
            if code:
                by_code[code].add(规范(name))
        conflicts = sorted(code for code,labels in by_code.items() if len(labels) > 1)
        if conflicts:
            raise 输入错误(f'{base.位置} 的辅助编码 {conflicts[0]} 对应多个名称，请确认后归并。')
        for (code,name,_), group in groups.items():
            output.append(replace(group[0],金额=汇总(group,4),行类型='辅助明细',辅助指定=(code,name,''),
                                  来源位置=tuple(r.位置 for r in group)))
        difference = tuple(a-b for a,b in zip(base.金额,汇总(selected,4),strict=True))
        if any(difference):
            if base.借贷方向 and any(r.借贷方向 not in ('','平',base.借贷方向) for r in selected):
                raise 输入错误(f'{base.位置} 辅助余额方向不同且无法逐列对回原科目。')
            known = [r for r in items if r.辅助记录[0]==未列明]
            supported = bool(known) and 汇总(known,4)==difference
            if not supported:
                for i in (1,2):
                    if (base.金额[i]>=0 and difference[i]<0 and all(r.金额[i]>=0 for r in selected)
                        or base.金额[i]<=0 and difference[i]>0 and all(r.金额[i]<=0 for r in selected)):
                        raise 输入错误(f'{base.位置} 已选辅助的借方或贷方发生额超过科目原数，无法确认剩余金额。')
                if any(base.金额[i]*difference[i]<0 for i in (0,3)):
                    balance.警示.append((base.位置,'辅助余额超过科目净余额','反向差额单列，请核实原表是否含反向余额'))
            issue = '期初或其他金额保留为未列明，未分摊给已有辅助项目'
            if difference[0]+difference[1]-difference[2] != difference[3]:
                issue += '；差额的期初加借减贷与期末不同，请复核源表'
            output.append(replace(base,金额=difference,行类型=未列明,辅助指定=('',未列明,issue)))
    balance.行 = output


def 应用辅助(books, supplements, rules, names):
    by_kind = {b.类别:b for b in books if b.补充核对 is not None}
    for book in books:
        if book.补充核对 is not None:
            continue
        if book.类别=='序时账' and ('核算项目明细表' in supplements or '余额表' in by_kind or '核算项目余额表' in supplements):
            book.补充核对 = []
            by_kind[book.类别] = book
        if book.类别=='余额表' and '核算项目余额表' in supplements:
            if any(row.行类型=='辅助明细' or any(有值(row.名称.get(s)) for s in (3,5,6,7,8,9)) for row in book.行):
                raise 输入错误(f'{book.文件} 已有单列的辅助对象，请按同一科目核实两份余额来源的重叠范围。')
            for row in book.行:
                row.行类型 = '科目汇总'
            book.补充核对 = []
            by_kind[book.类别] = book
    for key, kind in 辅助表.items():
        if key in supplements and kind not in by_kind:
            raise 输入错误(f'提供了{key}，还需提供能与它对应的鼎信诺{kind}。')
    journal,balance = by_kind.get('序时账'),by_kind.get('余额表')
    chosen = 选定来源(books,supplements,rules,names)
    parts = 整理分录(journal,supplements.get('核算项目明细表'),rules,names,chosen) if journal else {}
    if balance:
        整理余额(balance,supplements.get('核算项目余额表'),journal,parts,rules,names,chosen)
