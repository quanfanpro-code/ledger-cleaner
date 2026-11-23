from pathlib import Path
import csv
import json
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts' / '清洗账表.py'


class 清洗行为(unittest.TestCase):
    def setUp(self):
        self.work = Path(tempfile.mkdtemp(prefix='账表清洗验证-'))
        self.journal = self.work / '任意名称甲.xlsx'
        self.balance = self.work / '任意名称乙.xlsx'
        self.output = self.work / '结果'
        self.make_journal()
        self.make_balance()

    def make_journal(self, duplicate=False, loan_conflict=False):
        book = Workbook()
        book.active.title = '封面'
        sheet = book.create_sheet('任意凭证页')
        sheet.merge_cells('A2:H2')
        sheet['A2'] = '导出记录'
        for region, value in [('A6:B6','本位币'),('C6:C7','会计日期'),('D6:D7','凭证编号'),
                              ('E6:E7','账户组合编码'),('F6:F7','账户说明'),('G6:G7','摘要'),
                              ('H6:H7','科目名称'),('I6:I7','行号')]:
            sheet.merge_cells(region)
            sheet[region.split(':')[0]] = value
        sheet['A7'], sheet['B7'] = '贷方金额', '借方金额'
        sheet['J6'], sheet['J7'] = '原币', '借方金额'
        if duplicate:
            sheet['K6'], sheet['K7'] = '本位币', '借方金额'
        sheet.append((0,100,'2025/01/02','000001','001.0.0.2501010107.0.0.009.S03.B01.0',
                      '测试公司.缺省.-.长期借款-人民币-本金-质押加抵押加保证借款.-.缺省.合同利率5.25%.示例银行城东支行.示例银行.缺省',
                      '本金','长期借款-人民币-本金-质押加抵押加保证借款',1,9))
        sheet.append((100,0,'2025/01/02','000001','001.0.0.2231020101.0.0.009.001.0.0',
                      '测试公司.缺省.-.应付利息-长期借款利息-正常利息.-.缺省.合同利率5.25%.测试公司.缺省.缺省',
                      '利息','应付利息-长期借款利息-正常利息',2,9))
        if loan_conflict:
            sheet.append((0,0,'2025/01/02','000001','001.0.0.2501010107.0.0.009.S04.B04.0',
                          '测试公司.缺省.-.长期借款-人民币-本金-质押加抵押加保证借款.-.缺省.合同利率5.25%.另一银行城西支行.另一银行.缺省',
                          '另一个银行','长期借款-人民币-本金-质押加抵押加保证借款',3,9))
        sheet.append((0,-10,'2025年01月02日','000002','001.0.0.1002010101.00005.0.0.001.B01.0',
                      '测试公司.缺省.-.银行存款-人民币.示例银行城东支行00001234.缺省.-.测试公司.示例银行.缺省',
                      '红字','银行存款-人民币',4,99))
        sheet.append((-10,0,'2025.01.02','000002','001.0.0.2203010101.0.001.009.0.0.0',
                      '测试公司.缺省.-.预收账款-经营性-房产预收款-预收房款.-.销售-住宅.示例项目.缺省.缺省.缺省',
                      '客户缺失','预收账款-经营性-房产预收款-预收房款',5,99))
        book.save(self.journal)

    def make_balance(self, no_cache=False):
        book = Workbook()
        sheet = book.active
        sheet.title = '科目明细'
        sheet['F1'] = '2025年01月至12月'
        for region, value in [('A3:A4','帐户说明'),('B3:B4','账户编码组合'),('C3:D3','本年发生额'),
                              ('E3:E4','期末余额'),('F3:F4','期初余额'),('G3:G4','会计科目编码')]:
            sheet.merge_cells(region)
            sheet[region.split(':')[0]] = value
        sheet['C4'], sheet['D4'] = '贷方', '借方'
        sheet.append(('测试公司-.-.银行存款-人民币.示例银行城东支行00001234-.-.测试公司.示例银行-',
                      '001.0.0.1002010101.00005.0.0.001.B01.0',0,-10,90,100,'1002010101'))
        sheet.append(('测试公司-.-.银行存款-人民币.示例银行城东支行00001234-.-.第二公司.示例银行-',
                      '001.0.0.1002010101.00005.0.0.002.B01.0',0,0,-40,-40,'1002010101'))
        sheet.append(('测试公司-.-.应付利息-长期借款利息-正常利息.--.合同利率5\\.25%.测试公司--',
                      '001.0.0.2231020101.0.0.009.001.0.0',100,0,-100,0,'2231020101'))
        sheet.append(('测试公司-.-.长期借款-人民币-本金-质押加抵押加保证借款.--.合同利率5\\.25%.示例银行城东支行.示例银行-',
                      '001.0.0.2501010107.0.0.009.S03.B01.0',0,100,100,0,'2501010107'))
        if no_cache:
            sheet['D5'] = '=1+2'
        book.save(self.balance)

    def run_skill(self, success=True, files=None, optional=()):
        files = files or [self.balance,self.journal]
        result = subprocess.run([sys.executable, str(SCRIPT), '--输入', *map(str,files),
                                 '--输出目录', str(self.output), *optional],
                                capture_output=True, encoding='utf-8')
        if success:
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            return next(self.output.glob('账表清洗-*'))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(list(self.output.glob('账表清洗-*')))
        return result.stderr + result.stdout

    @staticmethod
    def load_rows(path, name):
        book = load_workbook(path, read_only=True, data_only=True)
        rows = list(book[name].values)
        book.close()
        return rows

    def test_合并表头重排负数和贷款匹配(self):
        result = self.run_skill()
        journal = self.load_rows(result / '序时账-借贷两列.xlsx', '序时账')
        balance = self.load_rows(result / '余额表-借贷两列.xlsx', '余额表')
        self.assertEqual(len(journal), 5)
        self.assertEqual(journal[0], ('日期','凭证类型','凭证号','摘要','科目编码','科目名称','借方金额','贷方金额',
                                      '辅助核算编码','辅助核算名称','一级编码','一级科目'))
        self.assertEqual(journal[3][-2:], ('1002','银行存款'))
        self.assertEqual([r[8:10] for r in journal[1:3]], [('S03','示例银行城东支行')]*2)
        self.assertEqual(journal[1][2], '000001')
        self.assertEqual(journal[3][6:10], (-10,0,'00005','示例银行城东支行00001234'))
        self.assertEqual(journal[4][9], '往来未记录')
        self.assertEqual(len(balance), 4)
        self.assertEqual(balance[1][2:4], ('00005','示例银行城东支行00001234'))
        self.assertEqual(balance[0], ('科目编码','科目名称','辅助核算编码','辅助核算名称','期初余额(借方)','期初余额(贷方)',
                                      '累计发生额(借方)','累计发生额(贷方)','期末余额(借方)','期末余额(贷方)','行类型'))
        self.assertEqual(balance[1][4:], (60,None,-10,0,50,None,'辅助明细'))
        self.assertEqual(balance[2][2:4], ('S03','示例银行城东支行'))
        self.assertTrue((result / '清洗核对.xlsx').is_file())

    def test_重复本位币列报位置(self):
        self.make_journal(duplicate=True)
        error = self.run_skill(success=False)
        self.assertIn('借方', error)
        self.assertIn('K', error)
        choice = self.work / '确认列.json'
        choice.write_text(json.dumps({'序时账':{'借方':'B'}},ensure_ascii=False),encoding='utf-8')
        result = self.run_skill(optional=('--指定列',str(choice)))
        journal = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        self.assertEqual(journal[1][6:8],(100,0))

    def test_公式未重算不按零继续(self):
        self.make_balance(no_cache=True)
        error = self.run_skill(success=False)
        self.assertIn('D5', error)
        self.assertIn('公式', error)

    def test_同一合同出现多个贷款行保持待核(self):
        self.make_journal(loan_conflict=True)
        result = self.run_skill()
        journal = self.load_rows(result / '序时账-借贷两列.xlsx', '序时账')
        self.assertEqual(journal[2][9], '贷款银行待核')

    def test_单表自动判断和未知金额单位保护(self):
        result = self.run_skill(files=[self.balance])
        self.assertTrue((result/'余额表-借贷两列.xlsx').is_file())
        self.assertFalse((result/'序时账-借贷两列.xlsx').exists())
        self.output = self.work/'单位结果'
        book = load_workbook(self.balance)
        book.active['A1'] = '金额单位：万元'
        book.save(self.balance)
        self.assertIn('换算',self.run_skill(success=False,files=[self.balance]))

    def test_自动识别分列序时账并按科目取维度(self):
        book = Workbook()
        sheet = book.active
        sheet.append(('凭证日期','凭证编号','科目编号','科目名称','本位币借方金额','本位币贷方金额',
                      '开户行','银行账号','客户名称','客户编码','供应商名称','供应商编码','账套编码',
                      '合同编号','贷款银行','贷款银行编码'))
        sheet.append(('2025/01/02','000017','1002010101','银行存款-人民币',100,0,
                      '示例银行城东支行',1234,'某客户','K01','某供应商','S01','001'))
        sheet['H2'].number_format = '00000000'
        sheet.append(('2025/01/02','000017','2202010301','应付账款-经营性-应付工程款-开票款',0,100,
                      None,None,'某客户','K01','某供应商','S01','001'))
        sheet.append(('2025/01/02','000018','2231020101','应付利息-长期借款利息-正常利息',0,20,
                      None,None,None,None,None,None,'001','C01','示例银行城东支行','B01'))
        book.save(self.journal)
        result = self.run_skill(files=[self.journal])
        rows = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        self.assertEqual(rows[1][8:10],('00001234','示例银行城东支行 00001234'))
        self.assertEqual(rows[2][8:10],('S01','某供应商'))
        self.assertEqual(rows[3][8:10],('B01','示例银行城东支行'))
        trace = self.load_rows(result/'清洗核对.xlsx','序时账来源')
        self.assertIsNone(trace[1][6])

    def test_查看结构和适配自定义余额表(self):
        book = Workbook()
        sheet = book.active
        sheet.title = '区间汇总'
        sheet['A1'] = '2025年01月至12月'
        sheet['C1'] = '账套编码001'
        sheet.merge_cells('A2:G2')
        sheet['A2'] = '汇总对照'
        sheet.append(('余额新','编号','摘要性质','增','减','余额旧','单位名','单位键'))
        sheet.append((40,'2202010301','应付账款-经营性-应付工程款-开票款',50,10,0,'甲供应商','S01'))
        sheet.append((60,'2202010301','应付账款-经营性-应付工程款-开票款',70,10,0,'甲供应商','S01'))
        sheet.append((15,'2202010301','应付账款-经营性-应付工程款-开票款',15,0,0,'甲供应商',None))
        sheet.append(('','合计','',135,20,0))
        sheet['A7'] = 115
        book.save(self.balance)
        report = subprocess.run([sys.executable,str(SCRIPT),'--输入',str(self.balance),'--查看结构'],
                                capture_output=True,encoding='utf-8',check=True)
        structure = json.loads(report.stdout)
        self.assertIn('A2:G2', structure[0]['工作表'][0]['合并区域'])
        self.assertEqual(structure[0]['工作表'][0]['样例行'][2]['单元格']['B'],'编号')
        plan = self.work/'本次版式.json'
        plan.write_text(json.dumps({self.balance.name: {'类型':'余额表','工作表':'区间汇总','表头末行':3,
            '公司标识':'001','列':{'期末':'A','科目编码':'B','科目名称':'C','借方':'D','贷方':'E','期初':'F',
                              '供应商':'G','供应商编码':'H'}}},ensure_ascii=False),encoding='utf-8')
        result = self.run_skill(files=[self.balance],optional=('--版式方案',str(plan)))
        rows = self.load_rows(result/'余额表-借贷两列.xlsx','余额表')
        self.assertEqual(len(rows),3)
        self.assertEqual(rows[1][2:4],('S01','甲供应商'))
        self.assertEqual(rows[1][6:9],(120,20,100))
        self.assertEqual(rows[2][2:4],(None,'甲供应商'))
        audit = self.load_rows(result/'清洗核对.xlsx','字段定位')
        self.assertIn(('余额表',self.balance.name,'区间汇总','科目编码','B','版式方案'),audit)
        trace = self.load_rows(result/'清洗核对.xlsx','余额表来源')
        self.assertEqual([r[2] for r in trace[1:]],[4,5,6])

    def test_无法确定公司范围时先确认(self):
        book = Workbook()
        sheet = book.active
        sheet.append(('科目编码','科目名称','借方','贷方','期初余额','期末余额'))
        sheet.append(('1001010101','库存现金-人民币',1,0,0,1))
        book.save(self.balance)
        self.assertIn('公司范围',self.run_skill(success=False,files=[self.balance]))

    def test_账户组合与分列可混用且保留来源(self):
        book = Workbook()
        sheet = book.active
        sheet['A1'] = '账套编码：001'
        sheet['A2'] = '2025年01月至01月'
        sheet.append(('科目编码','科目名称','借方','贷方','期初余额','期末余额','银行账号','开户行'))
        sheet.append(('1002010101','银行存款-人民币',-10,0,100,90,'00001234','示例银行城东支行'))
        book.save(self.balance)
        result = self.run_skill()
        journal_rows = self.load_rows(result/'清洗核对.xlsx','序时账来源')
        balance_rows = self.load_rows(result/'清洗核对.xlsx','余额表来源')
        self.assertEqual(journal_rows[3][4:7],('001','1002010101','001.0.0.1002010101.00005.0.0.001.B01.0'))
        self.assertEqual(balance_rows[1][4:7],('001','1002010101',None))
        checks = self.load_rows(result/'清洗核对.xlsx','金额核对')
        self.assertEqual(checks[-1][5],'公司、会计期间可对应')
        auxiliary_diffs = self.load_rows(result/'清洗核对.xlsx','跨表辅助差异')
        self.assertFalse(any(r[1] == '1002010101' for r in auxiliary_diffs[1:]))
        bank_matches = self.load_rows(result/'清洗核对.xlsx','跨表银行对应')
        self.assertEqual(bank_matches[1][:4],('001','1002010101','00005','00001234'))

    def test_分列表的同一项目不能代替借款合同(self):
        book = Workbook()
        sheet = book.active
        sheet.append(('日期','凭证号','科目编码','科目名称','借方','贷方','公司编码','项目编码','贷款银行','贷款银行编码'))
        sheet.append(('2025/01/02','001','2501010107','长期借款-人民币-本金-质押加抵押加保证借款',
                      100,0,'001','开发01','示例银行城东支行','B01'))
        sheet.append(('2025/01/02','001','2231020101','应付利息-长期借款利息-正常利息',0,100,'001','开发01'))
        book.save(self.journal)
        result = self.run_skill(files=[self.journal])
        rows = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        self.assertEqual(rows[1][8:10],('B01','示例银行城东支行'))
        self.assertEqual(rows[2][8:10],(None,'贷款银行待核'))

    def make_ding(self, empty_balance=False):
        journal = ['记账时间','凭证编号','凭证种类','编号','业务说明','科目编号','科目名称','借方发生额','贷方发生额']
        balance = ['科目编号','科目名称','借贷方向','是否明细科目','科目级次','账面期初数','账面借方发生额','账面贷方发生额','账面期末数']
        aux = ['核算项目类型名称','核算项目编号','核算项目名称']
        credit = ['2025-01-05','0007','记','02','材料货款','220201','应付账款-货款',0,100]
        table_rows = [
            [journal, ['2025-01-05','0007','记','01','材料货款','1002','银行存款',100,0], credit],
            [journal+aux,
             [*credit[:-1],60,'供应商','0001','甲供应商'], [*credit[:-1],40,'供应商','0002','乙供应商'],
             [*credit,'部门','001','采购部']],
            [balance, ['2202','应付账款','贷','否',1,25,0,100,125],
             ['1002','银行存款','借','是',1,5,100,0,105],
             ['220201','应付账款-货款','贷','是',2,25,0,100,125]],
            [balance+aux,
             ['220201','应付账款-货款','贷','是',2,5,0,60,65,'供应商','0001','甲供应商'],
             ['220201','应付账款-货款','贷','是',2,10,0,40,50,'供应商','0002','乙供应商'],
             ['220201','应付账款-货款','贷','是',2,25,0,100,125,'部门','001','采购部']],
        ]
        if empty_balance:
            table_rows[2].insert(0,['2025年1月至1月'])
            table_rows[3] = [['2025年1月至1月'], table_rows[3][0]]
        paths = [self.work / f'随机文件{i}.xls' for i in range(4)]
        for path, rows in zip(paths, table_rows, strict=True):
            with path.open('w', encoding='gb18030', newline='') as file:
                csv.writer(file, delimiter='\t').writerows(rows)
        return paths

    def test_鼎信诺四表按内容识别维度及金额层级(self):
        paths = self.make_ding()
        result = self.run_skill(files=paths[::-1])
        journal = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        balance = self.load_rows(result/'余额表-借贷两列.xlsx','余额表')
        self.assertEqual(len(journal), 3)
        self.assertEqual(journal[2][8:12], ('0001；0002','甲供应商；乙供应商','2202','应付账款'))
        self.assertIn(('220201','应付账款-货款','0001','甲供应商',-5,None,0,60,-65,None,'辅助明细'), balance)
        self.assertIn(('220201','应付账款-货款',None,'未列明辅助项目',-10,None,0,0,-10,None,'未列明辅助项目'), balance)
        audit = self.load_rows(result/'清洗核对.xlsx','辅助匹配')
        self.assertTrue(any('采购部' in row for row in audit))
        self.assertTrue(any('甲供应商' in row and 60 in row for row in audit))
        check = self.load_rows(result/'清洗核对.xlsx','金额核对')
        self.assertTrue(any(row[:5] == ('余额表','贷方发生额',100,100,0) for row in check[1:]))
        self.assertFalse(any(row[0] == '2202' for row in balance[1:]))

    def test_鼎信诺重复候选时禁止猜测分录归属(self):
        paths = self.make_ding()
        with paths[0].open('a', encoding='gb18030', newline='') as f:
            csv.writer(f, delimiter='\t').writerow(
                ['2025-01-05','0007','记','02','材料货款','220201','应付账款-货款',0,100])
        failure = self.run_skill(False, paths)
        self.assertIn('无法唯一匹配', failure)

    def test_鼎信诺空辅助余额只恢复已知发生额(self):
        result = self.run_skill(files=self.make_ding(empty_balance=True),optional=('--公司标识','模拟公司'))
        rows = self.load_rows(result/'余额表-借贷两列.xlsx','余额表')
        self.assertIn(('220201','应付账款-货款','0001','甲供应商',0,None,0,60,-60,None,'辅助明细'),rows)
        self.assertIn(('220201','应付账款-货款',None,'未列明辅助项目',-25,None,0,0,-25,None,'未列明辅助项目'),rows)
        logs = self.load_rows(result/'清洗核对.xlsx','辅助匹配')
        self.assertTrue(any(row[0] == '仅分配序时账已知发生额，期初未分配' for row in logs))

    def test_鼎信诺新版工作簿用合并表头辨认本位币(self):
        book = Workbook()
        sheet = book.active
        sheet.append(['示例报表'])
        titles = ['记账时间','凭证编号','凭证种类','业务说明','科目编号','科目名称']
        sheet.append([*titles,'原币',None,'本位币',None])
        sheet.append([None]*6+['借方金额','贷方金额','借方金额','贷方金额'])
        for col in range(1,7):
            sheet.merge_cells(start_row=2,start_column=col,end_row=3,end_column=col)
        sheet.merge_cells('G2:H2')
        sheet.merge_cells('I2:J2')
        sheet.append(['2025-01-05','0007','记','银行转款','1002','银行存款',999,0,100,0])
        sheet.append(['2025-01-05','0007','记','银行转款','2202','应付账款',0,999,0,100])
        path = self.work/'随便命名的新版导出.xlsx'
        book.save(path)
        result = self.run_skill(files=[path])
        rows = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        self.assertEqual([row[6:8] for row in rows[1:]],[(100,0),(0,100)])

    def test_不接受与原表冲突的版式方案(self):
        book = Workbook()
        sheet = book.active
        sheet.append(('科目编号','科目名称','借方','贷方','期初','期末','公司编码'))
        sheet.append(('1001010101','库存现金-人民币',10,0,0,10,'001'))
        sheet.append(('1001010101','库存现金-人民币',12,0,0,12,'001'))
        book.save(self.balance)
        columns = {'科目编码':'A','科目名称':'B','借方':'C','贷方':'D','期初':'E','期末':'F','公司编码':'G'}
        plan = self.work/'错列.json'
        for end,mapping,reason in ((1,{**columns,'借方':'D','贷方':'C'},'冲突'),(2,columns,'金额数据')):
            with self.subTest(reason=reason):
                plan.write_text(json.dumps({self.balance.name:{'类型':'余额表','工作表':'Sheet',
                    '表头末行':end,'列':mapping}},ensure_ascii=False),encoding='utf-8')
                response = self.run_skill(success=False,files=[self.balance],optional=('--版式方案',str(plan)))
                self.assertIn(reason,response)


if __name__ == '__main__':
    unittest.main()
