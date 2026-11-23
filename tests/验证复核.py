"""用正式入口验证账表逐行保真、科目级次和辅助归属。"""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook


入口 = Path(__file__).resolve().parents[1] / 'scripts' / '清洗账表.py'
流水 = ['记账时间','凭证编号','凭证种类','编号','业务说明','科目编号','科目名称','借方发生额','贷方发生额']
余额 = ['科目编号','科目名称','借贷方向','是否明细科目','科目级次','账面期初数','账面借方发生额','账面贷方发生额','账面期末数']
辅助 = ['核算项目类型名称','核算项目编号','核算项目名称']
月 = '2025年1月至1月'
年 = '2025年1月至12月'


class 复核(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix='账表复核-'))

    def 写(self, name, rows, company='甲公司', period=月):
        book = Workbook()
        sheet = book.active
        if company:
            sheet.append(['公司名称：'+company])
        if period:
            sheet.append([period])
        for row in rows:
            sheet.append(row)
        path = self.dir/(name+'.xlsx')
        book.save(path)
        return path

    def 跑(self, files, *args, fails=''):
        result = subprocess.run([sys.executable,str(入口),'--输入',*map(str,files),'--输出目录',str(self.dir),*args],
                                capture_output=True,encoding='utf-8')
        if fails:
            self.assertNotEqual(result.returncode,0,result.stdout)
            self.assertIn(fails,result.stderr)
        else:
            self.assertEqual(result.returncode,0,result.stderr)
            return Path(result.stdout.split('成果目录：')[-1].strip())

    def 读(self, out, tab):
        file = tab+'-借贷两列.xlsx' if tab in ('序时账','余额表') else '清洗核对.xlsx'
        wb = load_workbook(out/file,read_only=True,data_only=True)
        try:
            return list(wb[tab].values)[1:]
        finally:
            wb.close()

    def test_序时账逐行保留及分列账号(self):
        head = [*流水[:2],'凭证字','编号','凭证摘要',*流水[5:],'开户行','银行账号','供应商名称']
        f = self.写('甲',[head,['2025-1-5','0007','收','1','转款','1002','银行存款',-9,0,'甲银行','00054',None],
                        ['2025-1-5','0007','付','2','货款','2202','应付账款',0,-9,None,None,'乙公司']])
        rows = self.读(self.跑([f]),'序时账')
        self.assertEqual([r[1:8] for r in rows],[('收','0007','转款','1002','银行存款',-9,0),('付','0007','货款','2202','应付账款',0,-9)])
        self.assertEqual(rows[0][8:10],('00054','甲银行 00054'))
        self.assertEqual(rows[1][9],'乙公司')

    def test_重复表头换列不能串写(self):
        line = ['2025-1-5','01','记','1','转款','1002','银行存款',1,0]
        f = self.写('甲',[流水,line,[*流水[:2],流水[4],流水[3],流水[2],*流水[5:]],line])
        self.跑([f],fails='表头')

    def test_歧义分拆及未确认公司(self):
        line = ['2025-1-5','01','记',None,'货款','2202','应付账款',0,100]
        f = self.写('甲',[流水,line,[*line[:-1],60]])
        a = self.写('乙',[[*流水,*辅助],[*line[:-1],60,'供应商','A','甲商']])
        out = self.跑([f,a])
        self.assertEqual([r[7] for r in self.读(out,'序时账')],[100,60])
        self.assertFalse(any(r[9]=='甲商' for r in self.读(out,'序时账')))
        self.assertTrue(any('原表未区分分录' in r[0] and r[6]=='甲商' and r[10]==60 for r in self.读(out,'辅助匹配')))
        unknown = self.写('丙',[[*流水,*辅助],[*line,'供应商','A','甲商']],company='')
        self.跑([f,unknown],fails='公司')

    def test_同期间贷款银行补回(self):
        line = ['2025-1-5','01','记','1','借款','2501','长期借款',0,80]
        f = self.写('甲',[流水,line])
        a = self.写('乙',[[*流水,*辅助],[*line,'贷款银行','B','甲银行']])
        b = self.写('丙',[余额,['2501','长期借款','贷','是',1,0,0,80,80]])
        self.assertIn(('B','甲银行'),[r[2:4] for r in self.读(self.跑([f,a,b]),'余额表')])
        year = self.写('丁',[余额,['2501','长期借款','贷','是',1,0,0,80,80]],period=年)
        self.assertNotIn(('B','甲银行'),[r[2:4] for r in self.读(self.跑([f,a,year]),'余额表')])

    def test_不同维度和超额不能凑平(self):
        line = ['2501','长期借款','贷','是',1,0,0,80,80]
        b = self.写('甲',[余额,line])
        a = self.写('乙',[[*余额,*辅助],[*line,'贷款银行','B','甲银行'],[*line,'往来单位','C','甲银行']])
        rows = self.读(self.跑([b,a]),'余额表')
        self.assertEqual([r[2:4] for r in rows[1:]],[('B','甲银行')])
        over = self.写('丙',[[*余额,*辅助],[*line[:7],100,100,'贷款银行','B','甲银行']])
        self.跑([b,over],fails='超过')

    def test_分列方向层级总计和来源(self):
        b = self.写('甲',[['科目编号','科目名称','余额方向','期初','借方','贷方','期末','科目级次'],
                        ['2202','应付账款','贷',80,0,2,82,1],['220201','应付账款-货款','贷',80,0,2,82,2]])
        out = self.跑([b])
        self.assertEqual([r[4] for r in self.读(out,'余额表')],[-80,-80])
        self.assertIn(('余额表','贷方发生额',2,2,0),[r[:5] for r in self.读(out,'金额核对')])
        line = ['2202','应付账款','贷','是',1,0,0,20,20]
        b = self.写('乙',[余额,line,['合计',None,'贷',None,None,0,0,30,30]])
        a = self.写('丙',[[*余额,*辅助],[*line[:7],8,8,'供应商','A','甲商'],[*line[:7],12,12,'供应商','A','甲商']])
        out = self.跑([b,a])
        self.assertTrue(any('合计' in str(r) for r in self.读(out,'待核记录')))
        self.assertEqual([r[-1] for r in self.读(out,'余额表来源') if r[0]==a.name],[3,3])

    def test_外币单位和指定列(self):
        duplicate = self.写('指定',[流水+[流水[7]],['2025-1-5','01','收','1','款','1002','银行存款',50,0,99]])
        config = self.dir/'列.json'
        config.write_text(json.dumps({'序时账':{'借方':'H'}},ensure_ascii=False),encoding='utf-8')
        self.assertEqual(self.读(self.跑([duplicate],'--指定列',str(config)),'序时账')[0][6],50)
        path = self.写('单位',[['金额单位：万元'],流水,['2025-1-5','01','记','1','款','1002','银行存款',5,0]])
        self.跑([path],fails='单位')

    def test_原分录已有供应商时保留并用于余额(self):
        line = ['2025-1-5','01','记','1','货款','2202','应付账款',0,60]
        second = [*line[:3],'2',*line[4:]]
        f = self.写('甲',[[*流水,'供应商编码','供应商名称'],[*line,'P1','甲单位'],[*second,'P2','乙单位']])
        a = self.写('乙',[[*流水,*辅助],[*line,'供应商','P1','甲单位']])
        b = self.写('丙',[余额,['2202','应付账款','贷','是',1,0,0,120,120]])
        for files in ([f,a,b],[f,b]):
            out = self.跑(files)
            self.assertEqual([r[8:10] for r in self.读(out,'序时账')],[('P1','甲单位'),('P2','乙单位')])
            self.assertEqual([r[2:4] for r in self.读(out,'余额表')[1:]],[('P1','甲单位'),('P2','乙单位')])

    def test_凭证的辅助金额也不能超额(self):
        line = ['2025-1-5','01','记',None,'货款','2202','应付账款',0,60]
        f = self.写('甲',[流水,line,line])
        a = self.写('乙',[[*流水,*辅助],*[[*line,'供应商',str(i),f'单位{i}'] for i in range(3)]])
        self.跑([f,a],fails='超过')

    def test_同一供应商两表互补编码且部分金额不重复(self):
        line = ['2025-1-5','01','记','1','货款','2202','应付账款',0,100]
        for main_code,aux_code,amount in [('P01',None,100),(None,'P01',60)]:
            f = self.写('甲',[[*流水,'供应商编码','供应商名称'],[*line,main_code,'甲单位']])
            a = self.写('乙',[[*流水,*辅助],[*line[:-1],amount,'供应商',aux_code,'甲单位']])
            out = self.跑([f,a])
            self.assertEqual(self.读(out,'序时账')[0][7:10],(100,'P01','甲单位'))
            if amount<100:
                self.assertTrue(any('部分' in str(r) for r in self.读(out,'待核记录')))
                self.assertTrue(any('已包含' in str(r[0]) for r in self.读(out,'辅助匹配')))
                self.assertTrue(any('未在辅助序时账另列' in r[0] and r[10]==40 for r in self.读(out,'辅助匹配')))

    def test_普通账套参数和同科目辅助行核对(self):
        f = self.写('甲',[['日期','凭证号','科目编号','科目名称','借方','贷方'],
                          ['2025-1-5','01','2202','应付账款',0,100]],company='')
        self.assertEqual(self.读(self.跑([f],'--公司标识','甲公司'),'序时账')[0][7],100)
        b = self.写('乙',[['科目编号','科目名称','科目级次','供应商名称','期初','借方','贷方','期末'],
                          ['2202','应付账款',1,None,0,0,100,-100],
                          ['2202','应付账款',None,'甲',0,0,25,-25],['2202','应付账款',None,'乙',0,0,75,-75]])
        out = self.跑([b])
        self.assertEqual([r[7] for r in self.读(out,'余额表')],[100,25,75])
        self.assertIn(('余额表','贷方发生额',100,100,0),[r[:5] for r in self.读(out,'金额核对')])

    def test_普通序时账使用单独的辅助序时账(self):
        f = self.写('甲',[['日期','凭证号','凭证字','分录行号','摘要','科目编号','科目名称','借方','贷方'],
                          ['2025-1-5','01','记','01','货款','2202','应付账款',0,60]])
        a = self.写('乙',[[*流水,*辅助],['2025-1-5','01','记','01','货款','2202','应付账款',0,60,'供应商','P1','甲单位']])
        out = self.跑([f,a])
        self.assertEqual(self.读(out,'序时账')[0][7:10],(60,'P1','甲单位'))


if __name__ == '__main__':
    unittest.main()
