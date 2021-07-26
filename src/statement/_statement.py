from database import DataBase as _DataBase
from decimal import Decimal as _Decimal
from os.path import splitext as _splitext
import excel as _excel
import xlrd as _xlrd
from typing import List as _List
from typing import Dict as _Dict
from typing import Tuple as _Tuple
import json as _json
from xml.dom.minidom import Document as _XmlDoc
from unite_str import unite_str as _unite_str



class Statement:
    class HeaderMissingError(Exception):
        pass

    class Details:
        def __init__(self, data:_List[_Dict[str, str or _Decimal]]):
            self.data = data
             
        def lock(self, indexes:_List[str]):
            """锁定indexes列出的商品行
            """
            for index in indexes:
                if not index.isnumeric():
                    raise ValueError(f"“{index}”不是整数")
                elif int(index) not in range(1, 1 + len(self.data)):
                    raise ValueError(f"“{index}”越界")
                else:
                    index = int(index)
                    self.data[index - 1]['锁定'] = True

        def merge(self):
            """合并相同商品的商品行 不会合并锁定的商品行
            """
            details = {}
            for (idx, line) in enumerate(self.data):
                key = f"{line['商品名称']} | {line['规格型号']}"
                if line['锁定']:
                    key += str(line['序号'])
                    del line['序号']
                    details[key] = line
                else:
                    del line['序号']
                    if key not in details:
                        details[key] = line
                    else:
                        details[key]['数量'] += line['数量']
            self.data = []
            for key, value in details.items():
                detail = {'商品名称':key.split(' | ')[0], '规格型号':key.split(' | ')[1]}
                detail.update(value)
                self.data.append(detail)

    def __init__(self, database:_DataBase, path:str):
        """检查文件后缀 读取数据
        """

        def check_extention(path:str):
            """检查文件后缀名是否为xls或xlsx
            若不是则引发ValueError(f'不受支持的对账单类型:{后缀名}')
            """
            extention = _splitext(path)[1]
            if extention != '.xls' and extention != '.xlsx':
                raise ValueError(f'不受支持的对账单类型:{extention}')

        def read_sheet(path:str, sheet_idx:int) -> _List[_List]:
            """返回xls/xlsx的某一页的所有数据
            """
            if path.endswith('.xlsx'):
                workbook = _excel.open_excel(path)
                data = _excel.read_range(workbook, sheet_idx, [[1, 1], [_excel.get_rows_count(workbook, sheet_idx), _excel.get_columns_count(workbook, sheet_idx)]])
                _excel.close_excel(workbook, False)
            elif path.endswith('.xls'):
                workbook = _xlrd.open_workbook(path)
                sheet = workbook.sheet_by_index(sheet_idx)
                data = [[cell.value for cell in line] for line in sheet.get_rows()]
            return data

        def search_headers(sheet_data:_List[_List], searching_rules:_Dict[str, _Dict[str, _List[str]]]) -> _Dict[str, _Tuple[int, int]]:
            """根据searching_rules在sheet_data中查找表头
            """
            HEADERS_TO_FIND = {
                '商品名称',
                '规格型号',
                '单位',
                '数量',
                '含税单价',
                '未税单价',
                '客户',
                '序号'
            }
            header_locs = {}
            for (row, line) in enumerate(sheet_data):
                for (col, cell) in enumerate(line):
                    if type(cell) != str:continue # 跳过非字符单元格
                    for (header, rule) in searching_rules.items():
                        if header in header_locs:continue # 重复找到 只记录第一次找到的位置
                        for kw in rule['ignore_list']:
                            if kw in _unite_str(cell):break # 如果存在屏蔽字段则跳过
                        else: # 合格的单元格
                            for kw in rule['keywords']:
                                if kw in _unite_str(cell):
                                    header_locs[header] = (row, col)
                                    break
            # 检查表头查找情况
            headers_not_found = list(HEADERS_TO_FIND - set(header_locs.keys()))
            if '未税单价' in headers_not_found and '含税单价' in header_locs: # 未税单价和含税单价只需要一个就够了
                headers_not_found.remove('未税单价')
            elif '含税单价' in headers_not_found and '未税单价' in header_locs:
                headers_not_found.remove('含税单价')
            if len(headers_not_found) > 0:
                raise self.HeaderMissingError(f"在对账单中匹配这些表头失败:{'、'.join(headers_not_found)}")
            return header_locs

        def parse_sheet_data(sheet_data:_List[_List], header_locs:_Dict[str, _Tuple[int, int]]) -> _Tuple:
            """根据header_locs解析sheet_data 提取数据
            """
            def parse_customer(sheet_data:_List[_List], header_locs:_Dict[str, _Tuple[int, int]]) -> str:
                """解析客户信息
                返回客户名称
                """
                cell:str = sheet_data[header_locs['客户'][0]][header_locs['客户'][1]]
                cell = _unite_str(cell)
                customer = cell.split(':')[-1]
                return customer

            def parse_line_count(sheet_data:_List[_List], header_locs:_Dict[str, _Tuple[int, int]]) -> int:
                """解析行数
                返回行数
                """
                row = header_locs['序号'][0] + 1
                col = header_locs['序号'][1]
                line_count = 0
                for line in sheet_data[row:]:
                    cell = line[col]
                    if type(cell) == int or type(cell) == float and int(cell) == cell:
                        line_count += 1
                    else:
                        break
                return line_count

            def parse_details(sheet_data:_List[_List], header_locs:_Dict[str, _Tuple[int, int]]) -> _List[_Dict[str, str or _Decimal]]:
                """解析商品行详细信息
                返回列表
                """
                details = [{'锁定':False} for i in range(line_count)]
                for header, (row, col) in header_locs.items():
                    row_begin = row + 1
                    row_end = row_begin + line_count
                    for idx, line in enumerate(sheet_data[row_begin:row_end]):
                        cell = line[col]
                        if type(cell) == float or type(cell) == int:cell = _Decimal(str(cell))
                        if header != '含税单价':details[idx][header] = cell
                        else:details[idx]['未税单价'] = cell / _Decimal('1.13') # 含税转未税
                return details

            customer = parse_customer(sheet_data, header_locs)
            del header_locs['客户']
            line_count = parse_line_count(sheet_data, header_locs)
            details = parse_details(sheet_data, header_locs)
            return (
                customer,
                details
            )

        check_extention(path)
        self.path = path
        self.database = database
        self.sheet_data = read_sheet(path, 0)
        self.header_locs = search_headers(self.sheet_data, self.database.headers)
        self.customer, self.details = parse_sheet_data(self.sheet_data, self.header_locs)
        self.details = self.Details(self.details)

    def convert_to_xml(self) -> _List[_Dict[str, str or _Decimal]]:
        """将self.details.data分组 用分组后的数据创建符合开票软件接口规范的XML单据
        返回分组后的数据
        """
        def group_details(details:Statement.Details) -> _List[_Dict[str, str or _Decimal]]:
            """将details.data分组
            每组不超过八条
            每组金额不超过或等于130000
            返回：分组后的数据
            """
            grouped_detailed_data = [{'add_up':0, 'count':0, 'details':[]}]
            data = details.data.copy()
            while len(data) > 0:
                for count in range(int(data[0]['数量']), 0, -1):
                    if '.' in str(data[0]['数量']):count = _Decimal(str(count) + '.' + str(data[0]['数量']).split('.')[1]) # 如果数量不是整数 先去掉小数
                    if grouped_detailed_data[0]['count'] < 8 and grouped_detailed_data[0]['add_up'] + count * data[0]['未税单价'] < 100000:
                        if data[0]['锁定'] and count != int(data[0]['数量']):continue # 不拆分锁定的行
                        grouped_detailed_data[0]['count'] += 1
                        grouped_detailed_data[0]['add_up'] += count * data[0]['未税单价']
                        grouped_detailed_data[0]['details'].append(
                            {
                                '锁定' : data[0]['锁定'],
                                '序号' : len(grouped_detailed_data[0]['details']) + 1,
                                '商品名称' : _unite_str(data[0]['商品名称']),
                                '规格型号' : data[0]['规格型号'],
                                '单位' : data[0]['单位'],
                                '未税单价' : data[0]['未税单价'],
                                '数量' : count,
                                '金额' : data[0]['未税单价'] * count
                            }
                        )
                        data[0]['数量'] -= count
                        if data[0]['数量'] == 0:del data[0]
                        break
                else: # 满8条了或者金额饱和了:新增
                    grouped_detailed_data.insert(0, {'add_up':0, 'count':0, 'details':[]})
            return grouped_detailed_data
            
        groups = group_details(self.details)
        xml = _XmlDoc()
        xml.root = xml.appendChild(xml.createElement('Kp')) # 根
        xml.root.appendChild(xml.createElement('Version')).appendChild(xml.createTextNode('2.0')) # 版本
        xml.root = xml.root.appendChild(xml.createElement('Fpxx')) # 发票信息
        xml.root.appendChild(xml.createElement('Zsl')).appendChild(xml.createTextNode(str(len(groups)))) # 总数量
        xml.root = xml.root.appendChild(xml.createElement('Fpsj')) # 发票数据根
        for (group_idx, group_data) in enumerate(groups):
            xml.fproot = xml.root.appendChild(xml.createElement('Fp')) # 发票根
            xml.fproot.appendChild(xml.createElement('Djh')).appendChild(xml.createTextNode(str(group_idx + 1))) # 单据号
            xml.fproot.appendChild(xml.createElement('Gfmc')).appendChild(xml.createTextNode(self.customer)) # 购方名称
            customer_info = self.database.customers.get(self.customer)
            try:
                to_write =  customer_info['税号']
            except:
                to_write = '0000000000000000000'
            xml.fproot.appendChild(xml.createElement('Gfsh')).appendChild(xml.createTextNode(to_write)) # 购方税号
            try:
                to_write =  customer_info['银行账号']
            except:
                to_write = '未知'
            xml.fproot.appendChild(xml.createElement('Gfyhzh')).appendChild(xml.createTextNode(to_write)) # 购方银行账号
            try:
                to_write =  customer_info['地址电话']
            except:
                to_write = '未知'
            xml.fproot.appendChild(xml.createElement('Gfdzdh')).appendChild(xml.createTextNode(to_write)) # 购方地址电话
            xml.fproot.appendChild(xml.createElement('Bz')).appendChild(xml.createTextNode('')) # 备注
            xml.fproot.appendChild(xml.createElement('Fhr')).appendChild(xml.createTextNode('赵婷')) # 复核人
            xml.fproot.appendChild(xml.createElement('Skr')).appendChild(xml.createTextNode('陈杨')) # 收款人
            xml.fproot.appendChild(xml.createElement('Spbmbbh')).appendChild(xml.createTextNode('33.0')) # 商品编码版本号
            xml.fproot.appendChild(xml.createElement('Hsbz')).appendChild(xml.createTextNode('0')) # 含税标志
            xml.fproot = xml.fproot.appendChild(xml.createElement('Spxx')) # 商品信息根
            for (line_idx,line_data) in enumerate(group_data['details']):
                xml.lineroot = xml.fproot.appendChild(xml.createElement('Sph')) # 商品行根
                xml.lineroot.appendChild(xml.createElement('Xh')).appendChild(xml.createTextNode(str(line_idx + 1))) # 序号
                xml.lineroot.appendChild(xml.createElement('Spmc')).appendChild(xml.createTextNode(str(line_data['商品名称']))) # 商品名称
                xml.lineroot.appendChild(xml.createElement('Ggxh')).appendChild(xml.createTextNode(str(line_data['规格型号']))) # 规格型号
                xml.lineroot.appendChild(xml.createElement('Jldw')).appendChild(xml.createTextNode(str(line_data['单位']))) # 计量单位
                to_write = self.database.goods.get(line_data['商品名称'])
                if to_write == None:raise Exception(f'商品“{line_data["商品名称"]}”的分类编码未知')
                xml.lineroot.appendChild(xml.createElement('Spbm')).appendChild(xml.createTextNode(to_write)) # 商品编码
                xml.lineroot.appendChild(xml.createElement('Qyspbm')).appendChild(xml.createTextNode('')) # 企业商品编码
                xml.lineroot.appendChild(xml.createElement('Syyhzcbz')).appendChild(xml.createTextNode('0')) # 优惠政策标识
                xml.lineroot.appendChild(xml.createElement('Lslbz')).appendChild(xml.createTextNode('')) # 零税率标识
                xml.lineroot.appendChild(xml.createElement('Yhzcsm')).appendChild(xml.createTextNode('')) # 优惠政策说明
                xml.lineroot.appendChild(xml.createElement('Dj')).appendChild(xml.createTextNode(str(line_data['未税单价']))) # 单价
                xml.lineroot.appendChild(xml.createElement('Sl')).appendChild(xml.createTextNode(str(line_data['数量']))) # 数量
                xml.lineroot.appendChild(xml.createElement('Je')).appendChild(xml.createTextNode(str(line_data['金额']))) # 金额
                xml.lineroot.appendChild(xml.createElement('Slv')).appendChild(xml.createTextNode('0.13')) # 税率
                xml.lineroot.appendChild(xml.createElement('Se')).appendChild(xml.createTextNode(str(line_data['金额'] * _Decimal('0.13')))) # 税额
                xml.lineroot.appendChild(xml.createElement('Kce')).appendChild(xml.createTextNode('0')) # 扣除额
        with open(f'{self.path}.xml', mode='w', encoding='utf-8') as f:
            xml.writexml(f, indent='\t', addindent='\t', newl='\n', encoding='utf-8')
        return groups