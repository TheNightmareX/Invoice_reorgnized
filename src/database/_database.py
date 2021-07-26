import excel as _excel
from os.path import exists as _exists
from os import remove as _remove
from unite_str import unite_str as _unite_str


class DataBase:
    def __init__(self, path:str):
        """检查路径
        """
        if not path.endswith('.xls') and not path.endswith('.xlsx'):
            raise ValueError('文件类型不受支持')
        elif not _exists(path):
            raise FileNotFoundError('找不到指定的文件')
        self.path = path
        self.load()

    def load(self):
        """从文件加载数据
        """
        workbook = _excel.open_excel(self.path)
        self.goods = {}
        for name,code in _excel.read_range(workbook, '商品编码', [[2,1],[_excel.get_rows_count(workbook, '商品编码'),2]]):
            code = str(code)
            name = str(name)
            if len(code) < 19:
                self.goods[name] = f"{code:0<19}"
            else:
                self.goods[name] = code[:19]
        self.customers = {}
        for name,addr,code,account in _excel.read_range(workbook, '客户编码', [[2,1],[_excel.get_rows_count(workbook, '客户编码'),4]]):
            code,name,addr,account = str(code),str(name),str(addr),str(account)
            self.customers[name] = {'税号':code,'地址电话':addr,'银行账号':account}
        self.headers = {}
        for sect,kw,ign in _excel.read_range(workbook, '匹配规则', [[2,1],[_excel.get_rows_count(workbook, '匹配规则'),3]]):
            sect,kw,ign = str(sect),str(kw),str(ign)
            self.headers[sect] = {'keywords':kw.split(','),'ignore_list':ign.split(',')}
        _excel.close_excel(workbook, False)

    def flush(self):
        """将数据写入文件
        """
        if _exists(self.path):_remove(self.path)
        workbook = _excel.create_excel(self.path)
        _excel.sheet_rename(workbook, 0, '商品编码')
        _excel.write_range(workbook, '商品编码', 'A1', [('商品名称', '分类编码')] + list(self.goods.items()))
        for (idx, width) in enumerate((35, 20)):
            _excel.set_column_width(workbook, '商品编码', [1, idx + 1], width)
        _excel.create_sheet(workbook, '匹配规则')
        _excel.write_range(workbook, '匹配规则', 'A1', [['类别', '关键字列表', '屏蔽关键字列表']] + [[section, ','.join(cfg['keywords']), ','.join(cfg['ignore_list'])] for (section,cfg) in self.headers.items()])
        for (idx, width) in enumerate((10, 40, 40)):
            _excel.set_column_width(workbook, '匹配规则', [1, idx + 1], width)
        _excel.create_sheet(workbook, '客户编码')
        _excel.write_range(workbook, '客户编码', 'A1', [('客户名称', '地址电话', '税号', '银行账号')] + [(key, values['地址电话'], values['税号'], values['银行账号'])for (key, values) in self.customers.items()])
        for (idx, width) in enumerate((45, 100, 20, 70)):
            _excel.set_column_width(workbook, '客户编码', [1, idx + 1], width)
        _excel.close_excel(workbook)

    def update_from_txt(self, path:str):
        """从开票软件导出的txt文件加载数据并用这些数据更新
        """
        with open(path) as f:
            f_content = f.read().split('\n') 
            separator = f_content[0][12:-1]
        if f_content[0][1:5] == '客户编码':
            self.customers.update({_unite_str(line[1]):{'地址电话':line[4], '税号':line[3], '银行账号':line[5]} for line in [line.split(separator) for line in f_content[3:-1]]})
        elif f_content[0][1:5] == '商品编码':
            self.goods.update({_unite_str(line[1]):f"{line[11]:0<19}" for line in [line.split(separator) for line in f_content[3:-1]] if len(line) > 11})


