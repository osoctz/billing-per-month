import xlwt as xlwt


def wt_head(sheet):
    style = st_style()

    sheet.write(1, 0, '单位:', style)
    sheet.write_merge(1, 1, 1, 3, 'xx', style)

    sheet.write(1, 4, '部门:', style)
    sheet.write_merge(1, 1, 5, 7, 'xxx', style)

    sheet.write(1, 8, '姓名:', style)
    sheet.write_merge(1, 1, 9, 10, 'xx', style)

    sheet.write_merge(1, 1, 11, 12, '出差事由:', style)
    sheet.write_merge(1, 1, 13, 16, 'xxx', style)

    sheet.write_merge(2, 2, 0, 3, '出发', style)
    sheet.write_merge(2, 2, 4, 7, '到达', style)

    sheet.write_merge(2, 3, 8, 8, '交通工具', style)
    sheet.write_merge(2, 3, 9, 9, '交通费金额', style)

    sheet.write_merge(2, 2, 10, 12, '餐费补助', style)
    sheet.write_merge(2, 2, 13, 16, '住宿费', style)

    sheet.write(3, 0, '月', style)
    sheet.write(3, 1, '日', style)
    sheet.write(3, 2, '时间', style)
    sheet.write(3, 3, '地点', style)

    sheet.write(3, 4, '月', style)
    sheet.write(3, 5, '日', style)
    sheet.write(3, 6, '时间', style)
    sheet.write(3, 7, '地点', style)

    sheet.write(3, 10, '日期', style)
    sheet.write(3, 11, '天数', style)
    sheet.write(3, 12, '金额', style)

    sheet.write(3, 13, '住宿起止时间', style)
    sheet.write(3, 14, '天数', style)
    sheet.write(3, 15, '住宿人员', style)
    sheet.write(3, 16, '金额', style)


def st_style():
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = '黑体'
    font.height = 20 * 10
    font.bold = False  # 黑体
    # 设置边框
    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    # 设置单元格对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    style.font = font
    style.borders = borders
    return style


def wt_title(sheet, title='差　旅　费　报　销　明  细'):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = '黑体'
    font.height = 20 * 18
    font.bold = True  # 黑体
    font.underline = True  # 下划线

    # 设置单元格对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.alignment = alignment
    style.font = font
    sheet.write_merge(0, 0, 0, 16, title, style)


def wt_floor(sheet, line_num=5):
    style = st_style()
    sheet.write_merge(line_num + 1, line_num + 1, 0, 8, '合计', style)

    sheet.write(line_num + 1, 9, xlwt.Formula('SUM(J5:J' + str(line_num) + ')'), style)
    sheet.write(line_num + 1, 10, '', style)
    sheet.write(line_num + 1, 11, '', style)
    sheet.write(line_num + 1, 12, '', style)

    sheet.write(line_num + 1, 13, '', style)
    sheet.write(line_num + 1, 14, '', style)
    sheet.write(line_num + 1, 15, '', style)
    sheet.write(line_num + 1, 16, xlwt.Formula('SUM(Q5:Q' + str(line_num) + ')'), style)

    # 报销总额(单位：元）
    style1 = st_style()
    font = xlwt.Font()  # 为样式创建字体
    font.name = '黑体'
    font.height = 20 * 12
    font.bold = True  # 黑体
    style1.font = font

    sheet.write_merge(line_num + 2, line_num + 3, 0, 6, '报销总额(单位：元）', style1)
    sheet.write_merge(line_num + 2, line_num + 2, 7, 8, '人民币', style)

    sheet.write_merge(line_num + 2, line_num + 2, 9, 16, xlwt.Formula(
        'SUM(J' + str(line_num + 1) + ',M' + str(line_num + 1) + ',Q' + str(line_num + 1) + ')'), style1)

    sheet.write_merge(line_num + 3, line_num + 3, 7, 8, '(大写)', style)


    sheet.write_merge(line_num + 3, line_num + 3, 9, 16, xlwt.Formula(
        'SUM(J' + str(line_num + 2) + ',M' + str(line_num + 2) + ',Q' + str(line_num + 2) + ')'), style1)

class ChangesDetails:

    def __init__(self):
        self.wb = xlwt.Workbook(encoding='utf-8')

    def wt_sheet(self, sheet_name='差旅报销模版'):
        return self.wb.add_sheet(sheet_name)

    def wt_data(self, data, file):

        ws = self.wt_sheet('差旅报销')
        style = st_style()
        # 标题
        wt_title(ws)
        # 头
        wt_head(ws)

        for _r in range(len(data)):
            row = data[_r]
            for _c in range(len(row)):
                column = row[_c]
                ws.write(_r + 4, _c, column, style)

        # 尾部
        wt_floor(ws, len(data) + 4)
        self.wb.save(file)


if __name__ == '__main__':
    cd = ChangesDetails()

    data = [['a', 'b', 'c'], ['a1', 'b1', 'c1']]
    cd.wt_data(data, 'Excel_test.xls')
