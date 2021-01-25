import pdfplumber


class TripDetails:

    def __init__(self, path='1.pdf'):
        self.path = path

    def extract_tables(self):
        data = [];
        pdf = pdfplumber.open(self.path)

        for page in pdf.pages:
            # 获取当前页面的全部文本信息，包括表格中的文字
            for table in page.extract_tables():
                for i in range(len(table)):
                    # print(row)
                    row = table[i]
                    if i != 0 and len(row) == 9:
                        data.append(row)
                # print('---------- 分割线 ----------')
        pdf.close()
        return data

#
# if __name__ == '__main__':
#     trip = TripDetails('1.pdf');
#     trip.extract_tables()
