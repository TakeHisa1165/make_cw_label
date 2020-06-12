import xlwings as xw
import sys


class InputToLabel:
    def __init__(self, start_no, no_of_label, path, red, green, blue, sheet_name):
        self.path = path
        self.wb = xw.Book(self.path)
        self.ws = self.wb.sheets(sheet_name)
        self.start = start_no
        self.n = no_of_label
        self.cnt = 0
        self.k = 1
        self.ws.range((1, 1), (10000, 6)).color = (255, 255, 255)
        self.ws.range((1, 1), (10000, 6)).clear_contents()
        self.red = red
        self.green = green
        self.blue = blue
        self.input_to_label()

    def input_to_label(self):
        # 偶数列奇数列の判定　偶数の場合
        if self.start % 2 == 0:
            self.start_col = 4
            self.row = self.start // 2
            # 何段目のラベルを最初に書き込むか？を判定　1段目
            if self.row == 1:
                self.start_row = 1
                self.ws.range((self.start_row, self.start_col), (self.start_row+4, self.start_col+2)).value = self.ws.range((1, 8), (5, 10)).value
                self.ws.range((self.start_row, self.start_col), (self.start_row+1, self.start_col + 2)).color = (self.red, self.green, self.blue)
                self.start_row = self.start_row + 5
                self.cnt += 1

                self.start_even_col(k=self.k, start_row=self.start_row, cnt=self.cnt, n=self.n, ws=self.ws)

            # 何段目のラベルを最初に書き込むか？を判定　1段目以外
            else:
                self.cal_row = self.row - 1
                self.start_row = (self.cal_row * 5) + 1
                self.ws.range((self.start_row, self.start_col), (self.start_row+4, self.start_col+2)).value = self.ws.range((1, 8), (5, 10)).value
                self.ws.range((self.start_row, self.start_col), (self.start_row+1, self.start_col + 2)).color = (self.red, self.green, self.blue)
                self.start_row = self.start_row + 5
                self.cnt += 1

                self.start_even_col(k=self.k, start_row=self.start_row, cnt=self.cnt, n=self.n, ws=self.ws)

        # 偶数列奇数列の判定　奇数の場合
        elif self.start % 2 != 0:
            self.start_col = 1
            self.row = self.start // 2
            # スタート段数の特定　商のみ敬さんするが。0段目はないのでrowが０の場合は１にする
            if self.row == 0:
                self.start_row = 1
                self.start_odd_col(start_row=self.start_row, start_col=self.start_col, cnt=self.cnt, n=self.n,
                                   ws=self.ws)
            # スタート段数の特定　1段目以外
            else:
                self.cal_row = self.row
                self.start_row = (self.cal_row * 5) + 1
                self.start_odd_col(start_row=self.start_row, start_col=self.start_col, cnt=self.cnt, n=self.n,
                                   ws=self.ws)

    def start_even_col(self, k, start_row, cnt, n, ws):
        for k in range(start_row, 26, 5):
            for i in range(1, 6, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k + 1, i + 2)).color = (self.red, self.green, self.blue)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break

    def start_odd_col(self, start_row, start_col, cnt, n, ws):
        for k in range(start_row, 26, 5):
            for i in range(start_col, 7, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k + 1, i + 2)).color = (self.red, self.green, self.blue)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break


if __name__ == '__main__':
    app = InputToLabel()
    app.input_to_label()


