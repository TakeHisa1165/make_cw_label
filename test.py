import xlwings as xw
import sys

path = r"D:\デスクトップ\●カラー、ワッシャ test10面.xlsx"

wb = xw.Book(path)
ws = wb.sheets('029メカテクノ・結城')

ws.range((1, 1), (10000, 6)).color = (255, 255, 255)
ws.range((1, 1), (10000, 6)).clear_contents()

start = 1
n = 1
cnt = 0
k = 1

if start % 2 == 0:
    start_col = 4
    row = start // 2
    if row == 1:
        start_row = 1
        ws.range((start_row, start_col), (start_row+4, start_col+2)).value = ws.range((1, 8), (5, 10)).value
        ws.range((start_row, start_col), (start_row+1, start_col + 2)).color = (145, 210, 40)
        # ws.range((start_row, start_col), (start_row + 1, start_col + 2)).color = (145, 210, 40)
        start_row = start_row + 5
        cnt += 1

        for k in range (start_row, 26, 5):
            for i in range(1, 6, 3):
                ws.range((k, i), (k+4, i+2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k+1, i + 2)).color = (145, 210, 40)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break


    else:
        cal_row = row - 1
        start_row = (cal_row * 5) + 1
        ws.range((start_row, start_col), (start_row+4, start_col+2)).value = ws.range((1, 8), (5, 10)).value
        ws.range((start_row, start_col), (start_row+1, start_col + 2)).color = (145, 210, 40)
        start_row = start_row + 5
        cnt += 1
        for k in range (start_row, 26, 5):
            for i in range(1, 6, 3):
                ws.range((k, i), (k+4, i+2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k+1, i + 2)).color = (145, 210, 40)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break

elif start % 2 != 0:
    start_col = 1
    row = start // 2
    if row == 0:
        start_row = 1
        for k in range(start_row, 26, 5):
            for i in range(start_col, 7, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k + 1, i + 2)).color = (145, 210, 40)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break
    else:
        cal_row = row
        start_row = (cal_row * 5) + 1
        for k in range(start_row, 26, 5):
            for i in range(start_col, 7, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range((k, i), (k + 1, i + 2)).color = (145, 210, 40)
                cnt += 1
                if cnt == n:
                    break
            else:
                continue
            break



