import openpyxl as xl
from datetime import datetime
import random

# 現在時刻を取得
now = datetime.now()
formatted_now = now.strftime('%Y%m%d-%H%M%S')

# ファイルを読み込み、新しいシートを作成
wb = xl.load_workbook('exam.xlsx')
ws = wb.create_sheet(title='新しいシート' + str(formatted_now))

# シートに問題を書き込む
for i in range(1, 10):
    num1 = random.randint(-50, 50)
    num2 = random.randint(-50, 50)

    target_cell = ws.cell(i, 1)
    target_cell.value = f'({num1}) + ({num2})'

# シートを保存
wb.save('exam.xlsx')