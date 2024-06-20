import os
import csv
def toCSV(cherrys):
    # Create a directory under categorizedCSV/
    if not os.path.exists('categorizedCSV'):
        os.mkdir('categorizedCSV')
    writers = {}  # 用於追蹤每個標籤的CSV寫入器

    for cherry in cherrys:
        # 如果這個日期還沒有對應的CSV寫入器，就創建一個新的
        if cherry.arrival_date not in writers:
            csvfile = open(f'categorizedCSV/{cherry.arrival_date}.csv', 'w', newline='')
            fieldnames = ['公司', '品種', '重量', '大小', '件數']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writers[cherry.arrival_date] = (csvfile, writer)

        # 取得這個日期的CSV寫入器，並寫入資料
        _, writer = writers[cherry.arrival_date]
        writer.writerow({
            '公司': cherry.label,
            '品種': cherry.variety_code,
            '重量': cherry.pack_code,
            '大小': cherry.size_code,
            '件數': cherry.ctns
        })

    # 關閉所有的CSV檔案
    for csvfile, _ in writers.values():
        csvfile.close()