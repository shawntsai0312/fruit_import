import pandas as pd

def combineRows(data):
    # 創建DataFrame
    df = pd.DataFrame(data, columns=['arrival_date', 'variety_code', 'pack_code', 'size_code', 'label', 'ctns'])

    # 將ctns列轉換為整數類型
    df['ctns'] = df['ctns'].astype(int)

    # 按指定列分組並對ctns進行求和
    df_grouped = df.groupby(['arrival_date', 'variety_code', 'pack_code', 'size_code', 'label'], as_index=False).sum()

    # 將結果轉換回原始的列表格式
    result = df_grouped.values.tolist()

    return result
