"""
会計王 仕訳データ作成ツール - サンプルデータ作成スクリプト

サンプルのExcelファイルを作成します。
"""

import pandas as pd
from datetime import datetime, timedelta

# サンプルデータの作成
data = {
    '日付': [
        datetime(2025, 11, 1),
        datetime(2025, 11, 2),
        datetime(2025, 11, 5),
        datetime(2025, 11, 8),
        datetime(2025, 11, 10),
    ],
    '取引内容': [
        '商品売上',
        '消耗品購入',
        'オフィス家賃支払',
        '通信費支払',
        'コンサルティング売上',
    ],
    '金額': [
        110000,
        5500,
        150000,
        12100,
        220000,
    ],
    '科目': [
        '売上高',
        '消耗品費',
        '地代家賃',
        '通信費',
        '売上高',
    ],
    '相手科目': [
        '現金',
        '現金',
        '普通預金',
        '普通預金',
        '売掛金',
    ],
    '備考': [
        'A社への商品販売',
        '事務用品購入',
        '11月分オフィス賃料',
        'インターネット料金',
        'B社コンサルティング',
    ]
}

# DataFrameの作成
df = pd.DataFrame(data)

# Excelファイルとして保存
output_file = 'sample_transaction_data.xlsx'
df.to_excel(output_file, index=False, sheet_name='取引データ')

print(f"✅ サンプルExcelファイルを作成しました: {output_file}")
print(f"📊 データ件数: {len(df)}件")
print("\nこのファイルをアプリにアップロードしてテストしてください。")
