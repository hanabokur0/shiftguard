#!/usr/bin/env python3
"""
サンプル入力Excelファイルを生成するスクリプト
"""

import argparse
import pandas as pd
from datetime import datetime


def generate_sample(output_path: str, month: str = "2026-02"):
    """
    サンプルの入力Excelファイルを生成
    
    Args:
        output_path: 出力ファイルパス
        month: 対象月 (YYYY-MM形式)
    """
    
    # スタッフシートのサンプルデータ
    staff_data = [
        {
            'staff_id': 'A001',
            'name': '田中太郎',
            'role': '正社員',
            'desired_days': 20,
            'can_day': 1,
            'can_night': 1,
            'can_weekend_holiday': 1,
            'requested_off_dates': '2026-02-11,2026-02-23'
        },
        {
            'staff_id': 'A002',
            'name': '佐藤花子',
            'role': '正社員',
            'desired_days': 18,
            'can_day': 1,
            'can_night': 1,
            'can_weekend_holiday': 1,
            'requested_off_dates': '2026-02-14'
        },
        {
            'staff_id': 'A003',
            'name': '鈴木一郎',
            'role': 'パート',
            'desired_days': 15,
            'can_day': 1,
            'can_night': 0,
            'can_weekend_holiday': 0,
            'requested_off_dates': ''
        },
        {
            'staff_id': 'A004',
            'name': '高橋美咲',
            'role': '契約社員',
            'desired_days': 16,
            'can_day': 1,
            'can_night': 1,
            'can_weekend_holiday': 1,
            'requested_off_dates': '2026-02-20,2026-02-21'
        },
        {
            'staff_id': 'A005',
            'name': '伊藤健太',
            'role': '正社員',
            'desired_days': 22,
            'can_day': 1,
            'can_night': 1,
            'can_weekend_holiday': 1,
            'requested_off_dates': ''
        },
        {
            'staff_id': 'A006',
            'name': '渡辺優子',
            'role': 'パート',
            'desired_days': 12,
            'can_day': 1,
            'can_night': 0,
            'can_weekend_holiday': 1,
            'requested_off_dates': '2026-02-01,2026-02-15'
        },
    ]
    
    # 設定シートのサンプルデータ
    config_data = [{
        'month': month,
        'min_staff_day': 3,
        'min_staff_night': 2,
        'variable_extra_slots_month': 20,  # 変動枠（繁忙・欠員対応）
        'allow_solo_day': 0,  # 日勤ワンオペ許可（0=不可、1=可）
        'allow_solo_night': 0,  # 夜勤ワンオペ許可（0=不可、1=可）
        'max_consecutive_workdays': 6,
        'min_rest_hours': 11,
        'max_month_overtime_hours': 45,
        'standard_day_shift_hours': 8,
        'standard_night_shift_hours': 10
    }]
    
    # DataFrameに変換
    staff_df = pd.DataFrame(staff_data)
    config_df = pd.DataFrame(config_data)
    
    # Excelに書き出し
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        staff_df.to_excel(writer, sheet_name='staff', index=False)
        config_df.to_excel(writer, sheet_name='config', index=False)
    
    print(f"サンプルファイルを生成しました: {output_path}")
    print(f"対象月: {month}")
    print(f"スタッフ数: {len(staff_data)}")
    print("\n次のステップ:")
    print(f"  python shiftguard.py --input {output_path} --output output.xlsx")


def main():
    parser = argparse.ArgumentParser(description='サンプル入力Excelファイルを生成')
    parser.add_argument('--output', '-o', default='sample_input.xlsx', 
                       help='出力ファイル名 (default: sample_input.xlsx)')
    parser.add_argument('--month', '-m', default='2026-02',
                       help='対象月 YYYY-MM形式 (default: 2026-02)')
    
    args = parser.parse_args()
    
    generate_sample(args.output, args.month)


if __name__ == '__main__':
    main()
