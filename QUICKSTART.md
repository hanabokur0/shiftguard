# ShiftGuard クイックスタート

## 5分で始める

### 1. インストール（初回のみ）

```bash
# 依存パッケージをインストール
pip install -r requirements.txt
```

### 2. サンプルで試す

```bash
# サンプル入力ファイルを生成
python generate_sample.py

# シフトを生成
python shiftguard.py --input sample_input.xlsx --output result.xlsx

# 結果をExcelで開く
# result.xlsx の「schedule」「warnings」シートを確認
```

### 3. 自分のデータで使う

#### 方法A: サンプルを編集
1. `sample_input.xlsx` を開く
2. `staff` シートでスタッフ情報を編集
3. `config` シートで設定を調整
4. 保存して実行: `python shiftguard.py --input sample_input.xlsx --output my_shift.xlsx`

#### 方法B: 新規作成
1. Excelで新規ファイル作成
2. `staff` シートと `config` シートを作成（仕様は README.md 参照）
3. 実行: `python shiftguard.py --input my_data.xlsx --output result.xlsx`

## よくある最初の質問

**Q: エラーが出る**  
A: `pip install -r requirements.txt` を実行しましたか？

**Q: 警告が多すぎる**  
A: スタッフ数や `can_night` の設定を見直してください。人手不足だと警告が増えます。

**Q: ルールを変えたい**  
A: `rules.yml` を編集してください。

## トラブルシューティング

```bash
# Pythonバージョン確認
python --version  # 3.11以上が必要

# パッケージの再インストール
pip install -r requirements.txt --upgrade

# 詳細なエラー表示
python shiftguard.py --input your_file.xlsx --output result.xlsx 2>&1 | tee error.log
```

## 次のステップ

- README.md で詳細な仕様を確認
- rules.yml でルールをカスタマイズ
- shiftguard.py を読んでロジックを理解
- GitHub で Issue 報告や機能提案

---

問題があれば GitHub Issues へどうぞ！
