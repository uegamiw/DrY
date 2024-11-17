# ビルド例

```
python -m nuitka  --lto=no --standalone --onefile --windows-product-name=DrY --windows-file-description="Billing system for outside cases" --windows-product-version=0.0.1 --windows-company-name="KMC" --windows-icon-from-ico=icon.png DrY.py
```

# 使い方
1. ファイルをダウンロード
2. DrY.exeファイルと同一フォルダに 、masterフォルダを作成、さらにその中に指定の master.xlsx (非公開)を配置する。
3. 検査抽出ファイル (xlsx)　を DrY.exe 上にドラッグ&ドロップする。
4. コンソールに施設名が表示された場合は、該当する数字を入力した後、Enterを入力。
5. DrY.exeと同一フォルダに、結果ファイルが生成される。
