# ビルド

## 自動ビルド (GitHub Actions)

バージョンタグを push すると、GitHub Actions が Windows 用の `DrY.exe` を自動でビルドし、GitHub Release に添付します。

```
git tag v0.0.1
git push origin v0.0.1
```

`v` で始まるタグ (`v0.0.1` など) が対象です。バージョン番号はタグから自動で設定されます。成果物はワークフローの Artifact と GitHub Release の両方から取得できます。

## 手動ビルド例

```
python -m nuitka  --lto=no --standalone --onefile --output-filename=DrY.exe --windows-product-name=DrY --windows-file-description="Billing system for outside cases" --windows-product-version=0.0.1 --windows-company-name="KMC" --windows-icon-from-ico=icon.png main.py
```

# 使い方
1. ファイルをダウンロード
2. DrY.exeファイルと同一フォルダに 、masterフォルダを作成、さらにその中に指定の master.xlsx (非公開)を配置する。
3. 検査抽出ファイル (xlsx)　を DrY.exe 上にドラッグ&ドロップする。
4. コンソールに施設名が表示された場合は、該当する数字を入力した後、Enterを入力。
5. DrY.exeと同一フォルダに、結果ファイルが生成される。
