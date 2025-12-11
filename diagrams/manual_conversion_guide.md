# SVG手動変換ガイド

ネットワーク制限やツールのインストールが難しい場合の手動変換手順です。

## 最も簡単な方法：Mermaid Live Editor

### 手順

1. **ブラウザで開く**
   ```
   https://mermaid.live/
   ```

2. **コードをコピー**
   - `figure1.mmd`、`figure2.mmd`、`figure3.mmd`の内容をそれぞれコピー

3. **エディタに貼り付け**
   - 左側のエディタペインにコードを貼り付け
   - 右側にプレビューが自動表示される

4. **SVGをエクスポート**
   - 「Actions」メニューをクリック
   - 「Export SVG」を選択
   - ファイルをダウンロード

5. **ファイル名を変更**
   ```
   ダウンロードファイル → figure1.svg, figure2.svg, figure3.svg
   ```

## 図1のコピー方法

### macOS
```bash
cat figure1.mmd | pbcopy
```

### Linux (X11)
```bash
cat figure1.mmd | xclip -selection clipboard
```

### Windows (PowerShell)
```powershell
Get-Content figure1.mmd | Set-Clipboard
```

### 手動でコピー
```bash
cat figure1.mmd
```
表示された内容を手動でコピー

## オフラインでの変換

### VS Codeを使用（オフライン可能）

1. **拡張機能をインストール**
   - Extension: "Markdown Preview Mermaid Support"
   - または: "Mermaid Markdown Syntax Highlighting"

2. **ファイルを開く**
   ```
   code figure1.mmd
   ```

3. **Markdown Previewを表示**
   - ショートカット: `Ctrl+Shift+V` (Windows/Linux)
   - ショートカット: `Cmd+Shift+V` (macOS)

4. **SVGにエクスポート**
   - 図を右クリック
   - "Copy Image" または "Export to SVG" を選択

## ローカル環境でコマンドライン変換

お使いのローカルマシン（制限のない環境）で実行:

### mermaid-cliを使用

```bash
# インストール
npm install -g @mermaid-js/mermaid-cli

# 変換
mmdc -i figure1.mmd -o figure1.svg -b white
mmdc -i figure2.mmd -o figure2.svg -b white
mmdc -i figure3.mmd -o figure3.svg -b white
```

### オプション付き変換

```bash
# 幅を指定
mmdc -i figure1.mmd -o figure1.svg -w 1200

# 高解像度
mmdc -i figure1.mmd -o figure1.svg -s 2

# 背景色を指定
mmdc -i figure1.mmd -o figure1.svg -b transparent
```

## 品質確認

SVGファイルが正しく生成されたか確認:

### ファイルサイズ
```bash
ls -lh *.svg
```
- 正常: 数KB〜数十KB
- 異常: 0バイトまたは数百バイト

### 内容確認
```bash
head -n 5 figure1.svg
```
以下で始まるべき:
```xml
<svg ...>
```

### ブラウザで開く
```bash
# macOS
open figure1.svg

# Linux
xdg-open figure1.svg

# Windows
start figure1.svg
```

## トラブルシューティング

### 問題: 日本語が文字化けする

**解決方法:**
1. SVGにフォント情報を埋め込む
2. または、Webフォントを参照する

```xml
<!-- SVGファイルに追加 -->
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
  text { font-family: 'Noto Sans JP', sans-serif; }
</style>
```

### 問題: SVGが真っ白

**原因:** 透明背景
**解決方法:** 白背景で再生成

```bash
mmdc -i figure1.mmd -o figure1.svg -b white
```

### 問題: サイズが小さすぎる/大きすぎる

**解決方法:** widthパラメータで調整

```bash
# 幅1200pxで生成
mmdc -i figure1.mmd -o figure1.svg -w 1200

# 幅2400pxで生成（高解像度）
mmdc -i figure1.mmd -o figure1.svg -w 2400
```

## 特許出願用の推奨設定

```bash
# 推奨: 白背景、高解像度
mmdc -i figure1.mmd -o figure1_patent.svg -b white -w 1600 -s 2
```

## SVGからPNGへの変換（必要な場合）

### Inkscape使用
```bash
# 300dpi
inkscape figure1.svg --export-png=figure1.png --export-dpi=300

# 600dpi（高品質）
inkscape figure1.svg --export-png=figure1.png --export-dpi=600
```

### ImageMagick使用
```bash
convert -density 300 -background white figure1.svg figure1.png
```

### オンラインツール
- CloudConvert: https://cloudconvert.com/svg-to-png
- Convertio: https://convertio.co/svg-png/

## まとめ

最も確実で簡単な方法:
1. ✅ **Mermaid Live Editor** (https://mermaid.live/)
2. ✅ **VS Code拡張機能**
3. ✅ **ローカルでmermaid-cli**

これらの方法で確実にSVGファイルを生成できます。
