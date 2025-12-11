# 特許明細書 図面ファイル

このディレクトリには、特許明細書用の図面（Mermaid形式）が含まれています。

## ファイル一覧

- `figure1.mmd` - 図1: 特許出願データ分析装置の構成を示すブロック図
- `figure2.mmd` - 図2: 処理フローを示すフローチャート
- `figure3.mmd` - 図3: データ展開処理の例を示す図

## SVGへの変換方法

### 方法1: Mermaid Live Editor（最も簡単）

1. https://mermaid.live/ にアクセス
2. 各`.mmd`ファイルの内容をコピー
3. エディタに貼り付け
4. 「Actions」メニューから「Export SVG」を選択
5. ダウンロードしたSVGファイルを保存

**手順の詳細:**
```bash
# 図1のSVG生成
cat figure1.mmd | pbcopy  # macOS
cat figure1.mmd | xclip -selection clipboard  # Linux
```
→ https://mermaid.live/ に貼り付け → Export SVG

### 方法2: VS Code拡張機能

1. VS Codeで以下の拡張機能をインストール:
   - **Markdown Preview Mermaid Support**
   - または **Mermaid Markdown Syntax Highlighting**

2. `.mmd`ファイルを開く

3. Markdown Preview を表示（Ctrl+Shift+V）

4. 図を右クリック → "Export to SVG"

### 方法3: コマンドライン（ローカル環境で実行）

お使いのローカルマシンで以下を実行:

```bash
# mermaid-cliをインストール
npm install -g @mermaid-js/mermaid-cli

# 各図をSVGに変換
mmdc -i figure1.mmd -o figure1.svg
mmdc -i figure2.mmd -o figure2.svg
mmdc -i figure3.mmd -o figure3.svg

# 背景色を白に設定してSVG生成
mmdc -i figure1.mmd -o figure1.svg -b white
```

### 方法4: オンラインAPI（自動化）

```bash
# Kroki APIを使用（無料）
curl -X POST "https://kroki.io/mermaid/svg" \
  --data-binary "@figure1.mmd" \
  -o figure1.svg

curl -X POST "https://kroki.io/mermaid/svg" \
  --data-binary "@figure2.mmd" \
  -o figure2.svg

curl -X POST "https://kroki.io/mermaid/svg" \
  --data-binary "@figure3.mmd" \
  -o figure3.svg
```

### 方法5: Python（プログラマティック）

```python
import requests

def mermaid_to_svg(mmd_file, svg_file):
    with open(mmd_file, 'r', encoding='utf-8') as f:
        mermaid_code = f.read()

    response = requests.post(
        'https://kroki.io/mermaid/svg',
        data=mermaid_code.encode('utf-8'),
        headers={'Content-Type': 'text/plain'}
    )

    with open(svg_file, 'wb') as f:
        f.write(response.content)

    print(f"Generated: {svg_file}")

# 実行
mermaid_to_svg('figure1.mmd', 'figure1.svg')
mermaid_to_svg('figure2.mmd', 'figure2.svg')
mermaid_to_svg('figure3.mmd', 'figure3.svg')
```

## SVG品質の最適化

特許出願用に高品質なSVGを生成する場合:

### サイズと解像度の調整

```bash
# 幅を指定してSVG生成
mmdc -i figure1.mmd -o figure1.svg -w 1200

# 高DPI設定
mmdc -i figure1.mmd -o figure1.svg -s 2
```

### SVGからPNGへの変換（必要な場合）

```bash
# Inkscapeを使用（300dpi）
inkscape figure1.svg --export-png=figure1.png --export-dpi=300

# ImageMagickを使用
convert -density 300 figure1.svg figure1.png
```

## 推奨ワークフロー

特許出願用の図面を準備する場合:

1. **SVG生成**: 方法4（Kroki API）が最も簡単で自動化可能
2. **確認**: SVGファイルをブラウザで開いて確認
3. **PNG変換**: 必要に応じて300dpiのPNGに変換
4. **モノクロ確認**: 白黒印刷でも判読可能か確認

## 一括変換スクリプト

すべての図を一度にSVGに変換:

```bash
#!/bin/bash
# convert_all.sh

for mmd in *.mmd; do
    svg="${mmd%.mmd}.svg"
    echo "Converting $mmd to $svg..."
    curl -X POST "https://kroki.io/mermaid/svg" \
      --data-binary "@$mmd" \
      -o "$svg"
done

echo "All diagrams converted to SVG!"
```

実行:
```bash
chmod +x convert_all.sh
./convert_all.sh
```

## 図4について

図4（ヒートマップの表示例）はASCIIアートで作成されています。
SVG化が必要な場合は、別途グラフィックソフトで作成することを推奨します。

## トラブルシューティング

### Kroki APIが利用できない場合
```bash
# 代替: Mermaid Ink API
curl -s "https://mermaid.ink/svg/$(cat figure1.mmd | base64)" > figure1.svg
```

### 日本語フォントの問題
SVGに日本語が含まれる場合、フォント埋め込みが必要な場合があります:
```bash
# Inkscapeでフォント埋め込み
inkscape --export-text-to-path figure1.svg -o figure1_embedded.svg
```

## ライセンス

これらの図面は特許出願用に作成されたものです。
