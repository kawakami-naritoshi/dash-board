#!/bin/bash
# Mermaid図をSVGに一括変換するスクリプト
# Kroki APIを使用

echo "========================================"
echo "Mermaid to SVG Batch Converter"
echo "========================================"
echo ""

# カウンター
success=0
fail=0

# 各.mmdファイルを処理
for mmd in *.mmd; do
    if [ -f "$mmd" ]; then
        svg="${mmd%.mmd}.svg"
        echo "Converting $mmd to $svg..."

        if curl -X POST "https://kroki.io/mermaid/svg" \
            --data-binary "@$mmd" \
            -o "$svg" \
            --silent \
            --show-error \
            --max-time 30; then

            # ファイルサイズをチェック（0バイトでないか）
            if [ -s "$svg" ]; then
                echo "✓ Success: $svg"
                ((success++))
            else
                echo "✗ Failed: $svg (empty file)"
                rm "$svg"
                ((fail++))
            fi
        else
            echo "✗ Failed: $svg (curl error)"
            [ -f "$svg" ] && rm "$svg"
            ((fail++))
        fi
        echo ""
    fi
done

echo "========================================"
echo "Conversion complete:"
echo "  Success: $success"
echo "  Failed:  $fail"
echo "========================================"

if [ $fail -gt 0 ]; then
    echo ""
    echo "【代替方法】変換に失敗した場合:"
    echo ""
    echo "方法1: Mermaid Live Editor (推奨)"
    echo "  1. https://mermaid.live/ にアクセス"
    echo "  2. .mmdファイルの内容をコピー＆ペースト"
    echo "  3. 'Export SVG' をクリック"
    echo ""
    echo "方法2: Python スクリプト"
    echo "  python3 convert_to_svg.py"
    echo ""
    echo "方法3: VS Code"
    echo "  'Markdown Preview Mermaid Support' 拡張機能をインストール"
    echo "  .mmdファイルを開いて右クリック → 'Export to SVG'"
    echo ""
fi

exit $fail
