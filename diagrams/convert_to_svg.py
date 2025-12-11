#!/usr/bin/env python3
"""
Mermaid図をSVGに変換するスクリプト

使用方法:
    python convert_to_svg.py

このスクリプトはKroki APIを使用してMermaid図をSVGに変換します。
"""

import os
import sys
import requests
from pathlib import Path

def convert_mermaid_to_svg(mmd_file, svg_file):
    """
    MermaidファイルをSVGに変換する

    Args:
        mmd_file (str): 入力Mermaidファイル (.mmd)
        svg_file (str): 出力SVGファイル (.svg)

    Returns:
        bool: 成功時True、失敗時False
    """
    try:
        # Mermaidファイルを読み込む
        with open(mmd_file, 'r', encoding='utf-8') as f:
            mermaid_code = f.read()

        print(f"Converting {mmd_file} to {svg_file}...")

        # Kroki APIにPOSTリクエストを送信
        response = requests.post(
            'https://kroki.io/mermaid/svg',
            data=mermaid_code.encode('utf-8'),
            headers={'Content-Type': 'text/plain'},
            timeout=30
        )

        # レスポンスをチェック
        if response.status_code == 200:
            # SVGファイルに保存
            with open(svg_file, 'wb') as f:
                f.write(response.content)
            print(f"✓ Successfully generated: {svg_file}")
            return True
        else:
            print(f"✗ Error: HTTP {response.status_code}")
            print(f"  Response: {response.text[:200]}")
            return False

    except requests.exceptions.RequestException as e:
        print(f"✗ Network error: {e}")
        print(f"\n代替方法:")
        print(f"1. https://mermaid.live/ にアクセス")
        print(f"2. {mmd_file} の内容をコピー＆ペースト")
        print(f"3. 'Export SVG' をクリックしてダウンロード")
        return False
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

def main():
    """メイン処理"""
    print("=" * 60)
    print("Mermaid to SVG Converter")
    print("=" * 60)
    print()

    # 現在のディレクトリ
    current_dir = Path(__file__).parent

    # 変換するファイルのリスト
    mermaid_files = [
        'figure1.mmd',
        'figure2.mmd',
        'figure3.mmd'
    ]

    success_count = 0
    fail_count = 0

    for mmd_file in mermaid_files:
        mmd_path = current_dir / mmd_file
        svg_file = mmd_file.replace('.mmd', '.svg')
        svg_path = current_dir / svg_file

        if not mmd_path.exists():
            print(f"✗ File not found: {mmd_file}")
            fail_count += 1
            continue

        if convert_mermaid_to_svg(mmd_path, svg_path):
            success_count += 1
        else:
            fail_count += 1

        print()

    print("=" * 60)
    print(f"Conversion complete: {success_count} succeeded, {fail_count} failed")
    print("=" * 60)

    if fail_count > 0:
        print("\n【代替方法】ネットワークエラーの場合:")
        print("1. Mermaid Live Editor を使用:")
        print("   https://mermaid.live/")
        print()
        print("2. VS Code 拡張機能を使用:")
        print("   'Markdown Preview Mermaid Support' をインストール")
        print()
        print("3. ローカル環境で mermaid-cli を使用:")
        print("   npm install -g @mermaid-js/mermaid-cli")
        print("   mmdc -i figure1.mmd -o figure1.svg")

    return 0 if fail_count == 0 else 1

if __name__ == '__main__':
    sys.exit(main())
