import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
import re

# ExcelファイルのサポートのためのOPTIONAL IMPORTを追加
try:
    import openpyxl
except ImportError:
    st.warning("openpyxlがインストールされていません。Excelファイルを使用する場合は `pip install openpyxl` でインストールしてください。")

# ページ設定
st.set_page_config(
    page_title="特許出願データ分析ダッシュボード",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# カラーパレット（20色に拡張）
COLORS = [
    '#8dd3c7', '#FFD700', '#bebada', '#fb8072', '#80b1d3',
    '#fdb462', '#b3de69', '#fccde5', '#d9d9d9', '#bc80bd', 
    '#ccebc5', '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
    '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22'
]

# 分類数に応じた色を生成する関数
def get_colors_for_categories(n_categories):
    """分類数に応じて適切な色を返す"""
    if n_categories <= len(COLORS):
        return COLORS[:n_categories]
    else:
        # 20を超える場合は自動生成
        import plotly.colors as pc
        return pc.qualitative.Set3[:n_categories] if n_categories <= 12 else pc.qualitative.Dark24[:n_categories]

def preprocess_data(df):
    """データの前処理を行う"""
    try:
        # S3.1 出願年列追加
        df['year'] = pd.to_datetime(df['出願日']).dt.year
        
        # S3.2 出願人記号除去 & S3.3 出願人分割
        df['出願人/権利者'] = df['出願人/権利者'].str.replace('▲|▼', '', regex=True)
        df['applicants_list'] = df['出願人/権利者'].str.split(',')
        
        # S3.4 FI分割
        df['fi_list'] = df['FI'].fillna('').str.split(r',(?!\d)', regex=True)
        df['fi_list'] = df['fi_list'].apply(lambda x: [item for item in x if item.strip()])
        
        return df
    except Exception as e:
        st.error(f"データ前処理エラー: {str(e)}")
        return None

def expand_data(df):
    """データを展開して集計用データフレームを作成"""
    try:
        # S4.1 出願人展開
        expanded_applicants = []
        for _, row in df.iterrows():
            for applicant in row['applicants_list']:
                new_row = row.copy()
                new_row['出願人/権利者'] = applicant.strip()
                expanded_applicants.append(new_row)
        df_applicants = pd.DataFrame(expanded_applicants)
        
        # S4.2 FI展開
        expanded_fi = []
        for _, row in df.iterrows():
            for fi in row['fi_list']:
                new_row = row.copy()
                new_row['FI'] = fi.strip()
                expanded_fi.append(new_row)
        df_fi = pd.DataFrame(expanded_fi)
        
        # S4.3 FI/出願人展開
        expanded_applicants_fi = []
        for _, row in df.iterrows():
            for applicant in row['applicants_list']:
                for fi in row['fi_list']:
                    new_row = row.copy()
                    new_row['出願人/権利者'] = applicant.strip()
                    new_row['FI'] = fi.strip()
                    expanded_applicants_fi.append(new_row)
        df_applicants_fi = pd.DataFrame(expanded_applicants_fi)
        
        return df_applicants, df_fi, df_applicants_fi
    except Exception as e:
        st.error(f"データ展開エラー: {str(e)}")
        return None, None, None

def aggregate_data(df, df_applicants, df_fi, df_applicants_fi):
    """各種集計を実行"""
    try:
        # S5.1 出願人別集計
        applicant_counts = df_applicants['出願人/権利者'].value_counts().reset_index()
        applicant_counts.columns = ['出願人/権利者', '出願件数']
        
        # S5.2 FI別集計
        fi_counts = df_fi['FI'].value_counts().reset_index()
        fi_counts.columns = ['FI', '出願件数']
        
        # S5.3 出願年別集計
        year_counts = df['year'].value_counts().reset_index()
        year_counts.columns = ['出願年', '出願件数']
        year_counts = year_counts.sort_values('出願年')
        
        # S5.4 年別出願人別集計
        year_applicant_group = df_applicants.groupby(['year', '出願人/権利者']).size().reset_index(name='counts')
        
        # S5.5 年別FI別集計
        year_fi_group = df_fi.groupby(['year', 'FI']).size().reset_index(name='counts')
        
        # S5.6 出願人別FI別集計
        applicant_fi_group = df_applicants_fi.groupby(['出願人/権利者', 'FI']).size().reset_index(name='counts')
        
        # S6.1-6.7 上位データの抽出
        top_applicants = applicant_counts.head(10)
        top_fi = fi_counts.head(10)
        
        # 比率計算
        others_app_count = applicant_counts[10:]['出願件数'].sum() if len(applicant_counts) > 10 else 0
        top_applicant_ratio = top_applicants.copy()
        if others_app_count > 0:
            others_row = pd.DataFrame({'出願人/権利者': ['others'], '出願件数': [others_app_count]})
            top_applicant_ratio = pd.concat([top_applicant_ratio, others_row], ignore_index=True)
        
        others_fi_count = fi_counts[10:]['出願件数'].sum() if len(fi_counts) > 10 else 0
        top_fi_ratio = top_fi.copy()
        if others_fi_count > 0:
            others_row = pd.DataFrame({'FI': ['others'], '出願件数': [others_fi_count]})
            top_fi_ratio = pd.concat([top_fi_ratio, others_row], ignore_index=True)
        
        # トップ10のリスト
        top10_applicants = top_applicants['出願人/権利者'].tolist()
        top10_fi = top_fi['FI'].tolist()
        
        # 年別トップデータ
        year_top_applicant = year_applicant_group[year_applicant_group['出願人/権利者'].isin(top10_applicants)]
        year_top_fi = year_fi_group[year_fi_group['FI'].isin(top10_fi)]
        
        # 出願人/FI上位
        top_applicant_top_fi = applicant_fi_group[
            (applicant_fi_group['出願人/権利者'].isin(top10_applicants)) &
            (applicant_fi_group['FI'].isin(top10_fi))
        ]
        
        return {
            'year_counts': year_counts,
            'applicant_counts': applicant_counts,
            'fi_counts': fi_counts,
            'top_applicants': top_applicants,
            'top_fi': top_fi,
            'top_applicant_ratio': top_applicant_ratio,
            'top_fi_ratio': top_fi_ratio,
            'year_top_applicant': year_top_applicant,
            'year_top_fi': year_top_fi,
            'top_applicant_top_fi': top_applicant_top_fi,
            'top10_applicants': top10_applicants,
            'top10_fi': top10_fi
        }
    except Exception as e:
        st.error(f"データ集計エラー: {str(e)}")
        return None

def create_heatmap_data(data, row_col, col_col, value_col, row_items, col_items):
    """ヒートマップ用のデータマトリックスを作成"""
    # ピボットテーブルを作成
    pivot_data = data.pivot_table(
        values=value_col, 
        index=row_col, 
        columns=col_col, 
        fill_value=0
    )
    
    # 指定された行・列の順序でリオーダー
    pivot_data = pivot_data.reindex(index=row_items, columns=col_items, fill_value=0)
    
    return pivot_data

def plot_yearly_applications(year_counts):
    """年間出願件数推移のグラフ"""
    fig = px.line(year_counts, x='出願年', y='出願件数',
                  title='年間出願件数推移',
                  markers=True)
    fig.update_layout(height=400)
    return fig

def plot_top_applicants_bar(top_applicants):
    """トップ10出願人の横棒グラフ"""
    n_categories = len(top_applicants)
    colors = get_colors_for_categories(n_categories)
    
    fig = px.bar(top_applicants, x='出願件数', y='出願人/権利者',
                 title='トップ10出願人',
                 orientation='h',
                 color_discrete_sequence=colors)
    fig.update_layout(height=400, yaxis={'categoryorder':'total ascending'})
    return fig

def plot_share_chart(data, label_col, value_col, title):
    """シェアの円グラフ"""
    n_categories = len(data)
    colors = get_colors_for_categories(n_categories)
    
    fig = px.pie(data, values=value_col, names=label_col,
                 title=title,
                 color_discrete_sequence=colors)
    fig.update_layout(height=400)
    
    # 20分類に対応してテキストサイズを調整
    if n_categories > 15:
        fig.update_traces(textfont_size=10)
    return fig

def plot_trend_lines(data, x_col, y_col, color_col, title):
    """時系列トレンドの線グラフ"""
    n_categories = len(data[color_col].unique())
    colors = get_colors_for_categories(n_categories)
    
    fig = px.line(data, x=x_col, y=y_col, color=color_col,
                  title=title,
                  markers=True,
                  color_discrete_sequence=colors)
    fig.update_layout(height=500)
    
    # 20分類に対応して凡例を調整
    if n_categories > 15:
        fig.update_layout(
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02,
                font=dict(size=10)
            )
        )
    return fig

def plot_heatmap(matrix_data, title, color_scale='Blues'):
    """ヒートマップの作成（動的な文字色）"""
    # 分類数に応じて高さを調整
    n_rows = len(matrix_data.index)
    n_cols = len(matrix_data.columns)
    height = max(600, n_rows * 30)
    
    # カスタムカラースケールを定義（確実に白→濃い色のグラデーション）
    if color_scale == 'Blues':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(240, 248, 255)'],    # 非常に薄い青
            [0.3, 'rgb(173, 216, 230)'],    # 薄い青
            [0.5, 'rgb(135, 206, 250)'],    # 中程度の青
            [0.7, 'rgb(70, 130, 180)'],     # 濃い青
            [1.0, 'rgb(25, 25, 112)']       # 非常に濃い青
        ]
    elif color_scale == 'Greens':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(240, 255, 240)'],    # 非常に薄い緑
            [0.3, 'rgb(144, 238, 144)'],    # 薄い緑
            [0.5, 'rgb(60, 179, 113)'],     # 中程度の緑
            [0.7, 'rgb(34, 139, 34)'],      # 濃い緑
            [1.0, 'rgb(0, 100, 0)']         # 非常に濃い緑
        ]
    elif color_scale == 'Purples':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(248, 248, 255)'],    # 非常に薄い紫
            [0.3, 'rgb(221, 160, 221)'],    # 薄い紫
            [0.5, 'rgb(186, 85, 211)'],     # 中程度の紫
            [0.7, 'rgb(138, 43, 226)'],     # 濃い紫
            [1.0, 'rgb(75, 0, 130)']        # 非常に濃い紫
        ]
    else:
        # デフォルトは青
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],
            [0.1, 'rgb(240, 248, 255)'],
            [0.3, 'rgb(173, 216, 230)'],
            [0.5, 'rgb(135, 206, 250)'],
            [0.7, 'rgb(70, 130, 180)'],
            [1.0, 'rgb(25, 25, 112)']
        ]
    
    # ヒートマップ作成
    fig = px.imshow(matrix_data, 
                    labels=dict(x="", y="", color="出願件数"),
                    title=title,
                    aspect="auto")
    
    # カスタムカラースケールを適用
    fig.update_traces(
        colorscale=custom_colorscale,
        zmin=0,
        zmax=matrix_data.values.max(),
        showscale=True
    )
    
    # セルに数値を表示（値に応じて動的に色を変更）
    text_values = matrix_data.values
    max_val = matrix_data.values.max() if matrix_data.values.max() > 0 else 1
    
    # テキスト表示用配列を作成
    text_display = []
    text_colors = []
    for row in text_values:
        text_row = []
        color_row = []
        for val in row:
            if val > 0:
                text_row.append(str(int(val)))
                # 値の割合を計算（0-1の範囲）
                ratio = val / max_val
                # 50%以上の値の場合は白文字、それ以下は黒文字
                if ratio > 0.5:
                    color_row.append("white")
                else:
                    color_row.append("black")
            else:
                text_row.append("")
                color_row.append("black")
        text_display.append(text_row)
        text_colors.append(color_row)
    
    # 動的な文字色を適用
    fig.update_traces(
        text=text_display,
        texttemplate="%{text}",
        textfont={"size": 10},
        hovertemplate='行: %{y}<br>列: %{x}<br>出願件数: %{z}<extra></extra>'
    )
    
    # Plotlyのannotationsを使用して個別のセルに色を設定
    annotations = []
    for i, row_label in enumerate(matrix_data.index):
        for j, col_label in enumerate(matrix_data.columns):
            if text_display[i][j]:  # 空でない場合のみ
                annotations.append(
                    dict(
                        x=col_label,
                        y=row_label,
                        text=text_display[i][j],
                        showarrow=False,
                        font=dict(
                            color=text_colors[i][j],
                            size=10 if n_rows <= 15 and n_cols <= 15 else 8
                        ),
                        xref="x",
                        yref="y"
                    )
                )
    
    fig.update_layout(
        height=height,
        annotations=annotations
    )
    
    # 多分類対応でテキストサイズを調整
    if n_rows > 15 or n_cols > 15:
        fig.update_layout(
            xaxis={'tickfont': {'size': 10}},
            yaxis={'tickfont': {'size': 10}}
        )
    
    return fig

def analyze_problem_solution_data(df, df_applicants=None):
    """課題分類・解決手段分類の分析データを生成（オプション機能）"""
    # 課題分類と解決手段分類が存在するかチェック
    if '課題分類' not in df.columns or '解決手段分類' not in df.columns:
        return None
    
    try:
        # 空値を除外
        df_filtered = df.dropna(subset=['課題分類', '解決手段分類'])
        
        # データが十分にあるかチェック
        if len(df_filtered) == 0:
            st.warning("課題分類・解決手段分類のデータが不足しています。")
            return None
        
        # 課題分類の集計
        problem_counts = df_filtered['課題分類'].value_counts().reset_index()
        problem_counts.columns = ['課題分類', '出願件数']
        
        # 解決手段分類の集計
        solution_counts = df_filtered['解決手段分類'].value_counts().reset_index()
        solution_counts.columns = ['解決手段分類', '出願件数']
        
        # 課題×解決手段のクロス集計
        cross_tab = pd.crosstab(df_filtered['課題分類'], df_filtered['解決手段分類'], margins=False)
        
        # 年別課題分類
        if 'year' in df_filtered.columns:
            year_problem = df_filtered.groupby(['year', '課題分類']).size().reset_index(name='counts')
            year_solution = df_filtered.groupby(['year', '解決手段分類']).size().reset_index(name='counts')
        else:
            year_problem = None
            year_solution = None
        
        # 出願人別課題・解決手段（展開後データを使用）
        applicant_problem_cross = None
        applicant_solution_cross = None
        applicant_problem_counts = None
        applicant_solution_counts = None
        top_applicants_for_analysis = None
        
        if (df_applicants is not None and 
            '課題分類' in df_applicants.columns and 
            '解決手段分類' in df_applicants.columns and
            '出願人/権利者' in df_applicants.columns):
            
            # 出願人展開データから課題・解決手段のデータを抽出
            df_app_filtered = df_applicants.dropna(subset=['課題分類', '解決手段分類', '出願人/権利者'])
            
            if len(df_app_filtered) > 0:
                # 出願人別の課題・解決手段集計
                applicant_problem_counts = df_app_filtered.groupby(['出願人/権利者', '課題分類']).size().reset_index(name='counts')
                applicant_solution_counts = df_app_filtered.groupby(['出願人/権利者', '解決手段分類']).size().reset_index(name='counts')
                
                # 上位出願人を特定（分析対象を絞るため、20分類に対応して15出願人に拡張）
                top_applicants = df_app_filtered['出願人/権利者'].value_counts().head(15).index.tolist()
                top_applicants_for_analysis = top_applicants
                
                # 上位出願人のみでクロス集計を作成
                df_top_applicants = df_app_filtered[df_app_filtered['出願人/権利者'].isin(top_applicants)]
                
                if len(df_top_applicants) > 0:
                    # 出願人×課題のクロス集計
                    applicant_problem_cross = pd.crosstab(
                        df_top_applicants['出願人/権利者'], 
                        df_top_applicants['課題分類'], 
                        margins=False
                    )
                    
                    # 出願人×解決手段のクロス集計
                    applicant_solution_cross = pd.crosstab(
                        df_top_applicants['出願人/権利者'], 
                        df_top_applicants['解決手段分類'], 
                        margins=False
                    )
        
        # 分類数を動的に取得
        num_problems = len(problem_counts)
        num_solutions = len(solution_counts)
        
        return {
            'problem_counts': problem_counts,
            'solution_counts': solution_counts,
            'cross_tab': cross_tab,
            'year_problem': year_problem,
            'year_solution': year_solution,
            'applicant_problem_cross': applicant_problem_cross,
            'applicant_solution_cross': applicant_solution_cross,
            'applicant_problem_counts': applicant_problem_counts,
            'applicant_solution_counts': applicant_solution_counts,
            'top_applicants_for_analysis': top_applicants_for_analysis,
            'num_problems': num_problems,
            'num_solutions': num_solutions,
            'total_records': len(df_filtered)
        }
    
    except Exception as e:
        st.warning(f"課題・解決手段分析でエラーが発生しました: {str(e)}")
        return None

def plot_problem_solution_bar(data, x_col, y_col, title, orientation='v'):
    """課題・解決手段の棒グラフ"""
    n_categories = len(data)
    colors = get_colors_for_categories(n_categories)
    
    if orientation == 'h':
        fig = px.bar(data, x=y_col, y=x_col, 
                     title=title, orientation='h',
                     color_discrete_sequence=colors)
        fig.update_layout(height=max(500, n_categories * 25), yaxis={'categoryorder':'total ascending'})
    else:
        fig = px.bar(data, x=x_col, y=y_col,
                     title=title,
                     color_discrete_sequence=colors)
        fig.update_layout(height=500)
        fig.update_xaxis(tickangle=45)
    return fig

def plot_problem_solution_pie(data, names_col, values_col, title):
    """課題・解決手段の円グラフ"""
    n_categories = len(data)
    colors = get_colors_for_categories(n_categories)
    
    fig = px.pie(data, values=values_col, names=names_col,
                 title=title,
                 color_discrete_sequence=colors)
    fig.update_layout(height=500)
    # 20分類に対応してテキストサイズを調整
    if n_categories > 15:
        fig.update_traces(textfont_size=10)
    return fig

def plot_cross_tab_heatmap(cross_tab, title, color_scale='Blues'):
    """課題×解決手段のヒートマップ（動的な文字色）"""
    # 分類数に応じて高さを調整
    n_rows = len(cross_tab.index)
    n_cols = len(cross_tab.columns)
    height = max(600, n_rows * 30)
    
    # カスタムカラースケールを定義（確実に白→濃い色のグラデーション）
    if color_scale == 'Blues':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(240, 248, 255)'],    # 非常に薄い青
            [0.3, 'rgb(173, 216, 230)'],    # 薄い青
            [0.5, 'rgb(135, 206, 250)'],    # 中程度の青
            [0.7, 'rgb(70, 130, 180)'],     # 濃い青
            [1.0, 'rgb(25, 25, 112)']       # 非常に濃い青
        ]
    elif color_scale == 'Oranges':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(255, 245, 238)'],    # 非常に薄いオレンジ
            [0.3, 'rgb(255, 218, 185)'],    # 薄いオレンジ
            [0.5, 'rgb(255, 165, 0)'],      # 中程度のオレンジ
            [0.7, 'rgb(255, 140, 0)'],      # 濃いオレンジ
            [1.0, 'rgb(139, 69, 19)']       # 非常に濃いオレンジ
        ]
    elif color_scale == 'Greens':
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],    # 白
            [0.1, 'rgb(240, 255, 240)'],    # 非常に薄い緑
            [0.3, 'rgb(144, 238, 144)'],    # 薄い緑
            [0.5, 'rgb(60, 179, 113)'],     # 中程度の緑
            [0.7, 'rgb(34, 139, 34)'],      # 濃い緑
            [1.0, 'rgb(0, 100, 0)']         # 非常に濃い緑
        ]
    else:
        # デフォルトは青
        custom_colorscale = [
            [0.0, 'rgb(255, 255, 255)'],
            [0.1, 'rgb(240, 248, 255)'],
            [0.3, 'rgb(173, 216, 230)'],
            [0.5, 'rgb(135, 206, 250)'],
            [0.7, 'rgb(70, 130, 180)'],
            [1.0, 'rgb(25, 25, 112)']
        ]
    
    # ヒートマップ作成
    fig = px.imshow(cross_tab, 
                    labels=dict(x="解決手段分類", y="課題分類", color="出願件数"),
                    title=title,
                    aspect="auto")
    
    # カスタムカラースケールを適用
    fig.update_traces(
        colorscale=custom_colorscale,
        zmin=0,
        zmax=cross_tab.values.max(),
        showscale=True
    )
    
    # セルに数値を表示（値に応じて動的に色を変更）
    text_values = cross_tab.values
    max_val = cross_tab.values.max() if cross_tab.values.max() > 0 else 1
    
    # テキスト表示用配列を作成
    text_display = []
    text_colors = []
    for row in text_values:
        text_row = []
        color_row = []
        for val in row:
            if val > 0:
                text_row.append(str(int(val)))
                # 値の割合を計算（0-1の範囲）
                ratio = val / max_val
                # 50%以上の値の場合は白文字、それ以下は黒文字
                if ratio > 0.5:
                    color_row.append("white")
                else:
                    color_row.append("black")
            else:
                text_row.append("")
                color_row.append("black")
        text_display.append(text_row)
        text_colors.append(color_row)
    
    # 動的な文字色を適用
    fig.update_traces(
        text=text_display,
        texttemplate="%{text}",
        textfont={"size": 10},
        hovertemplate='課題分類: %{y}<br>解決手段分類: %{x}<br>出願件数: %{z}<extra></extra>'
    )
    
    # Plotlyのannotationsを使用して個別のセルに色を設定
    annotations = []
    for i, row_label in enumerate(cross_tab.index):
        for j, col_label in enumerate(cross_tab.columns):
            if text_display[i][j]:  # 空でない場合のみ
                annotations.append(
                    dict(
                        x=col_label,
                        y=row_label,
                        text=text_display[i][j],
                        showarrow=False,
                        font=dict(
                            color=text_colors[i][j],
                            size=10 if n_rows <= 15 and n_cols <= 15 else 8
                        ),
                        xref="x",
                        yref="y"
                    )
                )
    
    fig.update_layout(
        height=height,
        annotations=annotations
    )
    
    # 20分類に対応してテキストサイズを調整
    if n_rows > 15 or n_cols > 15:
        fig.update_layout(
            xaxis={'tickfont': {'size': 10}},
            yaxis={'tickfont': {'size': 10}}
        )
    
    return fig

def plot_year_trend_stacked(data, x_col, y_col, color_col, title):
    """年別トレンドのスタック棒グラフ"""
    n_categories = len(data[color_col].unique())
    colors = get_colors_for_categories(n_categories)
    
    fig = px.bar(data, x=x_col, y=y_col, color=color_col,
                 title=title,
                 color_discrete_sequence=colors)
    fig.update_layout(height=500)
    
    # 20分類に対応して凡例を調整
    if n_categories > 15:
        fig.update_layout(
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02,
                font=dict(size=10)
            )
        )
    return fig

# メイン処理
def main():
    st.title("📊 特許出願データ分析ダッシュボード")
    
    # 説明セクション
    with st.expander("ℹ️ ダッシュボードについて", expanded=True):
        st.markdown("""
        このダッシュボードは、特許出願データを包括的に分析・可視化するためのツールです。CSVファイルまたはExcelファイルをアップロードすることで、以下の分析が可能です：
        
        ### 📊 基本分析（必須）
        - **概要**：出願件数や期間などの基本統計と主要グラフ
        - **時系列分析**：年ごとの出願件数の推移、出願人・FIの時系列変化
        - **ランキング**：出願人およびFIのトップ10と分布状況
        - **ヒートマップ**：出願人/年、FI/年、出願人/FIの相関関係を表示
        
        ### 🎯 課題・解決手段分析（オプション）
        - **課題・解決手段分析**：課題分類と解決手段分類の分布・相関・トレンド分析
        - **出願人別分析**：出願人×課題、出願人×解決手段のクロス集計
        
        ### 📁 必要なファイル形式
        - **基本分析**：出願日、出願人/権利者、FIを含むCSV/Excelファイル
        - **課題・解決手段分析**：上記に加えて課題分類、解決手段分類列が必要（オプション）
        
        ※ 課題分類・解決手段分類列がない場合でも、基本分析は正常に動作します。
        """)
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "CSVまたはExcelファイルをアップロードしてください",
        type=['csv', 'xlsx'],
        help="必須：出願日、出願人/権利者、FI | オプション：課題分類、解決手段分類"
    )
    
    if uploaded_file is not None:
        try:
            # データ読み込み
            with st.spinner('データを読み込み中...'):
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                elif uploaded_file.name.endswith('.xlsx'):
                    df = pd.read_excel(uploaded_file)
                else:
                    st.error("サポートされていないファイル形式です")
                    return
                
            # 必要な列のチェック（基本分析用）
            required_columns = ['出願日', '出願人/権利者', 'FI']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"基本分析に必要な列が見つかりません: {missing_columns}")
                st.write("利用可能な列:", df.columns.tolist())
                return
            
            # オプション分析の利用可能性をチェック
            optional_columns = ['課題分類', '解決手段分類']
            available_optional_columns = [col for col in optional_columns if col in df.columns]
            has_optional_analysis = len(available_optional_columns) == 2
            
            # 利用可能な分析の表示
            st.success("✅ 基本分析（概要、時系列、ランキング、ヒートマップ）が利用可能です")
            
            if has_optional_analysis:
                st.success("✅ 課題・解決手段分析が利用可能です")
            else:
                missing_optional = [col for col in optional_columns if col not in df.columns]
                if missing_optional:
                    st.info(f"ℹ️ 課題・解決手段分析は利用できません（不足列: {missing_optional}）")
                else:
                    st.info("ℹ️ 課題・解決手段分析のデータが不足しています")
            
            # データ前処理
            with st.spinner('データを処理中...'):
                df_processed = preprocess_data(df)
                if df_processed is None:
                    return
                
                df_applicants, df_fi, df_applicants_fi = expand_data(df_processed)
                if df_applicants is None:
                    return
                
                aggregated_data = aggregate_data(df_processed, df_applicants, df_fi, df_applicants_fi)
                if aggregated_data is None:
                    return
                
                # 課題・解決手段分析（利用可能な場合のみ）
                try:
                    problem_solution_data = analyze_problem_solution_data(df_processed, df_applicants)
                    has_problem_solution = problem_solution_data is not None
                except Exception as e:
                    st.warning(f"課題・解決手段分析の処理中にエラーが発生しました: {str(e)}")
                    problem_solution_data = None
                    has_problem_solution = False
            
            # 基本統計の計算
            total_patents = len(df_processed)
            years = df_processed['year'].unique()
            min_year, max_year = int(years.min()), int(years.max())
            year_span = len(years)
            avg_patents_per_year = total_patents // year_span
            unique_fi_count = len(aggregated_data['fi_counts'])
            
            # タブの作成
            if has_problem_solution:
                tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 概要", "📊 時系列分析", "🏆 ランキング", "🔥 ヒートマップ", "🎯 課題・解決手段分析"])
                st.info("💡 全ての分析機能が利用可能です！")
            else:
                tab1, tab2, tab3, tab4 = st.tabs(["📈 概要", "📊 時系列分析", "🏆 ランキング", "🔥 ヒートマップ"])
                st.info("💡 基本分析機能が利用可能です。課題・解決手段分析を利用するには、課題分類・解決手段分類列を含むデータをアップロードしてください。")
            
            # 概要タブ
            with tab1:
                # 統計カード
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("総出願件数", f"{total_patents:,}")
                    st.caption(f"{min_year}年 - {max_year}年")
                
                with col2:
                    st.metric("期間", f"{year_span}年間")
                    st.caption(f"{min_year}年 - {max_year}年")
                
                with col3:
                    st.metric("年平均出願数", f"{avg_patents_per_year:,}")
                    st.caption("期間内の年平均値")
                
                with col4:
                    st.metric("ユニークFI数", f"{unique_fi_count:,}")
                    st.caption("全期間の総数")
                
                st.divider()
                
                # メイングラフ
                col1, col2 = st.columns(2)
                with col1:
                    fig_yearly = plot_yearly_applications(aggregated_data['year_counts'])
                    st.plotly_chart(fig_yearly, use_container_width=True)
                
                with col2:
                    fig_top_app = plot_top_applicants_bar(aggregated_data['top_applicants'])
                    st.plotly_chart(fig_top_app, use_container_width=True)
                
                # シェアグラフ
                col1, col2 = st.columns(2)
                with col1:
                    fig_app_share = plot_share_chart(
                        aggregated_data['top_applicant_ratio'], 
                        '出願人/権利者', '出願件数', 
                        '出願人シェア'
                    )
                    st.plotly_chart(fig_app_share, use_container_width=True)
                
                with col2:
                    fig_fi_share = plot_share_chart(
                        aggregated_data['top_fi_ratio'], 
                        'FI', '出願件数', 
                        'FIシェア'
                    )
                    st.plotly_chart(fig_fi_share, use_container_width=True)
            
            # 時系列分析タブ
            with tab2:
                # 年間出願件数
                fig_yearly_trend = plot_yearly_applications(aggregated_data['year_counts'])
                fig_yearly_trend.update_layout(title='出願年ごとの出願件数')
                st.plotly_chart(fig_yearly_trend, use_container_width=True)
                
                # 出願人トレンド
                fig_app_trend = plot_trend_lines(
                    aggregated_data['year_top_applicant'],
                    'year', 'counts', '出願人/権利者',
                    '出願人/権利者トップ10の年毎の出願件数'
                )
                st.plotly_chart(fig_app_trend, use_container_width=True)
                
                # FIトレンド
                fig_fi_trend = plot_trend_lines(
                    aggregated_data['year_top_fi'],
                    'year', 'counts', 'FI',
                    'FIトップ10の年毎の出願件数'
                )
                st.plotly_chart(fig_fi_trend, use_container_width=True)
            
            # ランキングタブ
            with tab3:
                col1, col2 = st.columns(2)
                
                with col1:
                    fig_app_ranking = plot_top_applicants_bar(aggregated_data['top_applicants'])
                    fig_app_ranking.update_layout(title='トップ10出願人/権利者の出願件数')
                    st.plotly_chart(fig_app_ranking, use_container_width=True)
                    
                    fig_app_share_ranking = plot_share_chart(
                        aggregated_data['top_applicant_ratio'], 
                        '出願人/権利者', '出願件数', 
                        '出願人/権利者別の出願件数の割合'
                    )
                    st.plotly_chart(fig_app_share_ranking, use_container_width=True)
                
                with col2:
                    fig_fi_ranking = px.bar(
                        aggregated_data['top_fi'], 
                        x='出願件数', y='FI',
                        title='トップ10のFIの出願件数',
                        orientation='h',
                        color_discrete_sequence=COLORS
                    )
                    fig_fi_ranking.update_layout(height=400, yaxis={'categoryorder':'total ascending'})
                    st.plotly_chart(fig_fi_ranking, use_container_width=True)
                    
                    fig_fi_share_ranking = plot_share_chart(
                        aggregated_data['top_fi_ratio'], 
                        'FI', '出願件数', 
                        'FI別の出願件数の割合'
                    )
                    st.plotly_chart(fig_fi_share_ranking, use_container_width=True)
            
            # ヒートマップタブ
            with tab4:
                # 出願人-年ヒートマップ
                st.subheader("年ごとの出願人/権利者別特許出願数ヒートマップ")
                years_sorted = sorted(df_processed['year'].unique())
                applicant_year_matrix = create_heatmap_data(
                    aggregated_data['year_top_applicant'],
                    '出願人/権利者', 'year', 'counts',
                    aggregated_data['top10_applicants'], years_sorted
                )
                fig_app_year = plot_heatmap(applicant_year_matrix, '', 'Blues')
                st.plotly_chart(fig_app_year, use_container_width=True)
                
                st.divider()
                
                # FI-年ヒートマップ
                st.subheader("年ごとのFI別特許出願数ヒートマップ")
                fi_year_matrix = create_heatmap_data(
                    aggregated_data['year_top_fi'],
                    'FI', 'year', 'counts',
                    aggregated_data['top10_fi'], years_sorted
                )
                fig_fi_year = plot_heatmap(fi_year_matrix, '', 'Greens')
                st.plotly_chart(fig_fi_year, use_container_width=True)
                
                st.divider()
                
                # 出願人-FIヒートマップ
                st.subheader("出願人とFIに基づく特許出願数ヒートマップ")
                applicant_fi_matrix = create_heatmap_data(
                    aggregated_data['top_applicant_top_fi'],
                    '出願人/権利者', 'FI', 'counts',
                    aggregated_data['top10_applicants'], aggregated_data['top10_fi']
                )
                fig_app_fi = plot_heatmap(applicant_fi_matrix, '', 'Purples')
                st.plotly_chart(fig_app_fi, use_container_width=True)
            
            # 課題・解決手段分析タブ
            if has_problem_solution:
                with tab5:
                    st.header("🎯 課題・解決手段分析")
                    
                    # 基本統計
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric(
                            "課題分類数", 
                            problem_solution_data['num_problems'],
                            help="ユニークな課題分類の数"
                        )
                    with col2:
                        st.metric(
                            "解決手段分類数", 
                            problem_solution_data['num_solutions'],
                            help="ユニークな解決手段分類の数"
                        )
                    with col3:
                        st.metric(
                            "分析対象件数", 
                            problem_solution_data['total_records'],
                            help="課題分類・解決手段分類が記録されている特許件数"
                        )
                    with col4:
                        if problem_solution_data['top_applicants_for_analysis']:
                            st.metric(
                                "分析対象出願人数", 
                                len(problem_solution_data['top_applicants_for_analysis']),
                                help="出願人×課題・解決手段分析で使用するトップ出願人数（最大15）"
                            )
                    
                    st.divider()
                    
                    # 課題分類と解決手段分類の分布
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📋 課題分類の分布")
                        
                        # 棒グラフ
                        fig_problem_bar = plot_problem_solution_bar(
                            problem_solution_data['problem_counts'], 
                            '課題分類', '出願件数', 
                            '課題分類別出願件数', 'h'
                        )
                        st.plotly_chart(fig_problem_bar, use_container_width=True)
                        
                        # 円グラフ
                        fig_problem_pie = plot_problem_solution_pie(
                            problem_solution_data['problem_counts'], 
                            '課題分類', '出願件数', 
                            '課題分類シェア'
                        )
                        st.plotly_chart(fig_problem_pie, use_container_width=True)
                    
                    with col2:
                        st.subheader("🔧 解決手段分類の分布")
                        
                        # 棒グラフ
                        fig_solution_bar = plot_problem_solution_bar(
                            problem_solution_data['solution_counts'], 
                            '解決手段分類', '出願件数', 
                            '解決手段分類別出願件数', 'h'
                        )
                        st.plotly_chart(fig_solution_bar, use_container_width=True)
                        
                        # 円グラフ
                        fig_solution_pie = plot_problem_solution_pie(
                            problem_solution_data['solution_counts'], 
                            '解決手段分類', '出願件数', 
                            '解決手段分類シェア'
                        )
                        st.plotly_chart(fig_solution_pie, use_container_width=True)
                    
                    st.divider()
                    
                    # 課題×解決手段のクロス集計ヒートマップ
                    st.subheader("🎯 課題分類 × 解決手段分類 相関分析")
                    fig_cross = plot_cross_tab_heatmap(
                        problem_solution_data['cross_tab'],
                        '課題分類と解決手段分類の組み合わせ',
                        'Blues'  # 青色グラデーション：白→濃い青
                    )
                    st.plotly_chart(fig_cross, use_container_width=True)
                    
                    # クロス集計の詳細表示
                    with st.expander("📊 クロス集計表の詳細"):
                        st.dataframe(problem_solution_data['cross_tab'], use_container_width=True)
                    
                    # 出願人×課題・解決手段のクロス集計（新機能）
                    if (problem_solution_data['applicant_problem_cross'] is not None and 
                        problem_solution_data['applicant_solution_cross'] is not None):
                        
                        st.divider()
                        st.subheader("🏢 出願人別分析")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.subheader("🏢 × 📋 出願人 × 課題分類")
                            fig_app_problem = plot_cross_tab_heatmap(
                                problem_solution_data['applicant_problem_cross'],
                                '出願人と課題分類の組み合わせ（上位出願人）',
                                'Oranges'  # オレンジ色グラデーション：白→濃いオレンジ
                            )
                            st.plotly_chart(fig_app_problem, use_container_width=True)
                            
                            with st.expander("📊 出願人×課題分類 詳細表"):
                                st.dataframe(problem_solution_data['applicant_problem_cross'], use_container_width=True)
                        
                        with col2:
                            st.subheader("🏢 × 🔧 出願人 × 解決手段分類")
                            fig_app_solution = plot_cross_tab_heatmap(
                                problem_solution_data['applicant_solution_cross'],
                                '出願人と解決手段分類の組み合わせ（上位出願人）',
                                'Greens'  # 緑色グラデーション：白→濃い緑
                            )
                            st.plotly_chart(fig_app_solution, use_container_width=True)
                            
                            with st.expander("📊 出願人×解決手段分類 詳細表"):
                                st.dataframe(problem_solution_data['applicant_solution_cross'], use_container_width=True)
                        
                        # 分析対象出願人の表示
                        if problem_solution_data['top_applicants_for_analysis']:
                            st.info(f"**分析対象出願人**: {', '.join(problem_solution_data['top_applicants_for_analysis'])}")
                    
                    # 年別トレンド（データが利用可能な場合）
                    if problem_solution_data['year_problem'] is not None:
                        st.divider()
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.subheader("📈 年別課題分類トレンド")
                            fig_year_problem = plot_year_trend_stacked(
                                problem_solution_data['year_problem'],
                                'year', 'counts', '課題分類',
                                '年別課題分類の出願件数推移'
                            )
                            st.plotly_chart(fig_year_problem, use_container_width=True)
                        
                        with col2:
                            st.subheader("📈 年別解決手段分類トレンド")
                            fig_year_solution = plot_year_trend_stacked(
                                problem_solution_data['year_solution'],
                                'year', 'counts', '解決手段分類',
                                '年別解決手段分類の出願件数推移'
                            )
                            st.plotly_chart(fig_year_solution, use_container_width=True)
                    
                    # トップ組み合わせの表示（動的な表示数、最大20に拡張）
                    st.divider()
                    max_combinations = min(20, len(problem_solution_data['cross_tab'].values.flatten()))
                    st.subheader(f"🏆 課題×解決手段 人気の組み合わせ Top {max_combinations}")
                    
                    # クロス集計から上位組み合わせを抽出
                    cross_tab_melted = problem_solution_data['cross_tab'].reset_index().melt(
                        id_vars='課題分類', 
                        var_name='解決手段分類', 
                        value_name='出願件数'
                    )
                    top_combinations = cross_tab_melted.sort_values('出願件数', ascending=False).head(max_combinations)
                    top_combinations = top_combinations[top_combinations['出願件数'] > 0]
                    
                    # ランキング表示（20個に対応してコンパクト表示）
                    if len(top_combinations) > 10:
                        # 10個以上の場合は2列表示
                        col_left, col_right = st.columns(2)
                        
                        for i, (_, row) in enumerate(top_combinations.iterrows(), 1):
                            target_col = col_left if i <= len(top_combinations) // 2 + len(top_combinations) % 2 else col_right
                            
                            with target_col:
                                with st.container():
                                    subcol1, subcol2, subcol3 = st.columns([1, 5, 1])
                                    with subcol1:
                                        st.markdown(f"**#{i}**")
                                    with subcol2:
                                        st.markdown(f"**{row['課題分類']}** × **{row['解決手段分類']}**")
                                    with subcol3:
                                        st.markdown(f"**{int(row['出願件数'])}件**")
                    else:
                        # 10個以下の場合は従来の表示
                        for i, (_, row) in enumerate(top_combinations.iterrows(), 1):
                            col1, col2, col3 = st.columns([1, 4, 1])
                            with col1:
                                st.metric("", f"#{i}")
                            with col2:
                                st.write(f"**{row['課題分類']}** × **{row['解決手段分類']}**")
                            with col3:
                                st.metric("件数", f"{int(row['出願件数'])}件")
        
        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")
            st.write("デバッグ情報:")
            if 'df' in locals():
                st.write("データフレームの形状:", df.shape)
                st.write("列名:", df.columns.tolist())
                st.write("最初の5行:")
                st.dataframe(df.head())

if __name__ == "__main__":
    main()