"""
===============================================================================
📊 CSVデータビジュアルアナライザー & 高度データ解析ダッシュボード
===============================================================================
【プログラム概要】
このアプリケーションは、CSVファイルを読み込み、以下の機能を提供するStreamlitベースのWebアプリです：

 1. 📈 データの図示（インタラクティブな可視化グラフ描画）
    - 折れ線グラフ、棒グラフ、散布図、ヒストグラム、箱ひげ図、円グラフ
 2. 🧪 データの高度分析（統計学的・機械学習・時系列解析）
    - 詳細記述統計（平均・中央値・標準偏差・信頼区間・歪度・尖度など）
    - 相関分析（ピアソン/スピアマンの相関係数・ヒートマップ・有意性p値）
    - 仮説検定・グループ比較（2群比較: t検定/Mann-Whitney U, 多群比較: ANOVA/Kruskal-Wallis）
    - 線形分析・回帰分析（OLS最小二乗法・決定係数R²・残差分析プロット）
    - 時系列分析（移動平均・ボリンジャーバンド・前比成長率・線形トレンド将来予測）

【動作モード】
サイドバーの「作業モードの選択」によって、図の表示のみ、データの分析のみ、または両方を自由に切り替えて作業できます。
===============================================================================
"""

import os
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import statsmodels.api as sm
import statsmodels.formula.api as smf
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_squared_error


# =============================================================================
# 1. ページ基本設定 (画面デザイン・スタイリング)
# =============================================================================
# Streamlitの基本ページレイアウトとタイトル・アイコンを設定
st.set_page_config(
    page_title="CSVデータアナライザー & 高度解析ダッシュボード",
    page_icon="📊",
    layout="wide",                  # 画面全体を広く使うワイドレイアウト
    initial_sidebar_state="expanded" # 起動時にサイドバーを開いた状態にする
)

# UI表示をより洗練させるためのカスタムCSSスタイル定義
st.markdown("""
    <style>
    /* メインコンテンツエリアの上部・下部余白を調整 */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }
    /* メトリクス表示カードの装飾スタイル */
    .stMetric {
        background-color: #f8f9fa;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #4f46e5;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    /* 結果・インサイト用カードのスタイル */
    .insight-card {
        background-color: #f0fdf4;
        border-left: 4px solid #16a34a;
        padding: 14px;
        border-radius: 6px;
        margin-bottom: 1rem;
    }
    /* 説明・情報提示用カードのスタイル */
    .info-card {
        background-color: #eff6ff;
        border-left: 4px solid #2563eb;
        padding: 14px;
        border-radius: 6px;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)


# =============================================================================
# 2. データ読み込み処理 (文字化け自動防止機能付き CSVローダー)
# =============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def load_csv_data(uploaded_file=None, filepath=None):
    """
    CSVファイルを読み込んで Pandas DataFrame に変換する関数です。
    日本語環境でよくある文字化け（文字コードエラー）を防ぐため、
    主要なエンコーディング (UTF-8, Shift-JIS, CP932, EUC-JP) を自動順次試行します。

    Parameters:
        uploaded_file: ユーザーがアップロードしたファイルオブジェクト (BytesIO)
        filepath (str): ローカルで読み込むサンプルファイルのパス

    Returns:
        pd.DataFrame or None: 読み込み成功時はデータフレーム、失敗時は None
    """
    # 試行する文字コードのリスト
    encodings = ['utf-8', 'shift_jis', 'cp932', 'euc-jp']
    
    for enc in encodings:
        try:
            if uploaded_file is not None:
                # ファイルポインタを先頭にリセット
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=enc)
            elif filepath is not None and os.path.exists(filepath):
                df = pd.read_csv(filepath, encoding=enc)
            else:
                return None
            return df
        except (UnicodeDecodeError, pd.errors.ParserError):
            # 文字コード不適合の場合は次の文字コードを試す
            continue
        except Exception:
            return None
    return None


# =============================================================================
# 3. サイドバー制御 (作業モード・データ読み込み・列選択)
# =============================================================================
st.sidebar.title("📊 CSV解析＆データ設定")
st.sidebar.markdown("---")

# -----------------------------------------------------------------------------
# 【設定1】作業モードの切り替え
# -----------------------------------------------------------------------------
st.sidebar.subheader("1. 作業モードの選択")
app_mode = st.sidebar.radio(
    "実行する機能を選択してください:",
    ["📈 図を表示する (データ可視化)", "🧪 データを分析する (高度解析)", "📊 両方を表示 (全機能)"],
    index=0,
    help="「図を表示する」でグラフ描画、「データを分析する」で各種統計・回帰分析を実行できます。"
)

st.sidebar.markdown("---")

# -----------------------------------------------------------------------------
# 【設定2】データソースの選択 (アップロード or サンプルデータ)
# -----------------------------------------------------------------------------
st.sidebar.subheader("2. データソースの選択")
data_source = st.sidebar.radio(
    "読み込むデータを選択してください:",
    ["サンプルデータを使用", "自分のCSVファイルをアップロード"]
)

df = None
if data_source == "自分のCSVファイルをアップロード":
    uploaded_file = st.sidebar.file_uploader(
        "CSVファイルを選択・ドロップしてください",
        type=["csv"],
        help="UTF-8またはShift_JIS形式のCSVファイルに対応しています。"
    )
    if uploaded_file is not None:
        df = load_csv_data(uploaded_file=uploaded_file)
else:
    # ローカルのサンプルCSVデータ（店舗売上・顧客データ）を使用
    sample_path = "sample_data.csv"
    if os.path.exists(sample_path):
        df = load_csv_data(filepath=sample_path)
        st.sidebar.success("サンプルデータ（売上・顧客データ）を読み込みました。")
    else:
        st.sidebar.warning("サンプルデータファイルが見つかりません。")

# データが読み込まれていない場合は処理を停止し案内メッセージを出力
if df is None or df.empty:
    st.title("📊 CSVデータビジュアルアナライザー & 高度解析ダッシュボード")
    st.info("👈 左側のサイドバーからCSVファイルをアップロードするか、サンプルデータを選択してください。")
    st.stop()


# =============================================================================
# 4. データ型の自動識別 & 列選択コントロール (データプロファイリング)
# =============================================================================
# 全項目名、数値列、カテゴリ（文字列）列を分類
all_columns = df.columns.tolist()
numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
categorical_columns = df.select_dtypes(include=['object', 'category']).columns.tolist()

# 日付型として解釈可能な列を自動判別 (60%以上の行が日付として認識できる場合)
date_columns = []
for col in all_columns:
    if col not in numeric_columns:
        try:
            converted = pd.to_datetime(df[col], errors='coerce')
            if converted.notnull().sum() > len(df) * 0.6:
                date_columns.append(col)
        except Exception:
            pass

# -----------------------------------------------------------------------------
# 【設定3】軸・グループ分け列の指定
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("3. グラフ・分析基本列の設定")

# X軸（横軸）列の選択
x_column = st.sidebar.selectbox(
    "X軸（横軸）にする列を選択:",
    options=all_columns,
    index=0,
    help="グラフ描画や分析の基軸となる項目を選択します。"
)

# Y軸（縦軸）列の初期選択（デフォルトで数値列を選択）
default_y = [col for col in numeric_columns if col != x_column][:2]
if not default_y and len(all_columns) > 1:
    default_y = [all_columns[1]]

y_columns = st.sidebar.multiselect(
    "Y軸（縦軸）にする数値列を選択 (複数選択可):",
    options=numeric_columns if numeric_columns else all_columns,
    default=default_y,
    help="グラフの縦軸値や統計・分析の対象となる数値列を選択します。"
)

# グループ分け・色分け列の選択 (オプション)
group_options = ["(なし)"] + categorical_columns + [c for c in all_columns if c not in categorical_columns]
group_column = st.sidebar.selectbox(
    "グループ分け・色分け列 (オプション):",
    options=group_options,
    index=0,
    help="カテゴリごとにグラフを色分けしたり、グループ間比較検定を行う場合に指定します。"
)
color_col = None if group_column == "(なし)" else group_column

# -----------------------------------------------------------------------------
# 【設定4】表示するグラフ種類の選択（図の表示モード時のみ有効）
# -----------------------------------------------------------------------------
show_line = show_bar = show_scatter = show_histogram = show_box = show_pie = False
if app_mode in ["📈 図を表示する (データ可視化)", "📊 両方を表示 (全機能)"]:
    st.sidebar.markdown("---")
    st.sidebar.subheader("4. 表示グラフタイプの選択")
    show_line = st.sidebar.checkbox("📈 折れ線グラフ", value=True)
    show_bar = st.sidebar.checkbox("📊 棒グラフ", value=True)
    show_scatter = st.sidebar.checkbox("🟡 散布図", value=True)
    show_histogram = st.sidebar.checkbox("📶 ヒストグラム", value=False)
    show_box = st.sidebar.checkbox("📦 箱ひげ図", value=False)
    show_pie = st.sidebar.checkbox("🥧 円グラフ", value=False)


# =============================================================================
# 5. メイン画面ヘッダー & データサマリー (データ概要のカード表示)
# =============================================================================
st.title("📊 CSVデータビジュアルアナライザー & 高度解析ダッシュボード")
st.caption(f"現在の選択モード: **{app_mode}**")

# データ全体の規模を示す指標（メトリクス）を表示
m1, m2, m3, m4 = st.columns(4)
m1.metric("総レコード数 (行)", f"{len(df):,} 行")
m2.metric("総項目数 (列)", f"{len(df.columns):,} 列")
m3.metric("数値データの列数", f"{len(numeric_columns)} 列")
m4.metric("検出された日付列数", f"{len(date_columns)} 列")

st.markdown("---")

# 折りたたみ可能なデータプレビューエリア
with st.expander("🔍 データのプレビューと基本データ型を確認する", expanded=False):
    t_prev, t_stat, t_info = st.tabs(["📄 データプレビュー", "📈 基本記述統計サマリー", "ℹ️ 列情報・欠損値チェック"])
    with t_prev:
        st.dataframe(df, use_container_width=True)
    with t_stat:
        if numeric_columns:
            st.dataframe(df.describe().T.style.highlight_max(axis=0), use_container_width=True)
        else:
            st.info("数値形式の列が含まれていません。")
    with t_info:
        info_df = pd.DataFrame({
            "データ型": df.dtypes.astype(str),
            "有効データ数": df.count(),
            "欠損値数": df.isnull().sum(),
            "欠損率 (%)": (df.isnull().sum() / len(df) * 100).round(2)
        })
        st.dataframe(info_df, use_container_width=True)


# =============================================================================
# 6. メイン機能エリア（作業モードに応じた動的タブ切り替え）
# =============================================================================
tab_viz = tab_stats = tab_linear = tab_timeseries = None

# モードに応じて生成するタブを切り替えます
if app_mode == "📈 図を表示する (データ可視化)":
    tabs = st.tabs(["📊 インタラクティブ可視化"])
    tab_viz = tabs[0]
elif app_mode == "🧪 データを分析する (高度解析)":
    tabs = st.tabs(["🧪 統計学的分析", "📐 線形分析 (回帰分析)", "⏳ 時系列分析"])
    tab_stats, tab_linear, tab_timeseries = tabs[0], tabs[1], tabs[2]
else:
    tabs = st.tabs([
        "📊 インタラクティブ可視化", 
        "🧪 統計学的分析", 
        "📐 線形分析 (回帰分析)", 
        "⏳ 時系列分析"
    ])
    tab_viz, tab_stats, tab_linear, tab_timeseries = tabs[0], tabs[1], tabs[2], tabs[3]

# Plotlyの共通グラフテーマ設定
theme_template = "plotly_white"

# -----------------------------------------------------------------------------
# モジュール 1: 📊 インタラクティブ可視化 (図の表示)
# -----------------------------------------------------------------------------
if tab_viz is not None:
    with tab_viz:
        st.header("📈 基本データ可視化ダッシュボード")
        
        if not y_columns:
            st.warning("⚠️ サイドバーの「3. グラフ・分析基本列の設定」から、少なくとも1つのY軸列を選択してください。")
        elif not any([show_line, show_bar, show_scatter, show_histogram, show_box, show_pie]):
            st.info("👈 サイドバーの「4. 表示グラフタイプの選択」から表示したいグラフのチェックボックスをオンにしてください。")
        else:
            # グラフを2列レイアウトで配置
            chart_cols = st.columns(2)
            c_idx = 0
            
            # 【グラフ1】折れ線グラフ (時系列推移や連続変化の可視化)
            if show_line:
                with chart_cols[c_idx % 2]:
                    st.subheader("📈 折れ線グラフ (Line Chart)")
                    fig_line = px.line(
                        df, x=x_column, y=y_columns, color=color_col,
                        title=f"【折れ線】 X: {x_column} vs Y: {', '.join(y_columns)}",
                        template=theme_template, markers=True
                    )
                    fig_line.update_layout(hovermode="x unified")
                    st.plotly_chart(fig_line, use_container_width=True)
                    c_idx += 1

            # 【グラフ2】棒グラフ (項目ごとの数値の大きさ比較)
            if show_bar:
                with chart_cols[c_idx % 2]:
                    st.subheader("📊 棒グラフ (Bar Chart)")
                    fig_bar = px.bar(
                        df, x=x_column, y=y_columns, color=color_col, barmode="group",
                        title=f"【棒グラフ】 X: {x_column} vs Y: {', '.join(y_columns)}",
                        template=theme_template
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                    c_idx += 1

            # 【グラフ3】散布図 (2変数間の相関・関係性の把握)
            if show_scatter:
                with chart_cols[c_idx % 2]:
                    st.subheader("🟡 散布図 (Scatter Plot)")
                    fig_scatter = px.scatter(
                        df, x=x_column, y=y_columns[0], color=color_col,
                        size=y_columns[1] if len(y_columns) >= 2 else None,
                        title=f"【散布図】 X: {x_column} vs Y: {y_columns[0]}",
                        template=theme_template
                    )
                    st.plotly_chart(fig_scatter, use_container_width=True)
                    c_idx += 1

            # 【グラフ4】ヒストグラム (データの度数分布の把握)
            if show_histogram:
                with chart_cols[c_idx % 2]:
                    st.subheader("📶 ヒストグラム (Histogram)")
                    target_col = y_columns[0]
                    fig_hist = px.histogram(
                        df, x=target_col, color=color_col, marginal="box",
                        title=f"【ヒストグラム】 分布: {target_col}",
                        template=theme_template
                    )
                    st.plotly_chart(fig_hist, use_container_width=True)
                    c_idx += 1

            # 【グラフ5】箱ひげ図 (四分位数・外れ値の可視化)
            if show_box:
                with chart_cols[c_idx % 2]:
                    st.subheader("📦 箱ひげ図 (Box Plot)")
                    fig_box = px.box(
                        df, x=x_column if x_column in categorical_columns or df[x_column].nunique() < 15 else None,
                        y=y_columns, color=color_col, points="all",
                        title="【箱ひげ図】 データ分布・四分位数・外れ値",
                        template=theme_template
                    )
                    st.plotly_chart(fig_box, use_container_width=True)
                    c_idx += 1

            # 【グラフ6】円グラフ (全体に対する割合の表示)
            if show_pie:
                with chart_cols[c_idx % 2]:
                    st.subheader("🥧 円グラフ (Pie Chart)")
                    pie_names = color_col if color_col else (x_column if df[x_column].nunique() <= 12 else None)
                    if pie_names:
                        fig_pie = px.pie(
                            df, names=pie_names, values=y_columns[0],
                            title=f"【円グラフ】 {pie_names} 別の {y_columns[0]} 割合",
                            template=theme_template, hole=0.3
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)
                    else:
                        st.info("💡 円グラフ表示には、グループ分け列またはX軸にカテゴリ列（項目数の少ない文字列列）を選択してください。")
                    c_idx += 1


# -----------------------------------------------------------------------------
# モジュール 2: 🧪 統計学的分析 (統計指標・相関・仮説検定)
# -----------------------------------------------------------------------------
if tab_stats is not None:
    with tab_stats:
        st.header("🧪 統計学的分析モジュール")
        st.markdown("記述統計量の計算、変数間の相関係数行列、およびグループ間の統計的仮説検定を実行します。")
        
        st_sub1, st_sub2, st_sub3 = st.tabs(["📈 詳細記述統計量", "🔗 相関分析 (ヒートマップ)", "🧪 仮説検定・グループ比較"])
        
        # -------------------------------------------------------------------------
        # サブ機能 2-1: 詳細記述統計量 (信頼区間・歪度・尖度含む)
        # -------------------------------------------------------------------------
        with st_sub1:
            st.subheader("📊 数値データの詳細統計情報")
            if numeric_columns:
                stats_list = []
                for col in numeric_columns:
                    series = df[col].dropna()
                    n = len(series)
                    mean_val = series.mean()
                    std_val = series.std()
                    median_val = series.median()
                    iqr_val = series.quantile(0.75) - series.quantile(0.25)
                    skew_val = stats.skew(series)
                    kurt_val = stats.kurtosis(series)
                    
                    # 95%信頼区間（t分布に基づく平均値の推定範囲）の計算
                    sem_val = stats.sem(series)
                    ci_range = stats.t.interval(0.95, df=n-1, loc=mean_val, scale=sem_val) if n > 1 else (mean_val, mean_val)
                    
                    stats_list.append({
                        "列名": col,
                        "サンプル数 (N)": n,
                        "平均値": round(mean_val, 3),
                        "中央値 (50%)": round(median_val, 3),
                        "標準偏差 (Std)": round(std_val, 3),
                        "分散 (Var)": round(series.var(), 3),
                        "四分位範囲 (IQR)": round(iqr_val, 3),
                        "歪度 (Skewness)": round(skew_val, 3),
                        "尖度 (Kurtosis)": round(kurt_val, 3),
                        "95% 信頼区間 (下限 - 上限)": f"[{ci_range[0]:.2f}, {ci_range[1]:.2f}]"
                    })
                
                detailed_stats_df = pd.DataFrame(stats_list)
                st.dataframe(detailed_stats_df, use_container_width=True)
                
                st.markdown("""
                **💡 各指標の解釈ポイント:**
                - **歪度 (Skewness)**: 分布の左右の偏り。0に近ければ左右対称、正なら「左寄り（右に長い裾）」、負なら「右寄り（左に長い裾）」。
                - **尖度 (Kurtosis)**: 分布の尖り具合。0に近ければ正規分布、正なら鋭いピーク、負ならなだらかな分布。
                """)
            else:
                st.info("数値型の列が選択されていません。")

        # -------------------------------------------------------------------------
        # サブ機能 2-2: 相関分析 (ピアソン・スピアマン & 有意性判定)
        # -------------------------------------------------------------------------
        with st_sub2:
            st.subheader("🔗 変数間の相関行列とヒートマップ")
            if len(numeric_columns) >= 2:
                corr_method = st.radio("相関係数の種類を選択:", ["ピアソン (積率相関 - 直線関係の強さ)", "スピアマン (順位相関 - 単調関係・外れ値に強い)"], horizontal=True)
                method_code = 'pearson' if 'ピアソン' in corr_method else 'spearman'
                
                # 相関行列の算出
                corr_matrix = df[numeric_columns].corr(method=method_code)
                
                # Plotlyによるヒートマップ描画
                fig_corr = px.imshow(
                    corr_matrix,
                    text_auto=".2f",
                    aspect="auto",
                    color_continuous_scale="RdBu_r",
                    range_color=[-1, 1],
                    title=f"【相関ヒートマップ】 ({corr_method.split(' ')[0]}相関)"
                )
                fig_corr.update_layout(template=theme_template)
                st.plotly_chart(fig_corr, use_container_width=True)
                
                # ペアごとの詳細な相関係数と有意性 (p値) の一覧表示
                st.markdown("##### 📌 変数ペア間の相関係数と統計的有意性")
                pairs = []
                for i in range(len(numeric_columns)):
                    for j in range(i+1, len(numeric_columns)):
                        col1, col2 = numeric_columns[i], numeric_columns[j]
                        val = corr_matrix.loc[col1, col2]
                        valid_df = df[[col1, col2]].dropna()
                        
                        # scipyによるp値の厳密な計算
                        if method_code == 'pearson':
                            _, p_val = stats.pearsonr(valid_df[col1], valid_df[col2])
                        else:
                            _, p_val = stats.spearmanr(valid_df[col1], valid_df[col2])
                        
                        pairs.append({
                            "変数 1": col1,
                            "変数 2": col2,
                            "相関係数": round(val, 3),
                            "相関の強さ": "非常に強い" if abs(val)>=0.7 else ("強い" if abs(val)>=0.4 else "弱い/なし"),
                            "p値 (有意性)": f"{p_val:.4f}",
                            "統計的有意差": "有意 (p < 0.05)" if p_val < 0.05 else "有意差なし (p ≥ 0.05)"
                        })
                
                pairs_df = pd.DataFrame(pairs).sort_values(by="相関係数", key=abs, ascending=False)
                st.dataframe(pairs_df, use_container_width=True)
            else:
                st.info("相関分析を実行するには、少なくとも2つの数値列が必要です。")

        # -------------------------------------------------------------------------
        # サブ機能 2-3: グループ間比較・仮説検定 (t検定 / ANOVA)
        # -------------------------------------------------------------------------
        with st_sub3:
            st.subheader("🧪 グループ間の統計的仮説検定 (t検定 / ANOVA)")
            
            if categorical_columns and numeric_columns:
                col_k1, col_k2 = st.columns(2)
                with col_k1:
                    group_var = st.selectbox("比較するカテゴリ列 (グループ分け):", categorical_columns, index=0)
                with col_k2:
                    target_num_var = st.selectbox("比較対象の数値列:", numeric_columns, index=0)
                
                # グループ一覧の取得
                groups = df[group_var].dropna().unique()
                st.write(f"検出されたグループ数 ({len(groups)} 群): {', '.join(list(map(str, groups)))}")
                
                # 【ケースA】 2群の比較（例: 性別、A/Bテストなど）
                if len(groups) == 2:
                    st.markdown("#### 🔹 2群の比較: Welchのt検定 / Mann-Whitney U検定")
                    g1_data = df[df[group_var] == groups[0]][target_num_var].dropna()
                    g2_data = df[df[group_var] == groups[1]][target_num_var].dropna()
                    
                    # パラメトリック検定 (t検定) とノンパラメトリック検定 (Mann-Whitney U)
                    t_stat, t_pval = stats.ttest_ind(g1_data, g2_data, equal_var=False)
                    u_stat, u_pval = stats.mannwhitneyu(g1_data, g2_data)
                    
                    res_df = pd.DataFrame([
                        {"検定手法": "Welchのt検定 (正規分布前提)", "統計量": round(t_stat, 3), "p値": f"{t_pval:.4f}", "判定": "統計的に有意差あり (p<0.05)" if t_pval < 0.05 else "有意差なし (p≥0.05)"},
                        {"検定手法": "Mann-Whitney U検定 (順位ベース)", "統計量": round(u_stat, 3), "p値": f"{u_pval:.4f}", "判定": "統計的に有意差あり (p<0.05)" if u_pval < 0.05 else "有意差なし (p≥0.05)"}
                    ])
                    st.dataframe(res_df, use_container_width=True)
                    
                    # 統計的インサイト判定文の出力
                    if t_pval < 0.05:
                        st.success(f"💡 **結論:** グループ「{groups[0]}」と「{groups[1]}」の間の【{target_num_var}】には、統計的に有意な平均差が見られます (p = {t_pval:.4f})。")
                    else:
                        st.info(f"💡 **結論:** グループ「{groups[0]}」と「{groups[1]}」の間の【{target_num_var}】には、統計的な有意差は認められませんでした (p = {t_pval:.4f})。")

                # 【ケースB】 3群以上の多群比較（例: 地域別、年代別など）
                elif len(groups) > 2:
                    st.markdown("#### 🔹 多群の比較: 一元配置分散分析 (ANOVA) / Kruskal-Wallis検定")
                    group_data_list = [df[df[group_var] == g][target_num_var].dropna() for g in groups]
                    
                    f_stat, f_pval = stats.f_oneway(*group_data_list)
                    kw_stat, kw_pval = stats.kruskal(*group_data_list)
                    
                    res_df = pd.DataFrame([
                        {"検定手法": "一元配置分散分析 ANOVA (パラメトリック)", "統計量 (F値)": round(f_stat, 3), "p値": f"{f_pval:.4f}", "判定": "いずれかの群に有意差あり (p<0.05)" if f_pval < 0.05 else "有意差なし (p≥0.05)"},
                        {"検定手法": "Kruskal-Wallis検定 (ノンパラメトリック)", "統計量 (H値)": round(kw_stat, 3), "p値": f"{kw_pval:.4f}", "判定": "いずれかの群に有意差あり (p<0.05)" if kw_pval < 0.05 else "有意差なし (p≥0.05)"}
                    ])
                    st.dataframe(res_df, use_container_width=True)
                    
                    if f_pval < 0.05:
                        st.success(f"💡 **結論:** {len(groups)} つのグループ間で【{target_num_var}】の分布に統計的有意差があります (p = {f_pval:.4f})。")
                    else:
                        st.info(f"💡 **結論:** {len(groups)} つのグループ間で【{target_num_var}】に有意差は認められませんでした (p = {f_pval:.4f})。")
                
                # グループごとの分布を箱ひげ図で視覚化
                fig_g_box = px.box(df, x=group_var, y=target_num_var, color=group_var, title=f"【グループ別比較箱ひげ図】 {group_var} ごとの {target_num_var} 分布")
                fig_g_box.update_layout(template=theme_template)
                st.plotly_chart(fig_g_box, use_container_width=True)
            else:
                st.info("グループ比較を実行するには、カテゴリ列と数値列の両方が必要です。")


# -----------------------------------------------------------------------------
# モジュール 3: 📐 線形分析 (OLS回帰分析・要因解析)
# -----------------------------------------------------------------------------
if tab_linear is not None:
    with tab_linear:
        st.header("📐 線形分析・回帰分析モジュール")
        st.markdown("目的変数（Y）と説明変数（X）の関係を最小二乗法（OLS）によりモデル化し、影響度や予測関係を分析します。")
        
        if len(numeric_columns) >= 2:
            reg_col1, reg_col2 = st.columns(2)
            with reg_col1:
                target_y = st.selectbox("🎯 目的変数 (Y - 予測・説明したい変数):", numeric_columns, index=0)
            with reg_col2:
                default_x_list = [c for c in numeric_columns if c != target_y]
                features_x = st.multiselect("🔍 説明変数 (X - 影響を与える変数):", numeric_columns, default=default_x_list[:1])
            
            if features_x:
                # 欠損値を除外したクリーンなデータを作成
                clean_df = df[[target_y] + features_x].dropna()
                
                if len(clean_df) > len(features_x) + 1:
                    Y = clean_df[target_y]
                    X = clean_df[features_x]
                    X_with_const = sm.add_constant(X) # 切片（定数項）を追加
                    
                    # OLS（最小二乗法）線形モデルの学習・当てはめ
                    model = sm.OLS(Y, X_with_const).fit()
                    
                    # 決定係数 R² および 誤差指標の算出
                    r2 = model.rsquared
                    adj_r2 = model.rsquared_adj
                    f_pvalue = model.f_pvalue
                    mse = mean_squared_error(Y, model.fittedvalues)
                    rmse = np.sqrt(mse)
                    
                    # 指標カードを表示
                    rc1, rc2, rc3, rc4 = st.columns(4)
                    rc1.metric("決定係数 (R²)", f"{r2:.4f}", help="モデルがデータをどれくらい説明できているか (0〜1)")
                    rc2.metric("調整済み R²", f"{adj_r2:.4f}", help="説明変数の数を考慮した決定係数")
                    rc3.metric("RMSE (二乗平均平方根誤差)", f"{rmse:.3f}", help="予測誤差の平均的な大きさ")
                    rc4.metric("モデル全体の P値", f"{f_pvalue:.4f}", help="モデル全体が統計的に意味を持つか")
                    
                    st.markdown("---")
                    
                    # 回帰係数パラメータテーブルの表示
                    st.subheader("📋 回帰係数および統計的有意性テーブル")
                    coeff_df = pd.DataFrame({
                        "項目名": model.params.index,
                        "回帰係数 (B)": model.params.values.round(4),
                        "標準誤差": model.bse.values.round(4),
                        "t値": model.tvalues.round(3),
                        "p値 (P>|t|)": model.pvalues.apply(lambda p: f"{p:.4f}"),
                        "95% 信頼区間": [f"[{l:.3f}, {u:.3f}]" for l, u in zip(model.conf_int()[0], model.conf_int()[1])],
                        "有意性": ["有意 (p<0.05)" if p < 0.05 else "有意差なし" for p in model.pvalues]
                    })
                    st.dataframe(coeff_df, use_container_width=True)
                    
                    # 回帰方程式のテキスト自動生成
                    eq_terms = [f"{model.params['const']:.3f}"] if 'const' in model.params else []
                    for feat in features_x:
                        val = model.params[feat]
                        sign = "+" if val >= 0 else "-"
                        eq_terms.append(f"{sign} {abs(val):.3f} × ({feat})")
                    formula_str = f"**{target_y}** = " + " ".join(eq_terms)
                    st.markdown(f"<div class='info-card'><b>推定量（推定された回帰方程式）:</b><br>{formula_str}</div>", unsafe_allow_html=True)
                    
                    # 回帰直線プロットおよび残差プロット
                    st.subheader("📈 回帰直線および残差分析グラフ")
                    g_col1, g_col2 = st.columns(2)
                    
                    with g_col1:
                        if len(features_x) == 1:
                            # 単回帰（説明変数が1つ）の場合：散布図上に回帰直線を描画
                            feat_single = features_x[0]
                            fig_reg = px.scatter(
                                clean_df, x=feat_single, y=target_y, trendline="ols",
                                trendline_color_override="red",
                                title=f"【単回帰直線】 {feat_single} vs {target_y}",
                                template=theme_template
                            )
                            st.plotly_chart(fig_reg, use_container_width=True)
                        else:
                            # 重回帰（説明変数が複数）の場合：実測値 vs 予測値プロット
                            fig_actual_pred = px.scatter(
                                x=model.fittedvalues, y=Y,
                                labels={"x": "モデル予測値", "y": "実測値"},
                                title=f"【実測値 vs 予測値】 {target_y}",
                                template=theme_template
                            )
                            min_v, max_v = min(Y.min(), model.fittedvalues.min()), max(Y.max(), model.fittedvalues.max())
                            fig_actual_pred.add_shape(type="line", x0=min_v, y0=min_v, x1=max_v, y1=max_v, line=dict(color="red", dash="dash"))
                            st.plotly_chart(fig_actual_pred, use_container_width=True)

                    with g_col2:
                        # 残差プロット（予測値に対する残差のばらつき＝モデル適合度の検証）
                        residuals = model.resid
                        fig_res = px.scatter(
                            x=model.fittedvalues, y=residuals,
                            labels={"x": "モデル予測値", "y": "残差 (Residual)"},
                            title="【残差プロット】 予測値 vs 残差",
                            template=theme_template
                        )
                        fig_res.add_hline(y=0, line_dash="dash", line_color="red")
                        st.plotly_chart(fig_res, use_container_width=True)

                else:
                    st.warning("分析を実行するための十分なデータ件数がありません。")
            else:
                st.info("少なくとも1つの説明変数 (X) を選択してください。")
        else:
            st.info("線形分析を実行するには、少なくとも2つの数値列が必要です。")


# -----------------------------------------------------------------------------
# モジュール 4: ⏳ 時系列分析 (移動平均・ボリンジャーバンド・トレンド予測)
# -----------------------------------------------------------------------------
if tab_timeseries is not None:
    with tab_timeseries:
        st.header("⏳ 時系列分析モジュール")
        st.markdown("時間経過に伴うデータの推移、移動平均・ボリンジャーバンド、成長率、将来予測トレンドを分析します。")
        
        # 時系列日付列の選択
        ts_date_col = st.selectbox(
            "📅 日付・時間列を選択:",
            options=all_columns,
            index=all_columns.index(date_columns[0]) if date_columns else 0
        )
        
        ts_target_cols = st.multiselect(
            "📊 時系列分析する数値列を選択:",
            options=numeric_columns,
            default=numeric_columns[:1] if numeric_columns else []
        )
        
        if ts_target_cols:
            # 日付列を datetime 型に変換し、時系列順（昇順）にソート
            ts_df = df.copy()
            ts_df['parsed_date'] = pd.to_datetime(ts_df[ts_date_col], errors='coerce')
            ts_df = ts_df.dropna(subset=['parsed_date']).sort_values('parsed_date').reset_index(drop=True)
            
            if len(ts_df) >= 3:
                target_ts_col = ts_target_cols[0]
                
                st.markdown("---")
                ts_tab1, ts_tab2, ts_tab3 = st.tabs(["📈 移動平均・ボリンジャーバンド", "📊 変動率・累積和", "🔮 トレンド将来予測"])
                
                # -----------------------------------------------------------------
                # サブ機能 4-1: 移動平均 (SMA) & ボリンジャーバンド (±2σ)
                # -----------------------------------------------------------------
                with ts_tab1:
                    st.subheader("📈 移動平均 (Moving Average) と移動標準偏差帯")
                    window_size = st.slider("移動平均のウィンドウサイズ (期間):", min_value=2, max_value=max(3, len(ts_df)//2), value=min(3, len(ts_df)))
                    
                    # 移動平均およびボリンジャーバンドの計算
                    ts_df['sma'] = ts_df[target_ts_col].rolling(window=window_size).mean()
                    ts_df['std'] = ts_df[target_ts_col].rolling(window=window_size).std()
                    ts_df['upper_band'] = ts_df['sma'] + (2 * ts_df['std'])
                    ts_df['lower_band'] = ts_df['sma'] - (2 * ts_df['std'])
                    
                    fig_ma = go.Figure()
                    # 実測値ライン
                    fig_ma.add_trace(go.Scatter(x=ts_df['parsed_date'], y=ts_df[target_ts_col], mode='lines+markers', name='実測値', line=dict(color='blue')))
                    # 移動平均ライン
                    fig_ma.add_trace(go.Scatter(x=ts_df['parsed_date'], y=ts_df['sma'], mode='lines', name=f'{window_size}期移動平均', line=dict(color='orange', width=2)))
                    # 上下バンドライン (±2σ)
                    fig_ma.add_trace(go.Scatter(x=ts_df['parsed_date'], y=ts_df['upper_band'], mode='lines', name='上限バンド (+2σ)', line=dict(color='gray', dash='dash')))
                    fig_ma.add_trace(go.Scatter(x=ts_df['parsed_date'], y=ts_df['lower_band'], mode='lines', name='下限バンド (-2σ)', line=dict(color='gray', dash='dash'), fill='tonexty', fillcolor='rgba(200,200,200,0.2)'))
                    
                    fig_ma.update_layout(title=f"【移動平均 & ボリンジャーバンド】 {target_ts_col}", template=theme_template, hovermode="x unified")
                    st.plotly_chart(fig_ma, use_container_width=True)

                # -----------------------------------------------------------------
                # サブ機能 4-2: 前期比成長率 (% Change) & 累積和 (Cumsum)
                # -----------------------------------------------------------------
                with ts_tab2:
                    st.subheader("📊 前期比成長率 (%) と 累積和 (Cumulative Sum)")
                    
                    ts_df['pct_change'] = ts_df[target_ts_col].pct_change() * 100
                    ts_df['cumsum'] = ts_df[target_ts_col].cumsum()
                    
                    tc1, tc2 = st.columns(2)
                    with tc1:
                        fig_pct = px.bar(
                            ts_df, x='parsed_date', y='pct_change',
                            title=f"【前期比成長率 (%)】 {target_ts_col}",
                            template=theme_template
                        )
                        st.plotly_chart(fig_pct, use_container_width=True)
                    with tc2:
                        fig_cum = px.line(
                            ts_df, x='parsed_date', y='cumsum', markers=True,
                            title=f"【累積和の推移】 {target_ts_col}",
                            template=theme_template
                        )
                        st.plotly_chart(fig_cum, use_container_width=True)

                # -----------------------------------------------------------------
                # サブ機能 4-3: 線形トレンドモデルによる将来予測
                # -----------------------------------------------------------------
                with ts_tab3:
                    st.subheader("🔮 線形トレンド推計による将来予測")
                    forecast_periods = st.slider("将来の予測ステップ数 (期間):", min_value=1, max_value=10, value=5)
                    
                    # 経過時間をインデックス化して線形回帰
                    ts_df['time_idx'] = np.arange(len(ts_df))
                    lr = LinearRegression()
                    lr.fit(ts_df[['time_idx']], ts_df[target_ts_col])
                    
                    # 未来のインデックスを生成して予測
                    last_time_idx = ts_df['time_idx'].iloc[-1]
                    future_time_idx = np.arange(last_time_idx + 1, last_time_idx + 1 + forecast_periods).reshape(-1, 1)
                    future_preds = lr.predict(future_time_idx)
                    
                    # データの日付間隔（日/月/年など）の中央値を推定して未来日付を算出
                    if len(ts_df) > 1:
                        date_diff = ts_df['parsed_date'].diff().median()
                    else:
                        date_diff = pd.Timedelta(days=1)
                    
                    last_date = ts_df['parsed_date'].iloc[-1]
                    future_dates = [last_date + (i + 1) * date_diff for i in range(forecast_periods)]
                    
                    forecast_df = pd.DataFrame({
                        "parsed_date": future_dates,
                        "predicted_val": future_preds
                    })
                    
                    # トレンド予測グラフの描画
                    fig_forecast = go.Figure()
                    fig_forecast.add_trace(go.Scatter(x=ts_df['parsed_date'], y=ts_df[target_ts_col], mode='lines+markers', name='実測値データ', line=dict(color='blue')))
                    fig_forecast.add_trace(go.Scatter(x=ts_df['parsed_date'], y=lr.predict(ts_df[['time_idx']]), mode='lines', name='過去トレンド適合線', line=dict(color='green', dash='dot')))
                    fig_forecast.add_trace(go.Scatter(x=forecast_df['parsed_date'], y=forecast_df['predicted_val'], mode='lines+markers', name='将来予測値', line=dict(color='red', width=3)))
                    
                    fig_forecast.update_layout(title=f"【将来予測トレンド】 {target_ts_col} (+{forecast_periods}ステップ先)", template=theme_template)
                    st.plotly_chart(fig_forecast, use_container_width=True)
                    
                    # 予測値テーブルの表示
                    st.markdown("##### 📋 予測データテーブル")
                    forecast_show_df = forecast_df.copy()
                    forecast_show_df.columns = ["予測日付", "予測値"]
                    forecast_show_df['予測値'] = forecast_show_df['予測値'].round(3)
                    st.dataframe(forecast_show_df, use_container_width=True)

            else:
                st.warning("時系列分析を実行するには、有効な日付・時間データが3行以上必要です。")
        else:
            st.info("時系列分析の対象数値列を選択してください。")


# =============================================================================
# 7. フッター解説 & ガイドライン (Footer)
# =============================================================================
st.markdown("---")
st.markdown("""
### 💡 各機能の使い方と活用ガイド
- **📊 インタラクティブ可視化**: データ型に合わせたインタラクティブグラフ（Plotly）を即座に描画します。
- **🧪 統計学的分析**: 詳細記述統計（平均・信頼区間・歪度等）、相関ヒートマップ、仮説検定（t検定・ANOVA）でデータの有意差を検証します。
- **📐 線形分析 (回帰分析)**: 因子（説明変数）が目的変数に与える影響度を最小二乗法で推計・数式化します。
- **⏳ 時系列分析**: 移動平均によるトレンドの平滑化、ボリンジャーバンド、成長率、および過去推移に基づく将来予測を行えます。
""")
