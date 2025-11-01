import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openai import OpenAI
from dotenv import load_dotenv

# 環境変数の読み込み
load_dotenv()


def convert_with_chatgpt(df, api_key):
    """
    ChatGPT APIを使用してExcelデータを仕訳データに変換
    """
    client = OpenAI(api_key=api_key)
    
    # データフレームをテキスト形式に変換
    data_text = df.head(50).to_string()
    
    # デバッグ用
    st.write("🔍 変換対象データ:")
    st.dataframe(df.head(10))
    
    # プロンプト作成
    prompt = f"""
以下のExcelデータを、会計王の仕訳データ受入形式のCSVに変換してください。

【Excelデータ】
{data_text}

【出力形式】
以下の列を持つCSV形式で出力してください:
- 伝票日付: YYYYMMDD形式
- 伝票番号: 連番
- 借方部門コード: 空欄でOK
- 借方部門名: 空欄でOK
- 借方科目コード: 数字3-4桁
- 借方科目名: 科目名
- 借方補助コード: 空欄でOK
- 借方補助名: 空欄でOK
- 借方税区分: 0=対象外、10=課税売上
- 借方税計算区分: 0=税込、1=税抜
- 借方金額: 数値のみ
- 借方税額: 数値のみ
- 貸方部門コード: 空欄でOK
- 貸方部門名: 空欄でOK
- 貸方科目コード: 数字3-4桁
- 貸方科目名: 科目名
- 貸方補助コード: 空欄でOK
- 貸方補助名: 空欄でOK
- 貸方税区分: 0=対象外、10=課税売上
- 貸方税計算区分: 0=税込、1=税抜
- 貸方金額: 数値のみ
- 貸方税額: 数値のみ
- 摘要: 取引内容の説明

【注意事項】
1. 日付がある場合はYYYYMMDD形式に変換
2. 金額は数値のみにして、カンマや円マークは除去
3. 借方と貸方の金額は必ず一致させる
4. CSVのヘッダー行も出力する
5. 出力はCSV形式のテキストのみで、説明文は不要

上記データを変換してCSVテキストを出力してください。
"""
    
    # ChatGPT API呼び出し
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "あなたは会計データ変換の専門家です。Excelデータを会計王の仕訳データ形式に正確に変換してください。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.3
    )
    
    # レスポンスからCSVテキストを取得
    csv_text = response.choices[0].message.content
    
    # デバッグ用
    st.write("🔍 ChatGPTからの応答:")
    st.code(csv_text[:500] + "..." if len(csv_text) > 500 else csv_text)
    
    # CSVテキストからDataFrameを作成
    # マークダウンのコードブロックを除去
    if "```" in csv_text:
        csv_text = csv_text.split("```")[1]
        if csv_text.startswith("csv"):
            csv_text = csv_text[3:]
    
    csv_text = csv_text.strip()
    
    # StringIOを使ってDataFrameに変換
    df_result = pd.read_csv(io.StringIO(csv_text))
    
    st.write("✅ 変換結果:")
    st.dataframe(df_result)
    
    return df_result


def convert_to_csv(df):
    """
    DataFrameをCSV形式の文字列に変換(Shift-JISエンコード)
    """
    output = io.StringIO()
    df.to_csv(output, index=False, encoding='shift_jis')
    return output.getvalue().encode('shift_jis')


# ページ設定
st.set_page_config(
    page_title="会計王 仕訳データ作成ツール",
    page_icon="📊",
    layout="wide"
)

# タイトル
st.title("📊 会計王 仕訳データ作成ツール")
st.markdown("ExcelファイルをアップロードしてChatGPTで解析し、会計王の仕訳データ受入形式に変換します。")

# セッション状態の初期化
if 'df_journal' not in st.session_state:
    st.session_state.df_journal = None
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None

# サイドバー
with st.sidebar:
    st.header("⚙️ 設定")
    
    # OpenAI APIキーの入力
    api_key = st.text_input(
        "OpenAI APIキー",
        type="password",
        value=os.getenv("OPENAI_API_KEY", ""),
        help="OpenAI APIキーを入力してください。.envファイルでも設定できます。"
    )
    
    st.divider()
    
    st.markdown("""
    ### 📋 使い方
    1. ExcelファイルをアップロードExcelファイルをアップロード
    2. データプレビューを確認
    3. 「ChatGPTで変換」ボタンをクリック
    4. 変換結果を確認・編集
    5. CSVファイルをダウンロード
    """)
    
    st.divider()
    
    st.markdown("""
    ### 📄 会計王データ形式
    - 伝票日付 (YYYYMMDD)
    - 借方科目・補助科目
    - 貸方科目・補助科目
    - 金額・税区分
    - 摘要
    """)

# メインエリア
tab1, tab2, tab3 = st.tabs(["📤 データアップロード", "✏️ データ編集", "📥 ダウンロード"])

with tab1:
    st.header("データのアップロード")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "Excelファイルをアップロードしてください",
        type=['xlsx', 'xls'],
        help="取引データが含まれるExcelファイルを選択してください"
    )
    
    if uploaded_file is not None:
        try:
            # Excelファイルの読み込み
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            st.session_state.uploaded_data = excel_data
            
            # シート選択
            sheet_names = list(excel_data.keys())
            selected_sheet = st.selectbox("シートを選択", sheet_names)
            
            # データプレビュー
            st.subheader("📊 データプレビュー")
            df_preview = excel_data[selected_sheet]
            st.dataframe(df_preview.head(20), use_container_width=True)
            
            st.info(f"行数: {len(df_preview)} / 列数: {len(df_preview.columns)}")
            
            # ChatGPT変換ボタン
            st.divider()
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col2:
                if st.button("🤖 ChatGPTで仕訳データに変換", use_container_width=True, type="primary"):
                    if not api_key:
                        st.error("❌ OpenAI APIキーを入力してください")
                    else:
                        with st.spinner("ChatGPTで変換中..."):
                            try:
                                # ChatGPT APIで変換
                                result_df = convert_with_chatgpt(df_preview, api_key)
                                st.session_state.df_journal = result_df
                                st.success("✅ 変換完了しました! 「データ編集」タブで確認してください。")
                                st.balloons()
                                # rerunの代わりに成功メッセージを表示
                            except Exception as e:
                                st.error(f"❌ 変換エラー: {str(e)}")
                                st.error(f"詳細: {type(e).__name__}")
                                import traceback
                                st.code(traceback.format_exc())
            
        except Exception as e:
            st.error(f"❌ ファイル読み込みエラー: {str(e)}")

with tab2:
    st.header("仕訳データの編集")
    
    if st.session_state.df_journal is not None:
        st.info("💡 データを直接編集できます。編集後は「ダウンロード」タブからCSVファイルを取得してください。")
        
        # データエディタ
        edited_df = st.data_editor(
            st.session_state.df_journal,
            num_rows="dynamic",
            use_container_width=True,
            height=600
        )
        
        # 編集内容を保存
        if st.button("💾 編集内容を保存"):
            st.session_state.df_journal = edited_df
            st.success("✅ 保存しました")
        
        # 統計情報
        st.divider()
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("伝票件数", len(edited_df))
        with col2:
            total_debit = edited_df['借方金額'].sum() if '借方金額' in edited_df.columns else 0
            st.metric("借方合計", f"¥{total_debit:,.0f}")
        with col3:
            total_credit = edited_df['貸方金額'].sum() if '貸方金額' in edited_df.columns else 0
            st.metric("貸方合計", f"¥{total_credit:,.0f}")
            
    else:
        st.warning("⚠️ まだデータが変換されていません。「データアップロード」タブでファイルをアップロードしてください。")

with tab3:
    st.header("CSVファイルのダウンロード")
    
    if st.session_state.df_journal is not None:
        # CSVプレビュー
        st.subheader("📄 エクスポートプレビュー")
        st.dataframe(st.session_state.df_journal, use_container_width=True)
        
        # CSVダウンロード
        csv_data = convert_to_csv(st.session_state.df_journal)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label="📥 会計王形式CSVをダウンロード",
                data=csv_data,
                file_name=f"kaikei_journal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
                type="primary"
            )
        
        st.success("✅ ダウンロードボタンをクリックしてファイルを保存してください")
        
    else:
        st.warning("⚠️ ダウンロードするデータがありません。")


if __name__ == "__main__":
    st.markdown("---")
    st.markdown("© 2025 会計王仕訳データ作成ツール")

