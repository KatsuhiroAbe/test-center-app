import streamlit as st
import pandas as pd
import json
import os
import random

# ==========================================
# 1. 設定とファイルのパス（クラウド対応のため相対パスに変更）
# ==========================================
EXCEL_PATH = "test_center.xlsx"
PROGRESS_FILE = "progress.json"

st.set_page_config(page_title="テストセンター対策アプリ", layout="wide")

# ==========================================
# 学習記録の管理（バックアップ対応）
# ==========================================
def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_progress(progress):
    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
    except:
        pass # クラウド環境での書き込みエラー回避

if 'progress' not in st.session_state:
    st.session_state.progress = load_progress()

if 'state_manager' not in st.session_state:
    st.session_state.state_manager = {}

if 'show_ans' not in st.session_state:
    st.session_state.show_ans = False

def get_status(q_id):
    if q_id not in st.session_state.progress:
        return {"learned": False, "confident": False}
    return st.session_state.progress[q_id]

def update_status(q_id, learned=None, confident=None):
    if q_id not in st.session_state.progress:
        st.session_state.progress[q_id] = {"learned": False, "confident": False}
    if learned is not None:
        st.session_state.progress[q_id]["learned"] = learned
    if confident is not None:
        st.session_state.progress[q_id]["confident"] = confident
    save_progress(st.session_state.progress)

# ==========================================
# 2. Excelデータの読み込み
# ==========================================
@st.cache_data
def load_data():
    if not os.path.exists(EXCEL_PATH):
        return None
    data = {}
    xls = pd.ExcelFile(EXCEL_PATH)
    
    df1 = pd.read_excel(xls, sheet_name="① 語句の意味", header=None)
    skip1 = [58, 136, 166, 168, 172, 182, 226, 306, 428]
    data["①語句の意味"] = extract_questions(df1, 2, 448, skip1, prefix="1_")

    df2 = pd.read_excel(xls, sheet_name="② 語句の用法", header=None)
    skip2 = [17, 63, 78, 82, 104, 111]
    data["②語句の用法"] = extract_questions(df2, 2, 114, skip2, prefix="2_")

    df3 = pd.read_excel(xls, sheet_name="④ 熟語の成り立ち", header=None)
    data["③熟語の成り立ち"] = extract_questions(df3, 2, 52, [], prefix="3_")

    df4 = pd.read_excel(xls, sheet_name="⑤ 空欄補充", header=None)
    a_ranges = list(range(2, 27)) + list(range(113, 137)) + list(range(189, 243))
    b_ranges = list(range(28, 55)) + list(range(244, 256))
    q4_list = []
    for r in a_ranges:
        row_idx = r - 1
        if row_idx < len(df4):
            q4_list.append({
                "id": f"4_A_{r}", "type": "A", "q": df4.iloc[row_idx, 1],
                "choices": [df4.iloc[row_idx, c] for c in range(2, 6) if not pd.isna(df4.iloc[row_idx, c])],
                "ans": df4.iloc[row_idx, 3]
            })
    for r in b_ranges:
        row_idx = r - 1
        if row_idx < len(df4):
            q4_list.append({
                "id": f"4_B_{r}", "type": "B", "q": df4.iloc[row_idx, 1],
                "choices": [df4.iloc[row_idx, 2]] if not pd.isna(df4.iloc[row_idx, 2]) else [],
                "ans": df4.iloc[row_idx, 3]
            })
    data["④空欄補充"] = q4_list

    df5 = pd.read_excel(xls, sheet_name="⑥ 文の並び替え", header=None)
    q5_list = []
    for r in range(2, 16):
        row_idx = r - 1
        if row_idx < len(df5):
            q5_list.append({
                "id": f"5_{r}", "q": df5.iloc[row_idx, 1],
                "choices": [df5.iloc[row_idx, c] for c in range(2, 7) if not pd.isna(df5.iloc[row_idx, c])],
                "ans": df5.iloc[row_idx, 3]
            })
    data["⑤文の並び替え"] = q5_list

    df6 = pd.read_excel(xls, sheet_name="⑧ 単語帳", header=None)
    skip6 = [24, 44, 71, 90, 111, 137, 168, 187, 258, 334, 359, 368, 375]
    q6_list = []
    for r in range(3, 585):
        if r in skip6: continue
        row_idx = r - 1
        if row_idx < len(df6):
            q6_list.append({
                "id": f"6_{r}",
                "q": df6.iloc[row_idx, 1],
                "yomi": df6.iloc[row_idx, 2],
                "ans": df6.iloc[row_idx, 3],
                "desc": df6.iloc[row_idx, 4] if not pd.isna(df6.iloc[row_idx, 4]) else ""
            })
    data["⑥単語帳"] = q6_list

    return data

def extract_questions(df, start_row, end_row, skips, prefix):
    q_list = []
    for r in range(start_row, end_row + 1):
        if r in skips: continue
        row_idx = r - 1
        if row_idx < len(df):
            q_list.append({
                "id": f"{prefix}{r}", "q": df.iloc[row_idx, 1],
                "c": df.iloc[row_idx, 2], "a": df.iloc[row_idx, 3]
            })
    return q_list

# ==========================================
# 3. アプリケーション UI
# ==========================================
def main():
    st.title("テストセンター対策アプリ")
    
    all_data = load_data()
    if all_data is None:
        st.error("エラー: `test_center.xlsx` が同じフォルダに見つかりません。GitHubにアップロードされているか確認してください。")
        return

    # ------------------------------------------
    # URLパラメータからの状態復元（リロード対策）
    # ------------------------------------------
    params = st.query_params
    default_cat = params.get("cat", list(all_data.keys())[0])
    default_mode = params.get("mode", "通常（順番通り）")
    default_idx = int(params.get("idx", 0))

    if default_cat not in all_data:
        default_cat = list(all_data.keys())[0]

    # サイドバー設定
    st.sidebar.header("設定")
    category = st.sidebar.radio("項目を選択", list(all_data.keys()), index=list(all_data.keys()).index(default_cat))
    filter_mode = st.sidebar.radio("出題モード", ["通常（順番通り）", "ランダム出題", "未学習のみ", "自信なしと未学習のみ"], index=["通常（順番通り）", "ランダム出題", "未学習のみ", "自信なしと未学習のみ"].index(default_mode))
    
    # クラウドでのデータ消失を防ぐバックアップ機能
    st.sidebar.divider()
    st.sidebar.subheader("💾 学習履歴の引き継ぎ")
    st.sidebar.caption("※クラウド上では時間が経つと履歴がリセットされるため、定期的にダウンロードして保存してください。")
    
    progress_json = json.dumps(st.session_state.progress, ensure_ascii=False, indent=2)
    st.sidebar.download_button("履歴をダウンロード", data=progress_json, file_name="progress_backup.json", mime="application/json")
    
    uploaded_file = st.sidebar.file_uploader("履歴をアップロード（復元）", type=["json"])
    if uploaded_file is not None:
        try:
            st.session_state.progress = json.load(uploaded_file)
            save_progress(st.session_state.progress)
            st.sidebar.success("復元しました！")
        except:
            st.sidebar.error("ファイルの読み込みに失敗しました。")

    state_key = f"{category}_{filter_mode}"

    # 出題リストの生成
    if state_key not in st.session_state.state_manager:
        filtered_q = []
        for q in all_data[category]:
            status = get_status(q["id"])
            if filter_mode == "未学習のみ":
                if not status["learned"]: filtered_q.append(q)
            elif filter_mode == "自信なしと未学習のみ":
                if not status["learned"] or not status["confident"]: filtered_q.append(q)
            else:
                filtered_q.append(q)
                
        if filter_mode == "ランダム出題":
            random.seed()
            random.shuffle(filtered_q)

        st.session_state.state_manager[state_key] = {
            "playlist": filtered_q,
            "idx": default_idx if state_key == f"{default_cat}_{default_mode}" else 0
        }
        st.session_state.show_ans = False

    st.sidebar.divider()
    if st.sidebar.button("🔄 現在の項目の進捗をリセット", use_container_width=True):
        if state_key in st.session_state.state_manager:
            del st.session_state.state_manager[state_key]
        st.query_params.clear()
        st.rerun()

    playlist_info = st.session_state.state_manager[state_key]
    playlist = playlist_info["playlist"]
    idx = playlist_info["idx"]

    # ------------------------------------------
    # 現在の状態をURLパラメータに保存（F5リロード対策）
    # ------------------------------------------
    st.query_params["cat"] = category
    st.query_params["mode"] = filter_mode
    st.query_params["idx"] = str(idx)

    if not playlist:
        st.info("該当する問題がありません！すべて学習済みか自信ありになっています。")
        return

    if idx >= len(playlist):
        st.success("この出題セットの問題をすべて終了しました！")
        if st.button("リストを更新して最初からやり直す", type="primary"):
            del st.session_state.state_manager[state_key]
            st.query_params.clear()
            st.rerun()
        return

    q_data = playlist[idx]
    q_id = q_data["id"]
    status = get_status(q_id)

    st.write(f"**進捗:** {idx+1} / {len(playlist)} 問目  |  現在の状態: {'🟢学習済み' if status['learned'] else '🔴未学習'} / {'🟢自信あり' if status['confident'] else '🔴自信なし'}")
    st.divider()

    st.subheader("問題")
    st.write(q_data["q"])

    if category in ["①語句の意味", "②語句の用法"] and not pd.isna(q_data.get("c")):
        st.write(f"**選択肢:** {q_data['c']}")
    elif category == "③熟語の成り立ち":
        st.write("**選択肢:**")
        for c in ["似た意味", "反対の意味", "前の字が後を修飾", "動詞＋目的語", "主語＋述語", "どれにも非該当"]:
            st.write(f"- {c}")
    elif category == "④空欄補充":
        if q_data["type"] == "A":
            st.write("**選択肢:**")
            for c in q_data["choices"]: st.write(f"- {c}")
        elif q_data["type"] == "B" and q_data["choices"]:
            st.write(f"**選択肢:** {q_data['choices'][0]}")
    elif category == "⑤文の並び替え":
        st.write("**選択肢:**")
        for c in q_data["choices"]: st.write(f"- {c}")
    elif category == "⑥単語帳" and not pd.isna(q_data.get("yomi")):
        st.write(f"**読み:** {q_data['yomi']}")

    st.write("")

    def proceed_next():
        st.session_state.state_manager[state_key]["idx"] += 1
        st.session_state.show_ans = False
        st.query_params["idx"] = str(st.session_state.state_manager[state_key]["idx"])

    if not st.session_state.show_ans:
        if st.button("解答を表示する", type="primary"):
            st.session_state.show_ans = True
            update_status(q_id, learned=True)
            st.rerun()
    else:
        st.divider()
        st.subheader("正答")
        
        if category in ["①語句の意味", "②語句の用法", "③熟語の成り立ち"]:
            st.success(q_data["a"])
        elif category in ["④空欄補充", "⑤文の並び替え"]:
            st.success(q_data["ans"])
        elif category == "⑥単語帳":
            st.success(f"{q_data['ans']}")
            if q_data["desc"]: st.info(f"補足: {q_data['desc']}")

        st.write("---")
        st.write("この問題の理解度を記録してください：")
        
        col1, col2, col3 = st.columns([1,1,2])
        with col1:
            if st.button("🔴 自信なし", use_container_width=True):
                update_status(q_id, confident=False)
                proceed_next()
                st.rerun()
        with col2:
            if st.button("🟢 自信あり", use_container_width=True):
                update_status(q_id, confident=True)
                proceed_next()
                st.rerun()

if __name__ == "__main__":
    main()