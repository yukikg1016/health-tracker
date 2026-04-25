"""
Health Tracker - Screenshot to Excel
起動: streamlit run health_tracker_app.py
"""

import datetime
import os
import pathlib
import sys
import tempfile

import streamlit as st
from dotenv import load_dotenv

BASE_DIR = pathlib.Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

# ── カスタムモジュールのキャッシュを毎回クリア（デプロイ後の古いコード防止）──
for _mn in [k for k in list(sys.modules.keys())
            if k.startswith(("excel_writer", "extractors", "gdrive_helper", "autosleep"))]:
    del sys.modules[_mn]

# .env 読み込み
load_dotenv(BASE_DIR / ".env")

# ── クッキー初期化（モジュールレベルで必須）────────────────────────────────
try:
    from streamlit_cookies_controller import CookieController
    _cookie = CookieController()
except Exception:
    _cookie = None

# ── パスワード保護 ────────────────────────────────────────────────────────────
def _check_password() -> bool:
    correct = os.environ.get("APP_PASSWORD", "yukihealth2026")

    # クッキーで永続ログイン確認
    if _cookie is not None:
        try:
            if _cookie.get("ht_auth") == "ok":
                st.session_state["authenticated"] = True
        except Exception:
            pass

    if st.session_state.get("authenticated"):
        return True

    st.markdown("## 🔐 Health Tracker")
    pw = st.text_input("パスワードを入力してください", type="password", key="pw_input")
    if st.button("ログイン", type="primary"):
        if pw == correct:
            st.session_state["authenticated"] = True
            if _cookie is not None:
                try:
                    _cookie.set("ht_auth", "ok", max_age=30 * 24 * 60 * 60)  # 30日間
                except Exception:
                    pass
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

if not _check_password():
    st.stop()

# クラウド環境かローカルかを判定
_IS_CLOUD = not (pathlib.Path.home() / "　Reasearch" / "Yuki データ.xlsx").exists()
DEFAULT_EXCEL = str(pathlib.Path.home() / "　Reasearch" / "Yuki データ.xlsx")

SHEET_OPTIONS = {
    "🍽️ Nutrition（食事・栄養）":    "nutrition",
    "😴 Sleep（睡眠）":              "sleep",
    "🩸 Labs（血液検査）":            "labs",
    "⚖️ In Body（体組成）":          "inbody",
    "💪 Workout":                    "workout",
    "🏋️ Performance（体力測定）":    "performance",
}

_PORT = os.environ.get("PORT", "8501")
st.set_page_config(page_title="Health Tracker", page_icon="🏃", layout="wide")
st.title("🏃 Health Tracker — Screenshot → Excel")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")
    provider = st.radio("AI Provider", ["gemini", "claude"], horizontal=True, key="provider")

    env_key = os.environ.get("GEMINI_API_KEY" if provider == "gemini" else "ANTHROPIC_API_KEY", "")
    if "api_key_input" not in st.session_state:
        st.session_state["api_key_input"] = env_key

    api_key = st.text_input(
        "APIキー", type="password", key="api_key_input",
        help="環境変数 GEMINI_API_KEY / ANTHROPIC_API_KEY でも設定可能",
    )

    # クラウドモード：Google Drive連携
    if _IS_CLOUD:
        from gdrive_helper import is_configured
        if is_configured():
            st.caption("☁️ Google Drive連携 有効")
        else:
            st.caption("⚠️ Google Drive未設定（SecretsにGDRIVE_FILE_IDとGOOGLE_SERVICE_ACCOUNT_JSONが必要）")
        excel_path = None  # クラウドではパス不要
    else:
        excel_path = st.text_input("Excelファイルパス", value=DEFAULT_EXCEL, key="excel_path_input")

    st.divider()
    st.caption("使い方: 画像をアップロード → 抽出 → 確認 → 書き込み")

# ── Main ──────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1])

with col1:
    sheet_label = st.selectbox("① シート種別", list(SHEET_OPTIONS.keys()))
    sheet_type = SHEET_OPTIONS[sheet_label]

    selected_date = st.date_input("② 日付", value=datetime.date.today())

    if True:
        # 入力方法の切り替え（スマホではカメラが便利）
        input_mode = st.radio(
            "③ 入力方法",
            ["📁 ファイルを選ぶ", "📷 カメラで撮影"],
            horizontal=True,
        )

        uploaded_files = []

        if input_mode == "📁 ファイルを選ぶ":
            files = st.file_uploader(
                "画像をアップロード（複数可・Nutritionは省略可）",
                type=["png", "jpg", "jpeg", "webp"],
                accept_multiple_files=True,
            )
            if files:
                uploaded_files = files

        else:
            # カメラ撮影（1枚ずつ、複数枚追加できるようセッションで管理）
            camera_photo = st.camera_input("カメラで撮影")
            if camera_photo:
                # セッションに蓄積
                if "camera_photos" not in st.session_state:
                    st.session_state["camera_photos"] = []
                # 同じ写真の重複追加を防ぐ
                existing_names = [f.name for f in st.session_state["camera_photos"]]
                if camera_photo.name not in existing_names:
                    st.session_state["camera_photos"].append(camera_photo)

            if "camera_photos" in st.session_state and st.session_state["camera_photos"]:
                uploaded_files = st.session_state["camera_photos"]
                st.caption(f"📷 {len(uploaded_files)} 枚撮影済み")
                cols = st.columns(min(len(uploaded_files), 3))
                for i, f in enumerate(uploaded_files):
                    cols[i % 3].image(f, use_container_width=True)
                if st.button("🗑️ 撮影リセット", use_container_width=True):
                    del st.session_state["camera_photos"]
                    st.rerun()

    # Performance のみ：テキスト入力欄（任意）
    performance_text = ""
    if sheet_type == "performance":
        st.markdown("---")
        performance_text = st.text_area(
            "③ 直接入力（任意）",
            placeholder="例：Push up: 30, Chin up: 12, Squat: 50\n（スクリーンショットがなければここに数字を入力してください）",
            height=80,
        )

    # Nutrition のみ：食事タイプ選択 + 説明欄 + サプリ
    nutrition_description = ""
    nutrition_meal_type = "dinner"
    nutrition_meal_time = ""
    if sheet_type == "nutrition":
        st.markdown("---")
        nutrition_meal_type = st.selectbox(
            "④ 食事タイプ",
            options=["breakfast", "lunch", "dinner", "snacks"],
            format_func=lambda x: {"breakfast": "🌅 Breakfast（朝食）", "lunch": "☀️ Lunch（昼食）",
                                    "dinner": "🌙 Dinner（夕食）", "snacks": "🍎 Snacks（間食）"}[x],
            index=2,
        )
        nutrition_meal_time = st.text_input(
            "⑤ 食事時間（任意）",
            placeholder="例：午後1時、13:00、8:30",
            help="Breakfast time / Lunch time / Dinner time 行に書き込まれます",
        )
        nutrition_description = st.text_area(
            "⑥ 食事の説明（任意）",
            placeholder="例：チキンカレー（ご飯200g、カレールー150g）、サラダ、水\n食材・量・料理名を書くとAIの精度が上がります",
            height=100,
            help="複数画像の場合、この説明は全画像に共通で使われます",
        )
        st.markdown("---")
        st.caption("💊 サプリメント（今日飲んだものをチェック）")
        from excel_writer.nutrition_writer import SUPPLEMENT_NAMES
        supplement_cols = st.columns(2)
        selected_supplements = []
        for i, name in enumerate(SUPPLEMENT_NAMES):
            if supplement_cols[i % 2].checkbox(name, key=f"supp_{name}"):
                selected_supplements.append(name)
        if selected_supplements and st.button("💊 サプリを書き込む", use_container_width=True):
            with st.spinner("書き込み中..."):
                try:
                    if _IS_CLOUD:
                        from gdrive_helper import download_to_temp, upload_from_path
                        _supp_path = download_to_temp()
                    else:
                        _supp_path = excel_path
                    from excel_writer.nutrition_writer import write_supplement_data
                    n = write_supplement_data(selected_supplements, selected_date, _supp_path)
                    if _IS_CLOUD:
                        upload_from_path(_supp_path)
                        pathlib.Path(_supp_path).unlink(missing_ok=True)
                        st.session_state.pop("cloud_excel_tmp_path", None)
                    st.success(f"✅ {n} 件のサプリを書き込みました（{selected_date}）")
                except Exception as e:
                    st.error(f"書き込みエラー: {e}")

    # ファイル選択モードのみプレビュー表示（カメラモードは上で表示済み）
    _input_mode = locals().get("input_mode", "📁 ファイルを選ぶ")
    if uploaded_files and sheet_type != "health_csv" and _input_mode == "📁 ファイルを選ぶ":
        if len(uploaded_files) == 1:
            st.image(uploaded_files[0], caption=uploaded_files[0].name, use_container_width=True)
        else:
            st.caption(f"📎 {len(uploaded_files)} 枚アップロード済み")
            cols = st.columns(min(len(uploaded_files), 3))
            for i, f in enumerate(uploaded_files):
                cols[i % 3].image(f, caption=f.name, use_container_width=True)

with col2:
    # ── 通常フロー（スクリーンショット） ──────────────────────────────────────
    # Nutrition/Performance はテキストのみでも動作可。それ以外は画像必須。
    if not uploaded_files:
        if sheet_type == "nutrition" and not nutrition_description.strip():
            st.info("← 画像をアップロードするか、食事の説明を入力してください")
            st.stop()
        elif sheet_type == "performance" and not performance_text.strip():
            st.info("← スクリーンショットをアップロードするか、上のテキスト欄に数値を入力してください")
            st.stop()
        elif sheet_type not in ("nutrition", "performance"):
            st.info("← スクリーンショットをアップロードしてください（複数可）")
            st.stop()

    if sheet_type == "nutrition":
        extract_btn_label = "⑥ データを抽出する"
    elif sheet_type == "performance":
        extract_btn_label = "④ データを確認する"
    else:
        extract_btn_label = "④ データを抽出する"
    if st.button(extract_btn_label, type="primary", use_container_width=True):
        if not api_key:
            st.error("APIキーを入力してください（サイドバー）")
            st.stop()
        if not _IS_CLOUD and (not excel_path or not pathlib.Path(excel_path).exists()):
            st.error(f"Excelファイルが見つかりません: {excel_path}")
            st.stop()
        if _IS_CLOUD:
            from gdrive_helper import is_configured
            if not is_configured():
                st.error("Google Drive未設定です。SecretsにGDRIVE_FILE_IDとGOOGLE_SERVICE_ACCOUNT_JSONを追加してください")
                st.stop()

        all_results = []

        # Nutrition/Performance で画像なしの場合はテキストのみで1件処理
        if sheet_type in ("nutrition", "performance") and not uploaded_files:
            spinner_msg = "AIがテキストから値を推定中..."
            fake_files = [None]  # ダミー1件
        else:
            spinner_msg = f"AIが {len(uploaded_files)} 枚を解析中..."
            fake_files = uploaded_files

        with st.spinner(spinner_msg):
            for idx, uploaded_file in enumerate(fake_files):
                if uploaded_file is None:
                    # テキストのみモード
                    tmp_path = None
                    fname = "（テキストのみ）"
                else:
                    suffix = pathlib.Path(uploaded_file.name).suffix
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_path = tmp.name
                    fname = uploaded_file.name

                try:
                    if sheet_type == "nutrition":
                        from extractors.nutrition_extractor import extract_nutrition_data
                        data = extract_nutrition_data(tmp_path, provider, api_key,
                                                      description=nutrition_description)
                    elif sheet_type == "sleep":
                        from extractors.sleep_extractor import extract_sleep_data
                        data = extract_sleep_data(tmp_path, provider=provider, api_key=api_key)
                    elif sheet_type == "labs":
                        from extractors.labs_extractor import extract_labs_data
                        data = extract_labs_data(tmp_path, provider, api_key)
                    elif sheet_type == "inbody":
                        from extractors.inbody_extractor import extract_inbody_data
                        data = extract_inbody_data(tmp_path, provider, api_key)
                    elif sheet_type == "workout":
                        from extractors.workout_extractor import extract_workout_data
                        data = extract_workout_data(tmp_path, provider, api_key)
                    elif sheet_type == "performance":
                        from extractors.performance_extractor import extract_performance_data
                        data = extract_performance_data(tmp_path, provider, api_key,
                                                        text=performance_text)

                    all_results.append({"file": fname, "data": data, "status": "ok"})

                except Exception as e:
                    all_results.append({"file": fname, "data": None,
                                        "status": "error", "error": str(e)})
                finally:
                    if tmp_path:
                        pathlib.Path(tmp_path).unlink(missing_ok=True)

        ok = sum(1 for r in all_results if r["status"] == "ok")
        ng = len(all_results) - ok
        if ok:
            st.success(f"解析完了！ {ok} 枚成功" + (f"、{ng} 枚エラー" if ng else ""))
        else:
            st.error("すべての画像でエラーが発生しました")

        if ng:
            for r in all_results:
                if r["status"] == "error":
                    st.warning(f"❌ {r['file']}: {r['error']}")

        st.session_state["all_results"] = all_results
        st.session_state["sheet_type"] = sheet_type
        st.session_state["selected_date"] = selected_date
        st.session_state["excel_path"] = excel_path
        st.session_state["nutrition_description"] = nutrition_description
        st.session_state["nutrition_meal_type"] = nutrition_meal_type
        st.session_state["nutrition_meal_time"] = nutrition_meal_time
        st.session_state["performance_text"] = performance_text

    # ── プレビュー ─────────────────────────────────────────────────────────────
    if "all_results" not in st.session_state:
        st.stop()

    all_results = st.session_state["all_results"]
    sheet_type_saved = st.session_state["sheet_type"]
    date_saved = st.session_state["selected_date"]
    # クラウドモード: プレビュー用にDriveからダウンロード（セッションにキャッシュ）
    if _IS_CLOUD:
        if "cloud_excel_tmp_path" not in st.session_state:
            from gdrive_helper import download_to_temp
            st.session_state["cloud_excel_tmp_path"] = download_to_temp()
        excel_path_saved = st.session_state["cloud_excel_tmp_path"]
    else:
        excel_path_saved = excel_path
    desc_saved = st.session_state.get("nutrition_description", "")
    meal_type_saved = st.session_state.get("nutrition_meal_type", "dinner")
    meal_time_saved = st.session_state.get("nutrition_meal_time", "")

    if sheet_type_saved != sheet_type:
        del st.session_state["all_results"]
        st.info("シートを変更しました。再度抽出してください。")
        st.stop()

    ok_results = [r for r in all_results if r["status"] == "ok"]
    if not ok_results:
        st.stop()

    import pandas as pd

    # Nutrition以外は複数画像をマージして1つのdictにする
    def merge_dicts(dicts: list[dict]) -> dict:
        """複数のdictをマージ。後の画像の値が優先（Noneは上書きしない）"""
        merged = {}
        for d in dicts:
            for k, v in d.items():
                if v is not None:
                    if isinstance(v, dict):
                        merged[k] = merge_dicts([merged.get(k, {}), v])
                    else:
                        merged[k] = v
        return merged

    # 表示・書き込み用データを準備
    if sheet_type == "nutrition":
        # Nutrition: 1枚=1行なので個別に処理
        display_items = [{"label": r["file"], "data": r["data"]} for r in ok_results]
    else:
        # その他: 全画像をマージして1つに
        merged_data = merge_dicts([r["data"] for r in ok_results])
        files = ", ".join(r["file"] for r in ok_results)
        display_items = [{"label": files, "data": merged_data}]
        if len(ok_results) > 1:
            st.info(f"📎 {len(ok_results)} 枚の画像をマージして {date_saved} の列に書き込みます")

    preview_title = "⑦ 抽出結果プレビュー" if sheet_type == "nutrition" else "⑤ 抽出結果プレビュー"
    st.subheader(preview_title)

    for item in display_items:
        data = item["data"]

        if len(display_items) > 1:
            st.markdown(f"**📄 {item['label']}**")

        if sheet_type == "nutrition":
            kcal = data.get("calories_kcal")
            summary = data.get("description_summary", "")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("カロリー", f"{kcal:.0f} kcal" if kcal else "—")
            m2.metric("タンパク質", f"{data.get('protein_g', 0):.1f} g")
            m3.metric("脂質", f"{data.get('fat_g', 0):.1f} g")
            m4.metric("炭水化物", f"{data.get('carbohydrates_g', 0):.1f} g")
            if summary:
                st.caption(f"🍽️ {meal_type_saved.capitalize()}  —  {summary}")

        with st.expander("生データ（JSON）", expanded=False):
            st.json(data)

        try:
            if sheet_type == "nutrition":
                from excel_writer.nutrition_writer import build_nutrition_preview
                import openpyxl
                from excel_writer.writer_base import ExcelWriter
                wb_tmp = openpyxl.load_workbook(excel_path_saved)
                ws_tmp = wb_tmp["Nutrition"]
                col_idx_tmp = ExcelWriter.find_date_column(ws_tmp, date_saved) or (ws_tmp.max_column + 1)
                preview = build_nutrition_preview(data, meal_type_saved, col_idx_tmp)
            elif sheet_type == "sleep":
                from excel_writer.sleep_writer import build_sleep_preview
                import openpyxl
                from excel_writer.writer_base import ExcelWriter
                wb_tmp = openpyxl.load_workbook(excel_path_saved)
                ws_tmp = wb_tmp[ExcelWriter.SHEET_SLEEP]
                col_idx = ExcelWriter.find_date_column(ws_tmp, date_saved) or (ws_tmp.max_column + 1)
                preview = build_sleep_preview(data, col_idx)
            elif sheet_type == "labs":
                from excel_writer.labs_writer import build_labs_preview
                import openpyxl
                from excel_writer.writer_base import ExcelWriter
                wb_tmp = openpyxl.load_workbook(excel_path_saved)
                ws_tmp = wb_tmp[ExcelWriter.SHEET_LABS]
                col_idx = ExcelWriter.find_date_column(ws_tmp, date_saved) or 2
                preview = build_labs_preview(ws_tmp, data.get("results", []), col_idx)
            elif sheet_type == "inbody":
                from excel_writer.inbody_writer import build_inbody_preview
                preview = build_inbody_preview(data, row_idx=2)
            elif sheet_type in ("workout", "performance"):
                from excel_writer.workout_writer import build_workout_preview
                import openpyxl as _opx_pv
                try:
                    _wb_pv = _opx_pv.load_workbook(excel_path_saved)
                    _ws_pv = next((s for n, s in [(n, _wb_pv[n]) for n in _wb_pv.sheetnames]
                                   if n.strip().lower() == "workout"), None)
                except Exception:
                    _ws_pv = None
                preview = build_workout_preview(data, row_idx=2, ws=_ws_pv)

            if sheet_type == "labs":
                df = pd.DataFrame([{k: v for k, v in row.items() if not k.startswith("_")}
                                    for row in preview])
                def highlight_confidence(row):
                    conf = row.get("信頼度", 1.0)
                    if conf == 0.0:
                        return ["background-color: #ffcccc"] * len(row)
                    elif conf < 1.0:
                        return ["background-color: #fff3cd"] * len(row)
                    return [""] * len(row)
                st.dataframe(df.style.apply(highlight_confidence, axis=1), use_container_width=True)
                unmatched = sum(1 for r in preview if not r["_write"])
                if unmatched:
                    st.warning(f"⚠️ {unmatched} 項目がマッチせず書き込みスキップされます")
            else:
                df = pd.DataFrame([{k: v for k, v in row.items() if not k.startswith("_")}
                                    for row in preview])
                st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.warning(f"プレビュー生成中のエラー: {e}")

        if len(display_items) > 1:
            st.divider()

    # ── 書き込みボタン ─────────────────────────────────────────────────────────
    n_label = f"（{len(ok_results)} 枚分）" if len(ok_results) > 1 else ""
    write_btn_label = f"{'⑧' if sheet_type == 'nutrition' else '⑥'} Excelに書き込む ✅{n_label}"

    if st.button(write_btn_label, type="primary", use_container_width=True):
        from excel_writer.writer_base import ExcelFileLockError

        with st.spinner("Excelに書き込み中..."):
            try:
                # クラウドモード: Google DriveからExcelをダウンロードして処理
                if _IS_CLOUD:
                    from gdrive_helper import download_to_temp
                    excel_path_saved = download_to_temp()

                if sheet_type == "nutrition":
                    from excel_writer.nutrition_writer import write_nutrition_data
                    merged_nutrition = merge_dicts([r["data"] for r in ok_results])
                    write_nutrition_data(merged_nutrition, date_saved, meal_type_saved, excel_path_saved,
                                         meal_time=meal_time_saved)
                    meal_label = {"breakfast": "Breakfast", "lunch": "Lunch",
                                  "dinner": "Dinner", "snacks": "Snacks"}.get(meal_type_saved, meal_type_saved)
                    st.success(f"✅ Nutrition シート [{meal_label}] に書き込みました！（{date_saved}）")
                else:
                    merged = display_items[0]["data"]
                    if sheet_type == "sleep":
                        from excel_writer.sleep_writer import write_sleep_data
                        write_sleep_data(merged, date_saved, excel_path_saved)
                    elif sheet_type == "labs":
                        from excel_writer.labs_writer import write_labs_data
                        write_labs_data(merged.get("results", []), date_saved, excel_path_saved)
                    elif sheet_type == "inbody":
                        from excel_writer.inbody_writer import write_inbody_data
                        write_inbody_data(merged, date_saved, excel_path_saved)
                    elif sheet_type in ("workout", "performance"):
                        from excel_writer.workout_writer import write_workout_data, get_column_a_dump
                        # Show column A structure for debugging row placement
                        try:
                            col_a = get_column_a_dump(excel_path_saved)
                            if col_a:
                                with st.expander("🔍 Workout sheet 行構造（デバッグ）", expanded=False):
                                    import pandas as pd
                                    st.dataframe(pd.DataFrame(col_a, columns=["行", "列A", "列B"]),
                                                 use_container_width=True)
                        except Exception as _de:
                            st.caption(f"デバッグ情報取得失敗: {_de}")
                        for r in ok_results:
                            dbg = write_workout_data(r["data"], date_saved, excel_path_saved)
                            st.caption(f"✍️ {dbg.get('workout_type')} → col {dbg.get('col')} | "
                                       f"書き込み: {list(dbg.get('written', {}).keys())}")
                    st.success(f"✅ Workout シートに書き込みました！（{date_saved}）")

                # クラウドモード: 書き込み済みExcelをGoogle Driveに自動アップロード
                if _IS_CLOUD:
                    from gdrive_helper import upload_from_path
                    upload_from_path(excel_path_saved)
                    pathlib.Path(excel_path_saved).unlink(missing_ok=True)
                    # キャッシュをクリア（次回は最新版をダウンロードする）
                    st.session_state.pop("cloud_excel_tmp_path", None)
                    st.success("☁️ Google Driveに自動保存しました")

                del st.session_state["all_results"]

            except ExcelFileLockError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"書き込みエラー: {e}")
                with st.expander("詳細"):
                    import traceback
                    st.code(traceback.format_exc())

    # クラウドモード: ダウンロードボタン表示
    if _IS_CLOUD and st.session_state.get("download_bytes"):
        st.download_button(
            label="⬇️ 更新済みExcelをダウンロード",
            data=st.session_state["download_bytes"],
            file_name=st.session_state.get("download_name", "Yuki データ.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
