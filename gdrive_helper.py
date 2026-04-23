"""
Google Drive helper — download/upload the Excel file.
Requires env vars:
  GOOGLE_SERVICE_ACCOUNT_JSON  (full JSON string of the service account key)
  GDRIVE_FILE_ID               (the file ID from the Drive URL)
"""

import io
import json
import os
import tempfile


def _get_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません")

    sa_info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def download_to_temp() -> str:
    """GoogleドライブのファイルをExcel形式で一時ファイルにダウンロードしてパスを返す"""
    from googleapiclient.http import MediaIoBaseDownload

    file_id = os.environ.get("GDRIVE_FILE_ID")
    if not file_id:
        raise ValueError("GDRIVE_FILE_ID が設定されていません")

    service = _get_service()

    # ファイルのMIMEタイプを確認
    meta = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime = meta.get("mimeType", "")

    if mime == "application/vnd.google-apps.spreadsheet":
        # Google Sheets形式 → xlsxとしてエクスポート
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        # 通常のExcelファイル
        request = service.files().get_media(fileId=file_id)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    buf.seek(0)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(buf.read())
    tmp.close()
    return tmp.name


def upload_from_path(local_path: str) -> None:
    """ローカルのExcelをGoogleドライブに上書きアップロード"""
    from googleapiclient.http import MediaFileUpload

    file_id = os.environ.get("GDRIVE_FILE_ID")
    if not file_id:
        raise ValueError("GDRIVE_FILE_ID が設定されていません")

    service = _get_service()
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )
    service.files().update(fileId=file_id, media_body=media).execute()


def is_configured() -> bool:
    """Drive連携に必要な環境変数が揃っているか確認"""
    return bool(
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        and os.environ.get("GDRIVE_FILE_ID")
    )
