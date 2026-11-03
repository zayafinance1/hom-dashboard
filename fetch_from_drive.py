"""ينزّل ملف الإكسل من Google Drive باستخدام Service Account (بدون ما الملف يتشارك للعامة)."""
import os, io, json, sys
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

creds = service_account.Credentials.from_service_account_info(
    json.loads(os.environ["GDRIVE_KEY"]),
    scopes=["https://www.googleapis.com/auth/drive.readonly"])
drive = build("drive", "v3", credentials=creds, cache_discovery=False)
fid = os.environ["GDRIVE_FILE_ID"]
import datetime as dt
meta = drive.files().get(fileId=fid, fields="modifiedTime,mimeType", supportsAllDrives=True).execute()
mt = meta["modifiedTime"]
cairo = dt.datetime.fromisoformat(mt.replace("Z", "+00:00")) + dt.timedelta(hours=3)
open("modified.txt", "w").write(cairo.strftime("%Y-%m-%d %H:%M"))
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
if meta["mimeType"] == "application/vnd.google-apps.spreadsheet":
    req = drive.files().export_media(fileId=fid, mimeType=XLSX)   # Google Sheet -> xlsx
else:
    req = drive.files().get_media(fileId=fid, supportsAllDrives=True)
buf = io.FileIO(sys.argv[1], "wb")
dl = MediaIoBaseDownload(buf, req, chunksize=8 * 1024 * 1024)
done = False
while not done:
    _, done = dl.next_chunk()
print("downloaded ->", sys.argv[1])
