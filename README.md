# 🎬 video2doc

**English** | [Русский](README.ru.md)

**Stop pausing tutorials to take notes.** video2doc turns YouTube videos and whole playlists into
ready-to-print DOCX/PDF manuals with smart screenshots and timestamped subtitles.

Perfect for lessons, tutorials, recipes, DIY guides, lectures, webinars and training videos —
no AI, no API keys, no cost. Everything runs on your own computer.

## ✨ Features

- 🔗 Works with a video file, one or several YouTube links, or an entire playlist
- 🖼️ Smart screenshots — captured when the picture changes, never blurry, no duplicates
- 📝 Text under every screenshot — YouTube subtitles in the original language, or local speech recognition (Whisper) when there are none
- 📄 Output as Word (DOCX) and PDF, plus `.txt` and `.srt` subtitles
- 📁 Neat folders: `Documents\video2doc\Project\Video\`
- 🔁 Playlists resume where they stopped — finished videos are skipped

## 🖼️ How screenshots are chosen

1. Detects moments when the picture on screen changes.
2. Waits ~1 s after the change and takes the frame only once the image is stable (no blurry transitions).
3. Drops near-identical frames.
4. If the screen doesn't change for a long time, still takes a screenshot at least every 40 s.
5. You can add your own moments: `--at 01:20,05:43`.

## ⚙️ Installation (Windows)

1. **Python 3.12** — from python.org, the **"Windows installer (64-bit)" (.exe)**, not .msix.
   Tick **"Add python.exe to PATH"**. Newer Python (3.14, 3.15) is not supported yet —
   if you have it, the program automatically restarts itself with 3.12.
2. **Libraries:**
   ```
   py -3.12 -m pip install -r requirements.txt
   ```
3. **PDF export:** Microsoft Word (via `docx2pdf`, already in requirements) or LibreOffice
   (`winget install TheDocumentFoundation.LibreOffice`).
4. Put `video2doc.py` and `video2doc.bat` in the same folder.

Optional: `winget install DenoLand.Deno` — removes YouTube warnings and reduces download errors.

## ▶️ Usage

Double-click **`video2doc.bat`**. The program asks:
1. **Project** — a number from the list, a new name, or Enter.
2. **YouTube link** (several links separated by spaces) or a path to a video file.
3. For a playlist link: Enter — only this video, `all` — the whole playlist, or numbers like `12,33,49`.

The result folder opens automatically.

### Command line

```
py -3.12 video2doc.py lesson.mp4
py -3.12 video2doc.py "https://youtu.be/AAA" "https://youtu.be/BBB" --project MyCourse
py -3.12 video2doc.py "https://www.youtube.com/playlist?list=XXXX" --playlist --items 1-10
```

| Option | What it does |
|---|---|
| `--project` | project folder inside `Documents\video2doc` |
| `--playlist`, `--items 1-10,15` | whole playlist or selected numbers |
| `--redo` | redo videos that are already finished |
| `--at 01:20,05:43` | extra screenshots at given moments |
| `--every 30` | a screenshot every 30 s instead of change detection |
| `--threshold 0.15` | change sensitivity (lower — more screenshots) |
| `--dup 4` | "same frame" threshold (higher — fewer screenshots) |
| `--min-gap 4` / `--max-gap 40` | min / max seconds between screenshots |
| `--whisper`, `--model small` | use speech recognition even if subtitles exist; Whisper model |

## 📂 Output

```
Documents\video2doc\
   My Course\                          ← project
      Lesson title\                    ← one video = one folder
         Lesson title.docx
         Lesson title.pdf
         frames\                       ← screenshots
         transcript.txt                ← text with timestamps
         subtitles.srt
         info.txt                      ← title, link, channel, date
         video.mp4
   Playlist name\
      playlist.txt                     ← all videos with links
      playlist_log.txt                 ← what's done, what failed
      001 - First video\
      002 - ...
```

## 🛠️ Troubleshooting

| Message | Fix |
|---|---|
| "Sideloading of unpublished apps is disabled" | you downloaded the .msix — use the Python .exe installer |
| "… is not recognized as the name of a cmdlet" | close and reopen PowerShell |
| `Failed to build 'av'` | Python 3.15 is running — use `video2doc.bat` or `py -3.12` |
| `No module named …` | repeat installation step 2 |
| `cublas64_12.dll is not found` | handled automatically — Whisper switches to CPU |
| `HTTP Error 403 / 429` | retried automatically after 60 s; otherwise `py -3.12 -m pip install -U yt-dlp` and try later |
| PDF not created | install Microsoft Word + `docx2pdf`, or LibreOffice |

## 📜 License

MIT — free to use, modify and share.
