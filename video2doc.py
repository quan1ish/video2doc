#!/usr/bin/env python3
"""
video2doc — turn videos (file, YouTube link(s) or a whole playlist) into
DOCX/PDF manuals with smart screenshots and timestamped subtitles.

Examples:
  py -3.12 video2doc.py                                    # no arguments — asks everything
  py -3.12 video2doc.py lesson.mp4
  py -3.12 video2doc.py "https://youtu.be/AAA" "https://youtu.be/BBB" --project MyCourse
  py -3.12 video2doc.py "https://www.youtube.com/playlist?list=XXXX" --playlist --items 1-10
  py -3.12 video2doc.py lesson.mp4 --at 01:20,05:43        # extra screenshots
  py -3.12 video2doc.py lesson.mp4 --every 30              # a screenshot every 30 s

Install (once, Python 3.12):
  py -3.12 -m pip install -r requirements.txt
  PDF: Microsoft Word (via docx2pdf) or LibreOffice.

Interface language: English or Russian, chosen from the Windows language.
Force it with --ui en / --ui ru or the VIDEO2DOC_LANG environment variable.
"""
import argparse, datetime, os, re, subprocess, shutil, sys, time
from pathlib import Path

# Python 3.14+ is not supported by the libraries yet — restart with 3.12 if available
if os.name == "nt" and sys.version_info[:2] >= (3, 14) and not os.environ.get("V2D_RELAUNCHED"):
    try:
        if subprocess.run(["py", "-3.12", "-c", ""], capture_output=True).returncode == 0:
            os.environ["V2D_RELAUNCHED"] = "1"
            sys.exit(subprocess.call(["py", "-3.12", os.path.abspath(__file__), *sys.argv[1:]]))
    except FileNotFoundError:
        pass


def _detect_ui():
    env = os.environ.get("VIDEO2DOC_LANG", "")
    if env:
        return "ru" if env[:2].lower() == "ru" else "en"
    cis = ("ru", "kk", "uk", "be", "uz", "ky")
    try:
        if os.name == "nt":
            import ctypes
            lid = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF
            return "ru" if lid in (0x19, 0x3F, 0x22, 0x23, 0x43, 0x40) else "en"
        import locale
        loc = locale.getlocale()[0] or os.environ.get("LANG", "")
        return "ru" if loc[:2].lower() in cis else "en"
    except Exception:
        return "en"

UI = _detect_ui()

def T(ru, en):
    """Interface text in the current language."""
    return ru if UI == "ru" else en


def find_ffmpeg():
    """ffmpeg from PATH, otherwise the one bundled with imageio-ffmpeg."""
    p = shutil.which("ffmpeg")
    if p:
        return p
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return None

FFMPEG = find_ffmpeg()
DOC_SUFFIX = ""  # the AI version sets " (ИИ)" so both documents can sit side by side

def is_done(folder):
    docs = [p for p in folder.glob("*.docx") if not p.name.startswith("~$")]
    if DOC_SUFFIX:
        return any(p.stem.endswith(DOC_SUFFIX) for p in docs)
    return any(not p.stem.endswith(" (ИИ)") for p in docs)


# ---------- utilities ----------
def fmt(t):
    t = int(t); return f"{t//3600:02d}:{t%3600//60:02d}:{t%60:02d}"

def parse_time(s):
    sec = 0.0
    for p in s.strip().split(":"):
        sec = sec * 60 + float(p)
    return sec

def duration(video):
    err = subprocess.run([FFMPEG, "-i", str(video)], capture_output=True, text=True,
                         encoding="utf-8", errors="ignore").stderr
    m = re.search(r"Duration:\s*(\d+):(\d+):([\d.]+)", err)
    return int(m[1]) * 3600 + int(m[2]) * 60 + float(m[3]) if m else 0.0

def documents_dir():
    """The Documents folder (handles OneDrive redirection)."""
    if os.name == "nt":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # CSIDL_PERSONAL
            if buf.value:
                return Path(buf.value)
        except Exception:
            pass
    return Path.home() / "Documents"

def safe_name(s, n=80):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s).strip(" .")
    return s[:n].strip(" .") or "video"

def _has(mod):
    import importlib.util
    return importlib.util.find_spec(mod) is not None

def now():
    return f"{datetime.datetime.now():%Y-%m-%d %H:%M}"

def save_text(tr, txt_path, srt_path):
    txt_path.write_text("\n".join(f"[{fmt(s)}] {t}" for s, e, t in tr), encoding="utf-8")
    ts = lambda x: f"{fmt(x)},{int((x % 1) * 1000):03d}"
    with open(srt_path, "w", encoding="utf-8") as f:
        for i, (s, e, t) in enumerate(tr, 1):
            f.write(f"{i}\n{ts(s)} --> {ts(e)}\n{t}\n\n")


# ---------- 1. YouTube ----------
def parse_vtt(path):
    """YouTube subtitles (.vtt) → [(start, end, text)] without repeats."""
    out, last = [], ""
    for b in Path(path).read_text(encoding="utf-8").split("\n\n"):
        lines = b.strip().splitlines()
        for i, l in enumerate(lines):
            m = re.match(r"(\d+):(\d+):([\d.]+) --> (\d+):(\d+):([\d.]+)", l)
            if m:
                g = [float(x) for x in m.groups()]
                s, e = g[0]*3600 + g[1]*60 + g[2], g[3]*3600 + g[4]*60 + g[5]
                for t in lines[i+1:]:
                    t = re.sub(r"<[^>]+>", "", t).strip()
                    if t and t != last:
                        out.append((s, e, t)); last = t
                break
    return out

def glob_escape(s):
    return re.sub(r"([\[\]*?])", r"[\1]", s)

def from_youtube(url, project_dir, lang, folder_name=None):
    """Download the video (up to 720p) and subtitles in the original language.
    Returns (video, subtitles or None, folder, title)."""
    import yt_dlp
    print(T("[1/4] Скачиваю с YouTube...", "[1/4] Downloading from YouTube..."))
    with yt_dlp.YoutubeDL({"noplaylist": True, "quiet": True}) as y:
        meta = y.extract_info(url, download=False)
    folder = project_dir / (folder_name or safe_name(meta.get("title") or meta.get("id", "video")))
    folder.mkdir(parents=True, exist_ok=True)
    print(T(f"  папка: {folder}", f"  folder: {folder}"))
    base = {"outtmpl": str(folder / "video.%(ext)s"), "noplaylist": True,
            "ffmpeg_location": FFMPEG, "retries": 10, "fragment_retries": 10}

    with yt_dlp.YoutubeDL({**base, "format": "bv*[height<=720]+ba/b[height<=720]/b",
                           "merge_output_format": "mp4"}) as y:
        info = y.extract_info(url)
        video = Path(y.prepare_filename(info)).with_suffix(".mp4")
    (folder / "info.txt").write_text(
        f"{T('Название', 'Title')}: {info.get('title')}\n"
        f"{T('Ссылка', 'Link')}: {info.get('webpage_url', url)}\n"
        f"{T('Канал', 'Channel')}: {info.get('uploader')}\n"
        f"{T('Длительность', 'Duration')}: {fmt(info.get('duration') or 0)}\n"
        f"{T('Обработано', 'Processed')}: {now()}\n", encoding="utf-8")

    manual = info.get("subtitles") or {}
    auto = info.get("automatic_captions") or {}
    orig = lang or info.get("language")
    if not orig:
        o = [k for k in auto if k.endswith("-orig")]
        orig = o[0][:-5] if o else None
    cands = []
    if orig:
        cands += [k for k in manual if k == orig or k.startswith(orig + "-")]
        cands += [k for k in (orig + "-orig", orig) if k in auto]
    elif manual:
        cands.append(next(iter(manual)))

    for code in dict.fromkeys(cands):
        try:
            with yt_dlp.YoutubeDL({**base, "skip_download": True, "writesubtitles": True,
                                   "writeautomaticsub": True, "subtitleslangs": [re.escape(code)],
                                   "subtitlesformat": "vtt", "quiet": True,
                                   "no_warnings": True}) as y:
                y.download([url])
        except Exception as e:
            print(T(f"  субтитры '{code}' не получены", f"  subtitles '{code}' not received") + f" ({str(e)[-60:]})")
            continue
        subs = sorted(folder.glob(glob_escape(video.stem) + f"*.{code}.vtt"))
        if subs:
            print(T(f"  субтитры YouTube: {code}", f"  YouTube subtitles: {code}"))
            return video, parse_vtt(subs[0]), folder, info.get("title") or folder.name
    return video, None, folder, info.get("title") or folder.name


# ---------- 2. text: subtitles or speech recognition ----------
def load_audio(video, sr=16000):
    """Audio via ffmpeg → numpy (16 kHz, mono)."""
    import numpy as np
    raw = subprocess.run([FFMPEG, "-nostdin", "-v", "error", "-i", str(video), "-f", "s16le",
                          "-ac", "1", "-ar", str(sr), "-"], capture_output=True, check=True).stdout
    return np.frombuffer(raw, np.int16).astype(np.float32) / 32768.0

def transcribe(video, model_name, lang):
    from faster_whisper import WhisperModel
    print(T(f"[2/4] Субтитров нет — распознаю речь (Whisper {model_name})...",
            f"[2/4] No subtitles — recognizing speech (Whisper {model_name})..."))
    audio = load_audio(video)
    for device in ("auto", "cpu"):  # GPU without CUDA libraries → fall back to CPU
        try:
            model = WhisperModel(model_name, device=device, compute_type="int8")
            segs, info = model.transcribe(audio, language=lang, vad_filter=True)
            out = []
            for s in segs:
                out.append((s.start, s.end, s.text.strip()))
                print(f"  {fmt(s.start)}  {s.text.strip()[:80]}")
            return out
        except Exception as e:
            if device == "cpu" or not re.search(r"cublas|cudnn|cuda", str(e), re.I):
                raise
            print(T("  Видеокарта недоступна для Whisper (нет библиотек CUDA) — считаю на процессоре...",
                    "  GPU not available for Whisper (no CUDA libraries) — using CPU..."))


# ---------- 3. screenshots ----------
TW, TH = 64, 36  # thumbnail size for comparing frames

def thumb(video, t):
    """Small grayscale frame (64×36) for comparison."""
    import numpy as np
    raw = subprocess.run([FFMPEG, "-v", "error", "-ss", f"{t:.2f}", "-i", str(video),
                          "-frames:v", "1", "-vf", f"scale={TW}:{TH},format=gray",
                          "-f", "rawvideo", "-"], capture_output=True).stdout
    return np.frombuffer(raw, np.uint8).astype(np.float32) if len(raw) == TW * TH else None

def diff(a, b):
    if a is None or b is None:
        return 255.0
    return float(abs(a - b).mean())

def scene_times(video, threshold):
    r = subprocess.run([FFMPEG, "-i", str(video), "-vf", f"select='gt(scene,{threshold})',showinfo",
                        "-f", "null", "-"], capture_output=True, text=True,
                       encoding="utf-8", errors="ignore")
    return [float(x) for x in re.findall(r"pts_time:([\d.]+)", r.stderr)]

def stable_time(video, t, dur):
    """Move forward until the frame stops changing (avoid transitions and blur)."""
    for _ in range(4):
        t2 = min(t + 0.7, dur - 0.1)
        if t2 <= t or diff(thumb(video, t), thumb(video, t2)) < 3:
            return t
        t = t2
    return t

def pick_times(video, every, threshold, min_gap, max_gap, dup, extra):
    """Screenshot moments: scene change → stable frame after it → no duplicates → no long gaps."""
    dur = duration(video)
    if every:
        cands = [float(t) for t in range(1, int(dur), every)]
    else:
        cands = [1.0] + [t + 1.0 for t in scene_times(video, threshold)]
        filled, prev = [], 0.0          # fill long stretches without changes
        for t in sorted(cands) + [dur]:
            while t - prev > max_gap:
                prev += max_gap
                filled.append(prev)
            prev = t
        cands = sorted(set(cands + filled))

    kept, last_t, last_img = [], -1e9, None
    for t in cands:
        if t >= dur - 0.2 or t - last_t < min_gap:
            continue
        t = stable_time(video, t, dur)
        img = thumb(video, t)
        if img is None or diff(img, last_img) < dup:   # nearly the same frame — skip
            continue
        kept.append(t); last_t, last_img = t, img
    return sorted(set(round(t, 1) for t in kept) | set(e for e in extra if e < dur))

def grab(video, t, path):
    subprocess.run([FFMPEG, "-y", "-v", "error", "-ss", f"{t:.2f}", "-i", str(video),
                    "-frames:v", "1", "-q:v", "2", "-vf", "scale='min(1280,iw)':-2", str(path)],
                   check=True)

def screenshots(video, times, folder):
    print(T(f"[3/4] Сохраняю {len(times)} скриншотов...", f"[3/4] Saving {len(times)} screenshots..."))
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir()
    shots = []
    for i, t in enumerate(times, 1):
        p = folder / f"{i:03d}_{fmt(t).replace(':', '-')}.jpg"
        grab(video, t, p)
        shots.append((t, p))
    return shots


# ---------- 4. document ----------
def build_docx(title, video, tr, shots, out_docx, link=""):
    from docx import Document
    from docx.shared import Cm, Pt
    print(T("[4/4] Собираю документ...", "[4/4] Building the document..."))
    d = Document()
    d.styles["Normal"].font.name = "Times New Roman"
    d.styles["Normal"].font.size = Pt(12)
    d.add_heading(title, 0)
    d.add_paragraph(f"{T('Длительность', 'Duration')}: {fmt(duration(video))}"
                    + (f"\n{T('Источник', 'Source')}: {link}" if link else ""))
    bounds = [t for t, _ in shots] + [1e9]
    for i, (t, img) in enumerate(shots):
        d.add_heading(f"[{fmt(t)}]", 2)
        d.add_picture(str(img), width=Cm(16))
        txt = " ".join(s for a, b, s in tr
                       if bounds[i] <= a + 0.01 < bounds[i + 1] or (i == 0 and a < bounds[0]))
        if txt:
            d.add_paragraph(txt)
    d.save(out_docx)

def find_soffice():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice and os.name == "nt":
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
            if os.path.exists(p):
                soffice = p
    return soffice

def to_pdf(docx_path):
    soffice = find_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_path.parent), str(docx_path)], capture_output=True)
        return
    try:
        from docx2pdf import convert  # Windows + MS Word
        convert(str(docx_path))
    except Exception:
        print(T("  PDF не создан: нужен MS Word (+ docx2pdf) или LibreOffice. DOCX готов.",
                "  PDF not created: needs MS Word (+ docx2pdf) or LibreOffice. DOCX is ready."))


# ---------- one video ----------
def process_one(a, src, project_dir, folder_name=None):
    link = ""
    tr = None
    if src.startswith("http"):
        video, tr, out, title = from_youtube(src, project_dir, a.lang, folder_name)
        link = src
    else:
        video = Path(src).resolve()
        title = video.stem
        out = project_dir / (folder_name or safe_name(title))
        out.mkdir(parents=True, exist_ok=True)
        (out / "info.txt").write_text(f"{T('Файл', 'File')}: {video}\n{T('Обработано', 'Processed')}: {now()}\n",
                                      encoding="utf-8")

    if tr is None or a.whisper:
        tr = transcribe(video, a.model, a.lang)
    else:
        print(T(f"[2/4] Субтитры YouTube: {len(tr)} фраз", f"[2/4] YouTube subtitles: {len(tr)} lines"))
    save_text(tr, out / "transcript.txt", out / "subtitles.srt")

    extra = [parse_time(x) for x in a.at.split(",") if x.strip()]
    times = pick_times(video, a.every, a.threshold, a.min_gap, a.max_gap, a.dup, extra)
    shots = screenshots(video, times, out / "frames")

    docx_path = out / f"{safe_name(title)}{DOC_SUFFIX}.docx"
    build_docx(title, video, tr, shots, docx_path, link)
    to_pdf(docx_path)
    print(T(f"Готово: {out}", f"Done: {out}"))
    return out


# ---------- playlist ----------
def playlist_entries(url):
    import yt_dlp
    m = re.search(r"[?&]list=([\w-]+)", url)
    purl = f"https://www.youtube.com/playlist?list={m[1]}" if m else url
    with yt_dlp.YoutubeDL({"extract_flat": "in_playlist", "quiet": True}) as y:
        info = y.extract_info(purl, download=False)
    items = [(i, e.get("title") or e["id"], f"https://www.youtube.com/watch?v={e['id']}")
             for i, e in enumerate(info.get("entries") or [], 1) if e and e.get("id")]
    return info.get("title") or T("Плейлист", "Playlist"), items

def parse_range(text, n):
    """'1-10,15,20-' → numbers; empty → all."""
    text = (text or "").replace(" ", "")
    if not text:
        return set(range(1, n + 1))
    nums = set()
    for part in text.split(","):
        if "-" in part:
            x, y = part.split("-", 1)
            nums |= set(range(int(x or 1), int(y or n) + 1))
        elif part:
            nums.add(int(part))
    return {x for x in nums if 1 <= x <= n}

def run_with_retry(a, url, project_dir, name):
    for attempt in (1, 2):
        try:
            process_one(a, url, project_dir, name)
            return "OK"
        except Exception as e:
            status = T("ОШИБКА: ", "ERROR: ") + re.sub(r"\x1b\[[0-9;]*m", "", str(e))
            print(f"  {status}")
            if attempt == 1 and re.search(r"403|429|Forbidden|Too Many", str(e)):
                print(T("  YouTube ограничил доступ — жду 60 с и пробую ещё раз...",
                        "  YouTube limited access — waiting 60 s and retrying..."))
                time.sleep(60)
                continue
            return status

def process_list(a, items, project_dir, log_name):
    """items: [(number, title, link)] → each video into its own folder "NNN - Title"."""
    log = project_dir / log_name
    ok = fail = skip = 0
    for k, (i, title, url) in enumerate(items, 1):
        name = f"{i:03d} - {safe_name(title, 70)}" if i else None
        print(f"\n===== [{k}/{len(items)}] {title} =====")
        if name and not a.redo and is_done(project_dir / name):
            print(T("  уже готово — пропускаю (--redo, чтобы переделать)",
                    "  already done — skipping (use --redo to redo)"))
            skip += 1
            continue
        status = run_with_retry(a, url, project_dir, name)
        ok += status == "OK"; fail += status != "OK"
        with open(log, "a", encoding="utf-8") as f:
            f.write(f"{now()} | {i or '-':>3} | {title} | {status}\n")
        if k < len(items):
            time.sleep(5)  # pause so YouTube doesn't rate-limit us
    print(T(f"\nГотово: успешно {ok}, пропущено {skip}, с ошибкой {fail}. Журнал: {log}",
            f"\nFinished: {ok} done, {skip} skipped, {fail} failed. Log: {log}"))
    return project_dir


# ---------- main ----------
def main():
    ap = argparse.ArgumentParser(description="Video → screenshots + subtitles (DOCX/PDF)")
    ap.add_argument("video", nargs="+", help="video file, or one or more YouTube links")
    ap.add_argument("--project", default="", help="project folder inside Documents\\video2doc")
    ap.add_argument("--root", default=None, help="main folder (default: Documents\\video2doc)")
    ap.add_argument("--playlist", action="store_true", help="process the playlist from the link")
    ap.add_argument("--items", default="", help="playlist numbers: 1-10,15 (empty — all)")
    ap.add_argument("--redo", action="store_true", help="redo videos that are already finished")
    ap.add_argument("--lang", default=None, help="speech language (usually detected automatically)")
    ap.add_argument("--whisper", action="store_true", help="recognize speech even if subtitles exist")
    ap.add_argument("--model", default="small", help="Whisper model: tiny/base/small/medium")
    ap.add_argument("--every", type=int, default=0, help="a screenshot every N s instead of change detection")
    ap.add_argument("--threshold", type=float, default=0.15, help="change sensitivity 0..1")
    ap.add_argument("--min-gap", type=float, default=4, help="min seconds between screenshots")
    ap.add_argument("--max-gap", type=float, default=40, help="max seconds without a screenshot")
    ap.add_argument("--dup", type=float, default=4, help="'same frame' threshold (higher — fewer screenshots)")
    ap.add_argument("--at", default="", help="extra screenshots: 01:20,05:43")
    ap.add_argument("--ui", choices=["en", "ru"], default=None, help="interface language")

    if len(sys.argv) == 1:  # no arguments — interactive mode
        root = documents_dir() / "video2doc"
        projects = sorted(p.name for p in root.glob("*") if p.is_dir()) if root.exists() else []
        print(T(f"Папка для результатов: {root}", f"Output folder: {root}"))
        if projects:
            print(T("Проекты:", "Projects:"), ", ".join(f"{i}) {p}" for i, p in enumerate(projects, 1)))
        pr = input(T("Проект (номер, новое название или Enter — «Без проекта»): ",
                     "Project (number, new name or Enter — \"No project\"): ")).strip()
        if pr.isdigit() and 1 <= int(pr) <= len(projects):
            pr = projects[int(pr) - 1]
        src = input(T("Ссылка YouTube (можно несколько через пробел) или путь к видео: ",
                      "YouTube link (several separated by spaces) or path to a video: ")).strip()
        srcs = src.split() if src.count("http") > 1 else [src.strip('"')]
        args = [*srcs, "--project", pr]
        if len(srcs) == 1 and "list=" in src:
            it = input(T("Это плейлист. Какие видео? Enter — только это видео, "
                         "«все» — весь плейлист, или номера, напр. 12,33,49: ",
                         "This is a playlist. Which videos? Enter — only this one, "
                         "'all' — the whole playlist, or numbers, e.g. 12,33,49: ")).strip().lower()
            if it:
                args.append("--playlist")
                if it not in ("все", "all", "y", "a"):
                    args += ["--items", it]
        a = ap.parse_args(args)
    else:
        a = ap.parse_args()

    global UI
    if a.ui:
        UI = a.ui
    need = ["faster_whisper", "docx", "numpy"]
    if any(s.startswith("http") for s in a.video): need.append("yt_dlp")
    if os.name == "nt" and not find_soffice(): need.append("docx2pdf")
    names = {"faster_whisper": "faster-whisper", "docx": "python-docx", "yt_dlp": "yt-dlp",
             "docx2pdf": "docx2pdf", "numpy": "numpy"}
    missing = [names[m] for m in need if not _has(m)]
    if missing:
        print(T("Не установлены библиотеки: ", "Missing libraries: ") + ", ".join(missing)
              + f"\nPython: {sys.executable}")
        if input(T("Установить сейчас? (y/n): ", "Install now? (y/n): ")).strip().lower().startswith(("y", "д")):
            subprocess.run([sys.executable, "-m", "pip", "install", *missing, "imageio-ffmpeg"])
            import importlib; importlib.invalidate_caches()
            if any(not _has(m) for m in need):
                raise SystemExit(T("Установка не удалась. Используйте Python 3.12: ",
                                   "Installation failed. Use Python 3.12: ")
                                 + "py -3.12 -m pip install " + " ".join(missing))
        else:
            raise SystemExit(T("Выполните:", "Run:") + f'\n  "{sys.executable}" -m pip install {" ".join(missing)}')
    if not FFMPEG:
        raise SystemExit(T("Не найден ffmpeg. Выполните:", "ffmpeg not found. Run:")
                         + "\n  py -3.12 -m pip install imageio-ffmpeg")

    root = Path(a.root) if a.root else documents_dir() / "video2doc"
    if a.playlist and a.video[0].startswith("http"):
        print(T("Получаю список видео плейлиста...", "Getting the playlist..."))
        ptitle, items = playlist_entries(a.video[0])
        project_dir = root / safe_name(a.project or ptitle)
        project_dir.mkdir(parents=True, exist_ok=True)
        (project_dir / "playlist.txt").write_text(
            "\n".join(f"{i:03d} | {t} | {u}" for i, t, u in items), encoding="utf-8")
        chosen = parse_range(a.items, len(items))
        todo = [x for x in items if x[0] in chosen]
        print(T(f"Плейлист «{ptitle}»: всего {len(items)}, к обработке {len(todo)}",
                f"Playlist \"{ptitle}\": {len(items)} videos, {len(todo)} to process"))
        out = process_list(a, todo, project_dir, "playlist_log.txt")
    else:
        project_dir = root / safe_name(a.project or T("Без проекта", "No project"))
        project_dir.mkdir(parents=True, exist_ok=True)
        if len(a.video) == 1:
            out = process_one(a, a.video[0], project_dir)
        else:
            out = process_list(a, [(0, s, s) for s in a.video], project_dir, "log.txt")
    if os.name == "nt" and len(sys.argv) == 1:
        os.startfile(out)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(f"\n{e.code}")
    except BaseException:
        import traceback
        print(T("\nОшибка:", "\nError:"))
        traceback.print_exc()
    finally:
        if len(sys.argv) == 1:
            input(T("\nНажмите Enter, чтобы закрыть...", "\nPress Enter to close..."))
