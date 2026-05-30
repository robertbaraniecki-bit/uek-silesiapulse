from flask import Flask
import subprocess
import webbrowser

app = Flask(__name__)

BASE = r'C:\Users\Oloros\OneDrive\Pulpit\SilesiaPulse_SEMINARIUM_28.05'

# ── SilesiaPulse (oryginalne endpointy) ──────────────────────────────────────

@app.route('/open/excel')
def open_excel():
    subprocess.Popen(
        f'start "" "{BASE}\\excel\\UEK_Baraniecki_2026.xlsm"',
        shell=True
    )
    return 'ok'

@app.route('/open/vscode')
def open_vscode():
    subprocess.Popen([
        r'C:\Users\Oloros\AppData\Local\Programs\Microsoft VS Code\bin\code.cmd',
        BASE
    ], shell=True)
    return 'ok'

@app.route('/open/gitkraken')
def open_gitkraken():
    subprocess.Popen([
        r'C:\Users\Oloros\AppData\Local\gitkraken\bin\gitkraken.cmd',
        '--path', r'C:\Projekty\uek-silesiapulse'
    ], shell=True)
    return 'ok'

@app.route('/open/cmd')
def open_cmd():
    subprocess.Popen(
        f'start /max cmd /k "cd /d {BASE}"',
        shell=True
    )
    return 'ok'

# ── Tool Stack (nowe endpointy) ───────────────────────────────────────────────

@app.route('/open/claude')
def open_claude():
    # Claude Desktop — Microsoft Store (PackageFamilyName: Claude_pzs8sxrjxfjjc)
    subprocess.Popen(
        'explorer shell:AppsFolder\\Claude_pzs8sxrjxfjjc!Claude',
        shell=True
    )
    return 'ok'

@app.route('/open/claude-design')
def open_claude_design():
    # Claude Design — otwiera w Chrome (nie domyślnej przeglądarce)
    subprocess.Popen(
        '"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" https://claude.ai',
        shell=True
    )
    return 'ok'

@app.route('/open/higgsfield')
def open_higgsfield():
    webbrowser.open('https://higgsfield.ai')
    return 'ok'

@app.route('/open/gdrive')
def open_gdrive():
    webbrowser.open('https://drive.google.com')
    return 'ok'

@app.route('/open/notion')
def open_notion():
    # Notion Desktop — klasyczny instalator
    subprocess.Popen(
        r'C:\Users\Oloros\AppData\Local\Programs\Notion\Notion.exe',
        shell=True
    )
    return 'ok'

@app.route('/open/obs')
def open_obs():
    # OBS Studio — domyślna ścieżka instalacji 64-bit
    subprocess.Popen(
        r'C:\Program Files\obs-studio\bin\64bit\obs64.exe',
        shell=True
    )
    return 'ok'

@app.route('/open/davinci')
def open_davinci():
    # DaVinci Resolve — domyślna ścieżka instalacji
    subprocess.Popen(
        r'C:\Program Files\Blackmagic Design\DaVinci Resolve\Resolve.exe',
        shell=True
    )
    return 'ok'

# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(port=5000)
