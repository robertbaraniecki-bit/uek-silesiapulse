from flask import Flask
import subprocess

app = Flask(__name__)

BASE = r'C:\Users\Oloros\OneDrive\Pulpit\SilesiaPulse_SEMINARIUM_28.05'

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

if __name__ == '__main__':
    app.run(port=5000)