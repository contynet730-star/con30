# 🌙 Fortune Divination App

1981/07/30・男性（下弦の月・銀／命数6）の鑑定を **毎朝6:00** に表示するミニアプリ。

## 使い方

### 1回だけ表示する
```bash
python3 fortune.py
```
`logs/fortune_YYYYMMDD.txt` にもその日の鑑定が保存されます。

### 毎朝6:00に自動表示する（cron）
```bash
bash schedule_6am.sh          # 登録
crontab -l                    # 確認
bash schedule_6am.sh remove   # 解除
```

### macOS で通知センターに出したい場合
`fortune.py` 末尾に以下を追記すると、標準出力に加えて通知を鳴らせます。
```python
import subprocess, shutil
if shutil.which("osascript"):
    subprocess.run(["osascript", "-e",
        'display notification "今日の鑑定が届きました" with title "🌙 Fortune"'])
```

### systemd timer を使う場合（Linux サーバ）
`/etc/systemd/system/fortune.timer` を作成:
```
[Unit]
Description=Daily fortune at 06:00

[Timer]
OnCalendar=*-*-* 06:00:00
Persistent=true

[Install]
WantedBy=timers.target
```
`/etc/systemd/system/fortune.service`:
```
[Service]
Type=oneshot
WorkingDirectory=/home/user/con30/fortune_app
ExecStart=/usr/bin/python3 fortune.py
```
有効化: `sudo systemctl enable --now fortune.timer`

## 構成
```
fortune_app/
├── fortune.py        # 鑑定本文を出力 & logs/ に保存
├── schedule_6am.sh   # cron 登録/解除スクリプト
├── README.md
└── logs/             # 日々の出力（自動生成）
```
