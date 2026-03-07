# Did You Close the Speaker? 🔊

**Safe power management for active studio monitors** powered by **TP-Link Tapo P115** smart plug.

Automatically turns off speakers **before** PC shutdown/sleep/restart — preventing damaging power-off pops on active monitors like iLoud Micro/MTMs.


## Features

| Feature | Status |
|---------|--------|
| System tray app with quick actions | ✅ |
| Safe Shutdown/Sleep/Restart | ✅ |
| Separate compact Settings window | ✅ |
| Windows Startup autorun | ✅ |
| Idle detection (auto speaker off) | ✅ |
| Auto resume speaker when active | ✅ |
| Sleep/Shutdown watchdog | ✅ |
| Single instance support | ✅ |
| PyInstaller exe build ready | ✅ |
| CLI fallback (main.py) | ✅ |

## Quick Start

### 1. Install
```bash
git clone <your-repo>
cd did-you-close-the-speaker
pip install -r requirements.txt
```

### 2. Configure
```bash
cp config.example.json config.json
# Edit with your Tapo P115 IP + account
```

### 3. Run
```bash
# GUI (system tray)
python gui.py

# Startup (no console window)
pythonw gui.py --startup

# Startup + auto speaker ON
pythonw gui.py --startup --speaker-on
```

## GUI Usage

**Main Window** (400x270 - super compact):
```
🔊 Did You Close the Speaker?
● Status: ON    ↻ Refresh
[Turn Speaker OFF]
───────────────
Shutdown  Sleep  Restart
───────────────
⚙️ Settings
```

**Settings Window** (opens on button click):
```
☐ Run at startup
☐ Run at startup & Speaker ON
☐ Turn off speaker before PC shutdown
☐ Turn off speaker when idle: [15min ▼]
☐ Auto turn on speaker when active
        [Close]
```

**System Tray Menu**:
```
Show Window
Speaker OFF ▼
───────────
Safe Shutdown
Safe Sleep  
Safe Restart
───────────
Quit
```

## Configuration

**config.json**:
```json
{
    "tapo_email": "your@email.com",
    "tapo_password": "yourpass", 
    "plug_ip": "192.168.0.123",
    "delay_after_power_off_sec": 2,
    "timeout_sec": 5
}
```

**settings.json** (auto-generated):
```json
{
    "start_with_windows": false,
    "speaker_on_at_startup": false,
    "watchdog_enabled": false,
    "idle_timer_enabled": false,
    "idle_timer_minutes": 15,
    "idle_auto_on": false
}
```

## Build Standalone EXE

```bash
pip install pyinstaller
pyinstaller gui.py \
    -n DYCTSpeaker \
    -w \
    --icon=assets/ico.ico \
    --add-data "assets;assets" \
    --add-data "config.json;." \
    --add-data "settings.json;."
```

**Result**: `dist/DYCTSpeaker.exe` (single file, portable)

## CLI Fallback (main.py)

```bash
python main.py shutdown    # Speaker OFF → Shutdown
python main.py sleep       # Speaker OFF → Sleep  
python main.py restart     # Speaker OFF → Restart
python main.py off         # Speaker OFF only
python main.py status      # Check status
```

Add `--force` to skip speaker control errors.

## Logs

```
logs/
└── dycts.log     # All actions + errors
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Startup shortcut wrong icon | Delete `DYCTSpeaker.lnk`, toggle "Run at startup" |
| Tapo connection fails | Check IP, credentials, network |
| Hibernate instead of sleep | `powercfg /h off` (Admin) |
| Multiple instances | Single-instance mutex prevents |

## Roadmap

- [x] CLI power management
- [x] **System tray GUI + Settings split**
- [x] Startup integration + idle detection
- [x] Watchdog + single instance
- [ ] Hotkey support
- [ ] macOS/Linux support (I don't need this now btw)


## License

MIT [LICENSE](LICENSE)
