# Did You Close the Speaker? 🔊

Safe power management for active studio monitors connected to a **TP-Link Tapo P115** smart plug.

Turns off your speakers before shutting down, sleeping, or restarting your PC — preventing potentially harmful power-off pops on active monitors.

## Why?

Active monitors (like the iLoud Micro/MTM series) can produce audible pops when power is cut abruptly. This tool ensures the smart plug powering your speakers is turned off *before* any system power action, giving the monitors a clean shutdown.

## Quick Start

```bash
# 1. Clone
git clone https://github.com/yourusername/did-you-close-the-speaker.git
cd did-you-close-the-speaker

# 2. Install
pip install -r requirements.txt

# 3. Configure
cp config.example.json config.json
# Edit config.json with your Tapo credentials and plug IP

# 4. Use
python main.py shutdown    # Speaker OFF → PC shutdown
python main.py sleep       # Speaker OFF → PC sleep
python main.py restart     # Speaker OFF → PC restart
python main.py off         # Speaker OFF only
python main.py on          # Speaker ON only
python main.py status      # Check plug status
```

## Commands

| Command | Description |
|---------|-------------|
| `shutdown` | Turn off speaker plug, wait, then shutdown |
| `sleep` | Turn off speaker plug, wait, then sleep |
| `restart` | Turn off speaker plug, wait, then restart |
| `off` | Turn off speaker plug only |
| `on` | Turn on speaker plug only |
| `status` | Show current plug status |

All power commands accept `--force` to proceed even if the plug control fails:

```bash
python main.py shutdown --force
```

## Configuration

Copy `config.example.json` to `config.json`:

```json
{
    "tapo_email": "your_tapo_email@example.com",
    "tapo_password": "your_tapo_password",
    "plug_ip": "192.168.0.123",
    "delay_after_power_off_sec": 2,
    "timeout_sec": 5
}
```

| Key | Description |
|-----|-------------|
| `tapo_email` | Your TP-Link / Tapo account email |
| `tapo_password` | Your TP-Link / Tapo account password |
| `plug_ip` | Local IP address of your Tapo P115 |
| `delay_after_power_off_sec` | Seconds to wait after turning off plug before power action |
| `timeout_sec` | Timeout for Tapo API calls |

> **Tip:** Assign a static IP to your P115 in your router's DHCP settings for reliability.

## Desktop Shortcut (Windows)

For quick access, create a shortcut on your desktop:

1. Right-click Desktop → New → Shortcut
2. Location: `pythonw.exe "C:\path\to\did-you-close-the-speaker\main.py" shutdown`
3. Name it "Safe Shutdown"

Or if using a PyInstaller build:

```bash
pip install pyinstaller
pyinstaller --onefile --name dycts main.py
# Creates dist/dycts.exe
```

Then point the shortcut to `dist/dycts.exe shutdown`.

## Sleep Note

Windows `SetSuspendState` may hibernate instead of sleep if hibernate is enabled. To ensure sleep mode:

```powershell
# Run as Administrator (one-time)
powercfg /h off
```

## Logs

All actions are logged to `logs/dycts.log`.

## Roadmap

- [x] v1: CLI with safe shutdown / sleep / restart / manual on-off
- [ ] v1.5: PyInstaller exe build
- [ ] v2: System tray app with quick actions
- [ ] v3: Windows shutdown event watchdog (auto-detect shutdown and intervene)

## License

MIT
