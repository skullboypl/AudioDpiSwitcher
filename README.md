# AudioDpiSwitcher

**Audio & DPI Switcher** – A lightweight Python-based **system tray application** for quickly switching:

- **Default audio devices** (media / communications)
- **Microphones** (media / communications)
- **Display DPI scaling** per monitor

It automatically detects connected hardware, supports **persistent monitor index mapping via EDID**, remembers settings, and includes **multi-language support** (EN/PL).

---

## ✨ Features

- 🎧 **Quick audio switching**:
  - Set default playback device (multimedia)
  - Set default playback device (communications)
  - Set default recording device (microphone – multimedia)
  - Set default recording device (microphone – communications)

- 🖥 **Per-monitor DPI scaling**:
  - Change DPI from predefined presets: 100%, 125%, 150%, 175%, 200%, 225%, 250%, 300%
  - Read current DPI per monitor
  - Map monitor EDID fingerprint to SetDpi index for consistent results

- 🔍 **Auto hardware detection**:
  - Lists all connected audio devices
  - Lists all physical monitors with EDID info (Manufacturer, Product, Serial)
  - Fallback pseudo-fingerprint if EDID not available

- 🌐 **Multi-language UI**:
  - English (default)
  - Polish (automatically selected if system language is Polish)

- ⚡ **Fast & responsive**:
  - Menu built from cached state for instant opening
  - Background refresh without blocking UI
  - Runs silently in the Windows system tray

---

## 📦 How It Works

1. On launch, the app collects:
   - Available audio devices
   - Current default devices
   - List of monitors and their fingerprints
   - Current DPI for each mapped monitor
2. Data is cached and used to render the tray menu instantly.
3. Any action (set device, change DPI) is performed in a background thread.
4. The app automatically refreshes its state after each change.

---

## 🛠 Technical Details

- **Language**: Python 3.x  
- **Dependencies**:
  - `pystray` – system tray integration
  - `Pillow` – tray icon drawing
  - Windows PowerShell – for audio and monitor information
- **Platform**: Windows 10/11
- **Additional tool**: `SetDpi.exe` for DPI control

**Data Storage**:
- Monitor mapping saved in:  
  `%APPDATA%\AudioDpiSwitcher\monitor_map.json`

---

## 🚀 Installation & Usage

### From Release (recommended)
1. Go to the [Releases](../../releases) section.
2. Download the latest `.exe` version.
3. This build is **EV Code Signed** – should run without SmartScreen warnings.
4. Run the application – it will appear in the system tray.

### From Source
1. Install Python 3.9+ and pip.
2. Install dependencies:
   ```bash
   pip install pystray pillow
3. Place SetDpi.exe in the same folder as the script.
4. Run: `python audio_taskbar_switcher.pyw`
---

## 📋 Menu Overview

- **Refreshed at** – Last state refresh time  
- **Refresh now** – Manual refresh  
- **Default Audio / Communications Audio / Default Mic / Communications Mic** – shows current default devices  
- **Per Monitor DPI** – shows DPI for each monitor with quick preset selection  
- **Index Mapping (by EDID)** – assign SetDpi index for each monitor fingerprint (persistent mapping)  
- **Quick Audio Switch** – set playback/microphone devices directly from the menu  
- **Exit** – quit the application  

---

## 🛠 Technical Details

- **Language**: Python 3.x  
- **Dependencies**:
  - `pystray` – system tray integration
  - `Pillow` – tray icon drawing
  - Windows PowerShell – for audio and monitor information
- **Platform**: Windows 10/11
- **Additional tool**: [`SetDPI`](https://github.com/imniko/SetDPI) for DPI control  
  (included in the release build – no need to download separately)

**Data Storage**:
- Monitor mapping saved in:  
  `%APPDATA%\AudioDpiSwitcher\monitor_map.json`

---

## 🙏 Acknowledgements 

- **[SetDPI by imniko](https://github.com/imniko/SetDPI)** – lightweight Windows tool used for reading and changing per-monitor DPI scaling.  
  Included in this application to provide DPI management functionality.

---


## 📄 License

This project is released under the **MIT License**.  
You are free to use, modify, and distribute it, provided that copyright notice is included.

---

## 📷 Screenshots

*(To be added)*  
 
---

## 🔒 Security Notice

The Windows binary provided in the **[Releases](../../releases)** section is signed with an  
**Extended Validation (EV) Code Signing Certificate**.  
This means:
- **No SmartScreen warnings** on Windows 10/11
- Verified publisher information
- Safer and more trusted installation experience

---

## 💡 Tips

- The application stores monitor mappings in:  
  `%APPDATA%\AudioDpiSwitcher\monitor_map.json`  
  You can delete this file to reset DPI index mapping.

- If you change connected monitors or audio devices,  
  use **Refresh now** in the menu to update the device list.

---
