import sys
import os
import json
import threading
import subprocess
import warnings
import time
import locale
from typing import List, Tuple, Dict, Optional

from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw, ImageFilter


def _base_dir():
    # przy .exe PyInstaller rozpakowuje zasoby do _MEIPASS
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = _base_dir()

warnings.filterwarnings("ignore", category=UserWarning)

IS_WIN = sys.platform == "win32"
CREATE_NO_WINDOW = 0x08000000 if IS_WIN else 0

APP_NAME = "Audio & DPI Switcher"
SETDPI_EXE = os.path.join(BASE_DIR, "SetDpi.exe")
APPDATA = os.getenv("APPDATA") or os.path.expanduser("~")
MAP_DIR = os.path.join(APPDATA, "AudioDpiSwitcher")
os.makedirs(MAP_DIR, exist_ok=True)
MAP_FILE = os.path.join(MAP_DIR, "monitor_map.json")

DPI_CHOICES = [100, 125, 150, 175, 200, 225, 250, 300]
MAX_IDX_CHOICES = 8  # ile IDX SetDpi przewidujemy do wyboru

tray_icon: Optional[Icon] = None


# ==================== I18N ====================
def _detect_system_lang() -> str:
    # 1) Spróbuj locale
    try:
        loc_tuple = locale.getlocale()
        loc = loc_tuple[0] if loc_tuple and loc_tuple[0] else ""
    except Exception:
        loc = ""
    loc = (loc or "").lower()
    if IS_WIN:
        try:
            cp = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "[cultureinfo]::InstalledUICulture.Name"],
                capture_output=True, text=True, encoding="utf-8", creationflags=CREATE_NO_WINDOW
            )
            winloc = (cp.stdout or "").strip().lower()
            if winloc:
                loc = winloc
        except Exception:
            pass

    if loc.startswith("pl"):
        return "pl"
    return "en"


TRANSLATIONS = {
    "en": {
        "refreshed_at": "⏱️ Refreshed: {time}",
        "refresh_now": "🔄 Refresh now",

        "default_audio": "🎧 Default audio: {name}",
        "comm_audio": "📞 Communications audio: {name}",
        "default_mic": "🎤 Default mic: {name}",
        "comm_mic": "📞 Communications mic: {name}",

        "monitor_line": "📏 {name} (IDX {idx}, fp …{fp}): {val}%",
        "set_dpi_for": "🖥 Set DPI: {name} (IDX {idx})",

        "index_map_root": "🧭 Index mapping (by EDID)",
        "index_map_item": "{name} → IDX {current} (fp …{fp})",

        "set_speakers_media": "🔊 Set speakers (media)",
        "set_speakers_calls": "🔊 Set speakers (calls)",
        "set_mic_media": "🎙️ Set mic (media)",
        "set_mic_calls": "🎙️ Set mic (calls)",

        "unknown": "(Unknown)",
        "exit": "❌ Exit",
    },
    "pl": {
        "refreshed_at": "⏱️ Odświeżono: {time}",
        "refresh_now": "🔄 Odśwież teraz",

        "default_audio": "🎧 Domyślne audio: {name}",
        "comm_audio": "📞 Komunikacyjne audio: {name}",
        "default_mic": "🎤 Domyślny mikrofon: {name}",
        "comm_mic": "📞 Komunikacyjny mikrofon: {name}",

        "monitor_line": "📏 {name} (IDX {idx}, fp …{fp}): {val}%",
        "set_dpi_for": "🖥 Ustaw DPI: {name} (IDX {idx})",

        "index_map_root": "🧭 Mapowanie indeksów (po EDID)",
        "index_map_item": "{name} → IDX {current} (fp …{fp})",

        "set_speakers_media": "🔊 Ustaw głośniki (multimedia)",
        "set_speakers_calls": "🔊 Ustaw głośniki (rozmowy)",
        "set_mic_media": "🎙️ Ustaw mikrofon (multimedia)",
        "set_mic_calls": "🎙️ Ustaw mikrofon (rozmowy)",

        "unknown": "(Nieznane)",
        "exit": "❌ Wyjście",
    },
}

CURRENT_LANG = _detect_system_lang()

def t(key: str, **kwargs) -> str:
    txt = TRANSLATIONS.get(CURRENT_LANG, TRANSLATIONS["en"]).get(key, key)
    return txt.format(**kwargs) if kwargs else txt


# ==================== CACHE & SYNC ====================
state_lock = threading.RLock()
state = {
    "outs": [],            # [(name,id)]
    "ins": [],             # [(name,id)]
    "defs": {},            # {output, output_comm, input, input_comm}
    "mons": [],            # [{'index':1,'name':'...','fp':'...'}]
    "dpi": {},             # {setdpi_idx: '125'}
    "idx_map": {},         # {fingerprint: idx}
    "last_update": 0.0,    # ts
}


# ==================== POWERSHELL HELPERS ====================
def run_ps(ps_script: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True, text=True, encoding="utf-8",
        creationflags=CREATE_NO_WINDOW
    )


# ==================== AUDIO ====================
def list_audio_devices() -> List[Dict]:
    ps = r"[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-AudioDevice -List | ConvertTo-Json -Depth 3"
    cp = run_ps(ps)
    try:
        data = json.loads(cp.stdout) if cp.stdout else []
        return [data] if isinstance(data, dict) else (data or [])
    except Exception:
        return []


def list_devices_by_type_from_raw(devices, device_type: str) -> List[Tuple[str, str]]:
    def is_type(d):
        df = d.get("Device", {}).get("DataFlow")
        return (device_type == "playback" and df == 0) or (device_type == "recording" and df == 1)
    return [(d.get("Name", ""), d.get("ID", "")) for d in devices if d.get("Name") and is_type(d)]


def get_default_ids() -> Dict[str, Optional[str]]:
    roles = {
        "output": "-Playback",
        "output_comm": "-PlaybackCommunication",
        "input": "-Recording",
        "input_comm": "-RecordingCommunication",
    }
    out = {}
    for key, param in roles.items():
        cp = run_ps(f"[Console]::OutputEncoding=[Text.Encoding]::UTF8; (Get-AudioDevice {param}).ID")
        out[key] = (cp.stdout or "").strip() or None
    return out


def set_default_device(device_id: str, role: str = "console"):
    if not device_id:
        return
    if role == "console":
        cmd = f"Set-AudioDevice -ID '{device_id}' -DefaultOnly"
    elif role == "communications":
        cmd = f"Set-AudioDevice -ID '{device_id}' -CommunicationOnly"
    else:
        cmd = f"Set-AudioDevice -ID '{device_id}'"
    run_ps(cmd)


# ==================== MONITORY & DPI ====================
def set_dpi(percent: int, monitor: Optional[int] = None):
    if not os.path.isfile(SETDPI_EXE):
        return
    args = [SETDPI_EXE, str(percent)]
    if monitor:
        args.append(str(monitor))
    subprocess.run(args, creationflags=CREATE_NO_WINDOW)


def get_dpi_value(monitor: int) -> Optional[str]:
    if not os.path.isfile(SETDPI_EXE):
        return None
    cp = subprocess.run(
        [SETDPI_EXE, "value", str(monitor)],
        capture_output=True, text=True, creationflags=CREATE_NO_WINDOW
    )
    val = (cp.stdout or "").strip()
    return val or None


def list_monitors() -> List[Dict]:
    """
    Zwraca listę fizycznych monitorów wraz z unikalnym 'fp' (EDID):
    [{index:int, name:str, fp:str}]
    """
    ps = r"""
$mons = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID
$i=1
$lst = foreach($m in $mons){
  $name = [Text.Encoding]::UTF8.GetString($m.UserFriendlyName) -replace "\0",""
  $man  = [Text.Encoding]::ASCII.GetString($m.ManufacturerName) -replace "\0",""
  $prod = ($m.ProductCodeID | ForEach-Object { $_.ToString("X2") }) -join ''
  $ser  = [Text.Encoding]::ASCII.GetString($m.SerialNumberID) -replace "\0",""
  if ([string]::IsNullOrWhiteSpace($ser)) { $ser = "NOSN" }
  [PSCustomObject]@{
    Index = $i
    Name = $name
    Manufacturer = $man
    Product = $prod
    Serial = $ser
  }
  $i++
}
$lst | ConvertTo-Json
"""
    cp = run_ps(ps)
    try:
        arr = json.loads(cp.stdout) if cp.stdout else []
        if isinstance(arr, dict):
            arr = [arr]
        result = []
        for m in arr or []:
            idx = int(m.get("Index", 0) or 0)
            name = m.get("Name") or f"Monitor #{idx}"
            man = (m.get("Manufacturer") or "").strip()
            prod = (m.get("Product") or "").strip()
            ser = (m.get("Serial") or "NOSN").strip()
            fp = f"{man}|{prod}|{ser}"  # fingerprint EDID
            result.append({"index": idx, "name": name, "fp": fp})
        return result
    except Exception:
        # fallback: bez EDID – użyjemy nazwy + index jako pseudo-fp (gorsze, ale działa)
        return [{"index": i, "name": f"Monitor #{i}", "fp": f"FAKE||{i}"} for i in (1, 2)]


# ==================== MAPOWANIE (fingerprint → IDX SetDpi) ====================
def load_map() -> Dict[str, int]:
    try:
        with open(MAP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return {str(k): int(v) for k, v in data.items()}
    except Exception:
        return {}


def save_map(m: Dict[str, int]):
    try:
        with open(MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(m, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Nie zapisano mapy monitorów:", e)


def set_mapping(fp: str, idx: int):
    with state_lock:
        new_map = dict(state["idx_map"])
        new_map[fp] = int(idx)
        state["idx_map"] = new_map
    save_map(new_map)
    refresh_state_async()


# ==================== ZBIERANIE STANU (ASYNC NA ŻĄDANIE) ====================
def collect_state():
    """Ciężkie rzeczy robimy tutaj; menu renderuje cache."""
    raw_devices = list_audio_devices()
    outs = list_devices_by_type_from_raw(raw_devices, "playback")
    ins = list_devices_by_type_from_raw(raw_devices, "recording")
    defs = get_default_ids()
    mons = list_monitors()
    idx_map = load_map()

    # odczyt DPI (po mapie fp -> idx SetDpi)
    dpi = {}
    for m in mons:
        idx = idx_map.get(m["fp"], m["index"])
        val = get_dpi_value(idx)
        if val:
            dpi[idx] = val

    with state_lock:
        state["outs"] = outs
        state["ins"] = ins
        state["defs"] = defs
        state["mons"] = mons
        state["idx_map"] = idx_map
        state["dpi"] = dpi
        state["last_update"] = time.time()


def refresh_state_async():
    threading.Thread(target=_refresh_worker, daemon=True).start()


def _refresh_worker():
    try:
        collect_state()
        if tray_icon:
            tray_icon.menu = Menu(lambda: list(_menu_items_dynamic()))
            tray_icon.update_menu()
    except Exception:
        pass


# ==================== HANDLERY MENU ====================
def make_set_dpi_handler(pct: int, idx: int):
    def handler(icon, item):
        def work():
            try:
                set_dpi(pct, idx)
            finally:
                refresh_state_async()
        threading.Thread(target=work, daemon=True).start()
    return handler


def make_set_audio_handler(device_id: str, role: str):
    def handler(icon, item):
        def work():
            try:
                set_default_device(device_id, role)
            finally:
                refresh_state_async()
        threading.Thread(target=work, daemon=True).start()
    return handler


def make_map_choice_handler(fp: str, idx: int):
    def handler(icon, item):
        set_mapping(fp, idx)
    return handler


def exit_handler(icon, item):
    def work():
        try:
            try:
                icon.visible = False
            except Exception:
                pass
            icon.stop()
        finally:
            os._exit(0)
    threading.Thread(target=work, daemon=True).start()


# ==================== MENU (błyskawiczne, z cache) ====================
def _short_fp(fp: str) -> str:
    return fp[-6:] if len(fp) > 6 else fp


def _menu_items_dynamic():
    with state_lock:
        outs = list(state["outs"])
        ins  = list(state["ins"])
        defs = dict(state["defs"])
        mons = list(state["mons"])
        idx_map = dict(state["idx_map"])
        dpi = dict(state["dpi"])
        ts = state["last_update"]

    def find_name(pool, _id):
        _id = (_id or "").strip().lower()
        for n, i in pool:
            if (i or "").strip().lower() == _id:
                return n
        return t("unknown")

    # Nagłówek i odświeżanie
    yield MenuItem(t("refreshed_at", time=time.strftime('%H:%M:%S', time.localtime(ts))), None, enabled=False)
    yield MenuItem(t("refresh_now"), lambda icon, item: refresh_state_async())
    yield Menu.SEPARATOR

    # Audio – etykiety
    yield MenuItem(t("default_audio", name=find_name(outs, defs.get('output'))), None, enabled=False)
    yield MenuItem(t("comm_audio", name=find_name(outs, defs.get('output_comm'))), None, enabled=False)
    yield MenuItem(t("default_mic", name=find_name(ins, defs.get('input'))), None, enabled=False)
    yield MenuItem(t("comm_mic", name=find_name(ins, defs.get('input_comm'))), None, enabled=False)

    # DPI – etykiety i szybkie presety (fingerprint → idx)
    for m in mons:
        idx = idx_map.get(m["fp"], m["index"])
        val = dpi.get(idx, "?")
        yield MenuItem(t("monitor_line", name=m['name'], idx=idx, fp=_short_fp(m['fp']), val=val), None, enabled=False)

    for m in mons:
        idx = idx_map.get(m["fp"], m["index"])
        sub = [MenuItem(f"{pct}%", make_set_dpi_handler(pct, idx)) for pct in DPI_CHOICES]
        yield MenuItem(t("set_dpi_for", name=m['name'], idx=idx), Menu(*sub))

    # Mapowanie indeksów (trwałe po EDID)
    map_menu = []
    max_choice = max(MAX_IDX_CHOICES, len(mons))
    for m in mons:
        current = idx_map.get(m["fp"], m["index"])
        choices = []
        for n in range(1, max_choice + 1):
            label = f"{n} {'✓' if n == current else ''}"
            choices.append(MenuItem(label, make_map_choice_handler(m["fp"], n)))
        map_menu.append(MenuItem(t("index_map_item", name=m["name"], current=current, fp=_short_fp(m["fp"])), Menu(*choices)))
    yield MenuItem(t("index_map_root"), Menu(*map_menu))

    # Szybkie audio – nagłówki
    if outs:
        yield MenuItem(t("set_speakers_media"),
                       Menu(*[MenuItem(n, make_set_audio_handler(i, "console")) for n, i in outs]))
        yield MenuItem(t("set_speakers_calls"),
                       Menu(*[MenuItem(n, make_set_audio_handler(i, "communications")) for n, i in outs]))
    if ins:
        yield MenuItem(t("set_mic_media"),
                       Menu(*[MenuItem(n, make_set_audio_handler(i, "console")) for n, i in ins]))
        yield MenuItem(t("set_mic_calls"),
                       Menu(*[MenuItem(n, make_set_audio_handler(i, "communications")) for n, i in ins]))

    yield Menu.SEPARATOR
    # yield MenuItem("🧩 Ustaw STANDARDOWE", lambda icon, item: threading.Thread(target=set_standard_devices, daemon=True).start()) 
    yield Menu.SEPARATOR
    yield MenuItem(t("exit"), exit_handler)


# ==================== STANDARD AUDIO PROFILE ====================
def set_standard_devices():
    mappings = [
        ("Voicemeeter AUX Input (VB-Audio Voicemeeter VAIO)", "communications"),
        ("Voicemeeter Input (VB-Audio Voicemeeter VAIO)", "console"),
        ("CABLE Output (VB-Audio Virtual Cable)", "console"),
        ("CABLE Output (VB-Audio Virtual Cable)", "communications"),
    ]
    devices = list_audio_devices()
    for target_name, role in mappings:
        for d in devices:
            if (d.get("Name") or "").strip() == target_name:
                set_default_device(d.get("ID"), role)
                break
    refresh_state_async()


# ==================== TRAY ====================
def create_icon_image() -> Image.Image:
    S = 256
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for y in range(S):
        c = 22 + int(20 * (y / S))
        d.line([(0, y), (S, y)], fill=(c, c, c, 255))
    def glow(shape_fn, xy, color, width=8, spread=18, steps=5, fill=None):
        r, g, b = color
        for i in range(steps, 0, -1):
            alpha = int(90 * (i / steps)) 
            off = spread * (i / steps)
            pad = (xy[0]-off, xy[1]-off, xy[2]+off, xy[3]+off)
            shape_fn(pad, outline=(r, g, b, alpha), width=width)
        shape_fn(xy, outline=(r, g, b, 255), width=width)
        if fill:
            shape_fn(xy, fill=fill + (255,))
    border = (24, 24, S-24, S-24)
    glow(d.rounded_rectangle, border, (180, 0, 255), width=10, spread=28, steps=6)
    d.rounded_rectangle(border, outline=(220, 220, 220, 180), width=6, radius=40)
    left  = (58, 78, 126, 186)
    right = (130, 78, 198, 186)
    band_y = 124
    glow(d.ellipse, left,  (0, 255, 230), width=18, spread=22, steps=5)
    glow(d.ellipse, right, (0, 255, 230), width=18, spread=22, steps=5)
    d.line((126, band_y, 130, band_y), fill=(0, 255, 230, 255), width=18) 
    ruler = (80, 188, 176, 208)
    glow(d.rectangle, ruler, (255, 240, 0), width=8, spread=18, steps=4, fill=(255, 210, 0))
    for x in range(86, 176, 16):
        d.line((x, 188, x, 208), fill=(120, 90, 0, 220), width=6)
    img = img.filter(ImageFilter.GaussianBlur(radius=0.3))
    out = img.resize((64, 64), resample=Image.LANCZOS)
    return out


def build_menu_fast():
    refresh_state_async()
    return list(_menu_items_dynamic())


def run_tray():
    global tray_icon
    try:
        collect_state()
    except Exception:
        pass
    tray_icon = Icon(APP_NAME, icon=create_icon_image(), title=APP_NAME)
    tray_icon.menu = Menu(lambda: build_menu_fast())  # szybkie, z cache + async refresh
    tray_icon.run()


# ==================== MAIN ====================
if __name__ == "__main__":
    try:
        run_tray()
    except Exception as e:
        print("Wystąpił błąd:", e)
        import traceback; traceback.print_exc()
