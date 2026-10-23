using System.Globalization;

namespace AudioDpiSwitcher;

internal enum AppLang
{
    Auto,
    En,
    Pl,
    De,
}

internal static class I18n
{
    private static AppLang _lang = AppLang.Auto;

    public static AppLang Current => _lang;

    public static void SetLanguage(string? langCode)
    {
        _lang = (langCode ?? "auto").Trim().ToLowerInvariant() switch
        {
            "pl" => AppLang.Pl,
            "en" => AppLang.En,
            "de" => AppLang.De,
            _ => AppLang.Auto,
        };
    }

    public static AppLang Resolve()
    {
        if (_lang != AppLang.Auto)
            return _lang;

        var ui = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.ToLowerInvariant();
        return ui switch
        {
            "pl" => AppLang.Pl,
            "de" => AppLang.De,
            _ => AppLang.En,
        };
    }

    public static string T(string key, params (string name, object value)[] args)
    {
        var dict = Resolve() switch
        {
            AppLang.Pl => Pl,
            AppLang.De => De,
            _ => En,
        };

        if (!dict.TryGetValue(key, out var text))
            text = En.TryGetValue(key, out var fallback) ? fallback : key;

        foreach (var (name, value) in args)
            text = text.Replace("{" + name + "}", value?.ToString() ?? "", StringComparison.Ordinal);

        return text;
    }

    private static readonly Dictionary<string, string> En = new()
    {
        ["app_name"] = "Audio & DPI Switcher",
        ["refreshed_at"] = "Refreshed: {time}",
        ["refresh_now"] = "Refresh now",
        ["default_audio"] = "Default audio: {name}",
        ["comm_audio"] = "Communications audio: {name}",
        ["default_mic"] = "Default mic: {name}",
        ["comm_mic"] = "Communications mic: {name}",
        ["monitor_line"] = "{name} (IDX {idx}): {val}%",
        ["settings"] = "Settings",
        ["windows_settings"] = "Windows sound",
        ["set_defaults_tray"] = "Remember current audio",
        ["exit"] = "Exit",
        ["unknown"] = "(Unknown)",
        ["settings_title"] = "Audio & DPI Switcher",
        ["group_audio"] = "Preferred audio",
        ["audio_hint"] = "These devices are put back if Windows changes them.",
        ["group_dpi"] = "Display scaling",
        ["group_language"] = "Language",
        ["label_output"] = "Speakers - media",
        ["label_output_comm"] = "Speakers - calls",
        ["label_input"] = "Microphone - media",
        ["label_input_comm"] = "Microphone - calls",
        ["auto_restore"] = "Restore these devices when Windows changes them",
        ["apply"] = "Save",
        ["close"] = "Close",
        ["lang_auto"] = "System default",
        ["lang_en"] = "English",
        ["lang_pl"] = "Polish",
        ["lang_de"] = "German",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "SetDpi index",
        ["set_dpi"] = "Apply scale",
        ["save_map"] = "Save index map",
        ["none"] = "(none)",
        ["refresh_devices"] = "Refresh",
        ["use_current_defaults"] = "Use as preferred",
        ["tab_preferred"] = "Preferred",
        ["tab_current"] = "Current",
        ["tab_display"] = "Display",
        ["tab_app"] = "App",
        ["preferred_hint"] = "Choose the devices you want to keep. Saved ones come back if Windows changes them.",
        ["current_hint"] = "These are the devices Windows is using right now.",
        ["display_hint"] = "Scaling applies to the selected monitor. Resolution is shown for reference.",
        ["app_hint"] = "Interface language and updates from GitHub Releases.",
        ["auto_update"] = "Install automatically when a newer GitHub release is available",
        ["check_update"] = "Check for updates",
        ["update_status"] = "This version: {current}. GitHub: {latest}",
        ["update_ok"] = "You are up to date ({version}).",
        ["update_fail"] = "Could not reach GitHub.",
        ["update_dev"] = "This is a development copy, so the exe will not be replaced.",
        ["update_downloading"] = "Downloading {version}...",
        ["update_restarting"] = "Installing {version}. The app will restart.",
    };

    private static readonly Dictionary<string, string> Pl = new()
    {
        ["app_name"] = "Audio & DPI Switcher",
        ["refreshed_at"] = "Odświeżono: {time}",
        ["refresh_now"] = "Odśwież teraz",
        ["default_audio"] = "Domyślne audio: {name}",
        ["comm_audio"] = "Komunikacyjne audio: {name}",
        ["default_mic"] = "Domyślny mikrofon: {name}",
        ["comm_mic"] = "Komunikacyjny mikrofon: {name}",
        ["monitor_line"] = "{name} (IDX {idx}): {val}%",
        ["settings"] = "Ustawienia",
        ["windows_settings"] = "Dźwięk Windows",
        ["set_defaults_tray"] = "Zapamiętaj obecne audio",
        ["exit"] = "Wyjście",
        ["unknown"] = "(Nieznane)",
        ["settings_title"] = "Audio & DPI Switcher",
        ["group_audio"] = "Preferowane audio",
        ["audio_hint"] = "Te urządzenia wracają, gdy Windows sam je zmieni.",
        ["group_dpi"] = "Skalowanie ekranu",
        ["group_language"] = "Język",
        ["label_output"] = "Głośniki - multimedia",
        ["label_output_comm"] = "Głośniki - rozmowy",
        ["label_input"] = "Mikrofon - multimedia",
        ["label_input_comm"] = "Mikrofon - rozmowy",
        ["auto_restore"] = "Przywracaj te urządzenia, gdy Windows je zmieni",
        ["apply"] = "Zapisz",
        ["close"] = "Zamknij",
        ["lang_auto"] = "Domyślny systemowy",
        ["lang_en"] = "Angielski",
        ["lang_pl"] = "Polski",
        ["lang_de"] = "Niemiecki",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "Indeks SetDpi",
        ["set_dpi"] = "Ustaw skalę",
        ["save_map"] = "Zapisz mapowanie",
        ["none"] = "(brak)",
        ["refresh_devices"] = "Odśwież",
        ["use_current_defaults"] = "Użyj jako preferowane",
        ["tab_preferred"] = "Preferowane",
        ["tab_current"] = "Aktualne",
        ["tab_display"] = "Ekran",
        ["tab_app"] = "Aplikacja",
        ["preferred_hint"] = "Wybierz urządzenia, które mają zostać. Zapisane wracają, gdy Windows je zmieni.",
        ["current_hint"] = "To są urządzenia wybrane w tej chwili w Windows.",
        ["display_hint"] = "Skala dotyczy wybranego monitora. Rozdzielczość jest tylko informacją.",
        ["app_hint"] = "Język interfejsu i aktualizacje z GitHub Releases.",
        ["auto_update"] = "Instaluj automatycznie, gdy na GitHubie jest nowsza wersja",
        ["check_update"] = "Sprawdź aktualizacje",
        ["update_status"] = "Ta wersja: {current}. GitHub: {latest}",
        ["update_ok"] = "Masz aktualną wersję ({version}).",
        ["update_fail"] = "Nie udało się połączyć z GitHubem.",
        ["update_dev"] = "To kopia deweloperska, więc plik exe nie zostanie podmieniony.",
        ["update_downloading"] = "Pobieranie {version}...",
        ["update_restarting"] = "Instaluję {version}. Aplikacja uruchomi się ponownie.",
    };

    private static readonly Dictionary<string, string> De = new()
    {
        ["app_name"] = "Audio & DPI Switcher",
        ["refreshed_at"] = "Aktualisiert: {time}",
        ["refresh_now"] = "Jetzt aktualisieren",
        ["default_audio"] = "Standard-Audio: {name}",
        ["comm_audio"] = "Kommunikations-Audio: {name}",
        ["default_mic"] = "Standard-Mikrofon: {name}",
        ["comm_mic"] = "Kommunikations-Mikrofon: {name}",
        ["monitor_line"] = "{name} (IDX {idx}): {val}%",
        ["settings"] = "Einstellungen",
        ["windows_settings"] = "Windows-Sound",
        ["set_defaults_tray"] = "Aktuelles Audio merken",
        ["exit"] = "Beenden",
        ["unknown"] = "(Unbekannt)",
        ["settings_title"] = "Audio & DPI Switcher",
        ["group_audio"] = "Bevorzugtes Audio",
        ["audio_hint"] = "Diese Geräte werden zurückgesetzt, wenn Windows sie ändert.",
        ["group_dpi"] = "Anzeige-Skalierung",
        ["group_language"] = "Sprache",
        ["label_output"] = "Lautsprecher - Medien",
        ["label_output_comm"] = "Lautsprecher - Anrufe",
        ["label_input"] = "Mikrofon - Medien",
        ["label_input_comm"] = "Mikrofon - Anrufe",
        ["auto_restore"] = "Diese Geräte wiederherstellen, wenn Windows sie ändert",
        ["apply"] = "Speichern",
        ["close"] = "Schließen",
        ["lang_auto"] = "Systemstandard",
        ["lang_en"] = "Englisch",
        ["lang_pl"] = "Polnisch",
        ["lang_de"] = "Deutsch",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "SetDpi-Index",
        ["set_dpi"] = "Skalierung anwenden",
        ["save_map"] = "Indexzuordnung speichern",
        ["none"] = "(keine)",
        ["refresh_devices"] = "Aktualisieren",
        ["use_current_defaults"] = "Als bevorzugt verwenden",
        ["tab_preferred"] = "Bevorzugt",
        ["tab_current"] = "Aktuell",
        ["tab_display"] = "Anzeige",
        ["tab_app"] = "App",
        ["preferred_hint"] = "Geräte wählen, die bleiben sollen. Gespeicherte kommen zurück, wenn Windows sie ändert.",
        ["current_hint"] = "Das sind die Geräte, die Windows gerade verwendet.",
        ["display_hint"] = "Die Skalierung gilt für den gewählten Monitor. Die Auflösung dient nur zur Orientierung.",
        ["app_hint"] = "Sprache und Updates über GitHub Releases.",
        ["auto_update"] = "Automatisch installieren, wenn auf GitHub eine neuere Version liegt",
        ["check_update"] = "Nach Updates suchen",
        ["update_status"] = "Diese Version: {current}. GitHub: {latest}",
        ["update_ok"] = "Du hast die aktuelle Version ({version}).",
        ["update_fail"] = "GitHub ist nicht erreichbar.",
        ["update_dev"] = "Das ist eine Entwicklerkopie, die exe wird nicht ersetzt.",
        ["update_downloading"] = "Lade {version} herunter...",
        ["update_restarting"] = "Installiere {version}. Die App startet neu.",
    };
}
