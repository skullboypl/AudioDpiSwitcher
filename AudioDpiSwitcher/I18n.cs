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
        ["settings"] = "Settings…",
        ["windows_settings"] = "Windows Settings",
        ["set_defaults_tray"] = "🔊 Set as defaults",
        ["exit"] = "Exit",
        ["unknown"] = "(Unknown)",
        ["settings_title"] = "Audio & DPI Switcher – Settings",
        ["group_audio"] = "Audio",
        ["group_dpi"] = "DPI / Monitors",
        ["group_language"] = "Language",
        ["label_output"] = "Default playback (media)",
        ["label_output_comm"] = "Default playback (communications)",
        ["label_input"] = "Default microphone (media)",
        ["label_input_comm"] = "Default microphone (communications)",
        ["auto_restore"] = "Auto-restore preferred devices when Windows changes them",
        ["apply"] = "Apply",
        ["close"] = "Close",
        ["lang_auto"] = "System default",
        ["lang_en"] = "English",
        ["lang_pl"] = "Polish",
        ["lang_de"] = "German",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "SetDpi index",
        ["set_dpi"] = "Set DPI",
        ["save_map"] = "Save index map",
        ["none"] = "(none)",
        ["refresh_devices"] = "Refresh devices",
        ["use_current_defaults"] = "Use current as defaults",
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
        ["settings"] = "Ustawienia…",
        ["windows_settings"] = "Ustawienia Windows",
        ["set_defaults_tray"] = "🔊 Ustaw domyślne",
        ["exit"] = "Wyjście",
        ["unknown"] = "(Nieznane)",
        ["settings_title"] = "Audio & DPI Switcher – Ustawienia",
        ["group_audio"] = "Audio",
        ["group_dpi"] = "DPI / Monitory",
        ["group_language"] = "Język",
        ["label_output"] = "Domyślne odtwarzanie (multimedia)",
        ["label_output_comm"] = "Domyślne odtwarzanie (komunikacja)",
        ["label_input"] = "Domyślny mikrofon (multimedia)",
        ["label_input_comm"] = "Domyślny mikrofon (komunikacja)",
        ["auto_restore"] = "Auto-przywracanie preferowanych urządzeń, gdy Windows je zmieni",
        ["apply"] = "Zastosuj",
        ["close"] = "Zamknij",
        ["lang_auto"] = "Domyślny systemowy",
        ["lang_en"] = "Angielski",
        ["lang_pl"] = "Polski",
        ["lang_de"] = "Niemiecki",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "Indeks SetDpi",
        ["set_dpi"] = "Ustaw DPI",
        ["save_map"] = "Zapisz mapowanie",
        ["none"] = "(brak)",
        ["refresh_devices"] = "Odśwież urządzenia",
        ["use_current_defaults"] = "Pobierz aktualne jako domyślne",
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
        ["settings"] = "Einstellungen…",
        ["windows_settings"] = "Windows-Einstellungen",
        ["set_defaults_tray"] = "🔊 Als Standard setzen",
        ["exit"] = "Beenden",
        ["unknown"] = "(Unbekannt)",
        ["settings_title"] = "Audio & DPI Switcher – Einstellungen",
        ["group_audio"] = "Audio",
        ["group_dpi"] = "DPI / Monitore",
        ["group_language"] = "Sprache",
        ["label_output"] = "Standardwiedergabe (Medien)",
        ["label_output_comm"] = "Standardwiedergabe (Kommunikation)",
        ["label_input"] = "Standardmikrofon (Medien)",
        ["label_input_comm"] = "Standardmikrofon (Kommunikation)",
        ["auto_restore"] = "Bevorzugte Geräte automatisch wiederherstellen, wenn Windows sie ändert",
        ["apply"] = "Übernehmen",
        ["close"] = "Schließen",
        ["lang_auto"] = "Systemstandard",
        ["lang_en"] = "Englisch",
        ["lang_pl"] = "Polnisch",
        ["lang_de"] = "Deutsch",
        ["monitor"] = "Monitor",
        ["dpi_preset"] = "DPI",
        ["idx_map"] = "SetDpi-Index",
        ["set_dpi"] = "DPI setzen",
        ["save_map"] = "Indexzuordnung speichern",
        ["none"] = "(keine)",
        ["refresh_devices"] = "Geräte aktualisieren",
        ["use_current_defaults"] = "Aktuelle als Standard übernehmen",
    };
}
