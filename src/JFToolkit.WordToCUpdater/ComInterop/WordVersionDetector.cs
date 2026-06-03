using Microsoft.Win32;

namespace JFToolkit.WordToCUpdater.ComInterop;

/// <summary>
/// Detects installed Word version at runtime without any compile-time coupling.
/// </summary>
internal static class WordVersionDetector
{
    /// <summary>
    /// Probes the registry to determine which Word version is installed.
    /// Returns NotInstalled if no Word installation is found.
    /// </summary>
    public static WordVersion Detect()
    {
        // Check ClickToRun (Office 365 / Office 2016+ C2R)
        foreach (var hive in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, hive);
            using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            
            if (key?.GetValue("VersionToReport") is string versionString
                && Version.TryParse(versionString.Split('.')[0], out var major))
            {
                return major.Major switch
                {
                    16 => WordVersion.Office365,
                    15 => WordVersion.Word2013,
                    _  => WordVersion.Unknown
                };
            }
        }

        // Check traditional MSI installs
        foreach (var hive in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, hive);
            using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot");
            
            if (key?.GetValue("Path") is string)
                return WordVersion.Office365;
        }

        // Last resort: try to create an instance
        try
        {
            var t = Type.GetTypeFromProgID("Word.Application");
            return t != null ? WordVersion.Installed : WordVersion.NotInstalled;
        }
        catch
        {
            return WordVersion.NotInstalled;
        }
    }
}

public enum WordVersion
{
    NotInstalled,
    Installed,      // Detected via ProgID, version unknown
    Word2007,
    Word2010,
    Word2013,
    Word2016,
    Office365,      // ClickToRun, continuously updated
    Unknown         // Installed but version unrecognised
}
