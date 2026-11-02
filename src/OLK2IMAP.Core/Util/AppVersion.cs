using System.Reflection;

namespace OLK2IMAP.Core.Util;

/// <summary>Version shown in the client title bar / status line.</summary>
public static class AppVersion
{
    public static string Informational =>
        Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString()
        ?? "0.0.0";

    /// <summary>Short display string, e.g. "OLK2Any 0.1.0".</summary>
    public static string Display => $"OLK2Any {Informational.Split('+')[0]}";
}
