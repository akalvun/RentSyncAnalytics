using System;
using System.IO;
using System.Xml.Serialization;

namespace RentSync.AddIn
{
    /// <summary>
    /// Every path the add-in writes to, in one place. ThisAddIn and the
    /// settings dialog must agree on these, and duplicating the
    /// Path.Combine chain in two files is how they stop agreeing.
    /// </summary>
    public static class AppPaths
    {
        public static string RootDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "RentSync");
            }
        }

        public static string LogDirectory
        {
            get { return Path.Combine(RootDirectory, "logs"); }
        }

        public static string SettingsFile
        {
            get { return Path.Combine(RootDirectory, "settings.xml"); }
        }
    }

    /// <summary>
    /// User-editable settings that survive an Excel restart.
    ///
    /// Serialised with <see cref="XmlSerializer"/> deliberately: it ships in
    /// the .NET Framework BCL, so the Framework-bound add-in layer gains no
    /// package dependency for something this small. The API token is NOT
    /// here — secrets belong in Windows Credential Manager, never in a file
    /// this class can round-trip.
    /// </summary>
    [Serializable]
    [XmlRoot("RentSyncSettings")]
    public sealed class UserSettings
    {
        /// <summary>Read the bundled sample JSON instead of calling the API.</summary>
        public bool UseDemoData { get; set; }

        /// <summary>Base address of the REST API, used when demo mode is off.</summary>
        public string ApiBaseUrl { get; set; }

        /// <summary>How long a cached snapshot is still considered fresh.</summary>
        public int CacheFreshnessMinutes { get; set; }

        public UserSettings()
        {
            // Defaults match the shipped demo configuration.
            UseDemoData = true;
            ApiBaseUrl = "https://api.rentsync.example/v1/";
            CacheFreshnessMinutes = 30;
        }

        /// <summary>
        /// Never throws. A missing, unreadable, or corrupt settings file
        /// falls back to defaults: a bad file must not stop Excel from
        /// starting, and the user can always fix it from the dialog.
        /// </summary>
        public static UserSettings Load()
        {
            try
            {
                var path = AppPaths.SettingsFile;
                if (!File.Exists(path))
                    return new UserSettings();

                using (var stream = File.OpenRead(path))
                {
                    var serializer = new XmlSerializer(typeof(UserSettings));
                    var loaded = serializer.Deserialize(stream) as UserSettings;
                    return loaded != null ? loaded.Sanitised() : new UserSettings();
                }
            }
            catch
            {
                return new UserSettings();
            }
        }

        /// <summary>
        /// Writes via a temporary file and then replaces the original, so an
        /// interrupted save cannot leave a half-written file behind.
        /// </summary>
        public void Save()
        {
            var path = AppPaths.SettingsFile;
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            var temp = path + ".tmp";
            using (var stream = File.Create(temp))
            {
                var serializer = new XmlSerializer(typeof(UserSettings));
                serializer.Serialize(stream, this);
            }

            if (File.Exists(path))
                File.Delete(path);
            File.Move(temp, path);
        }

        /// <summary>Clamp values that a hand-edited file could put out of range.</summary>
        private UserSettings Sanitised()
        {
            if (string.IsNullOrWhiteSpace(ApiBaseUrl))
                ApiBaseUrl = new UserSettings().ApiBaseUrl;
            if (CacheFreshnessMinutes < 1) CacheFreshnessMinutes = 1;
            if (CacheFreshnessMinutes > 10080) CacheFreshnessMinutes = 10080; // one week
            return this;
        }
    }
}
