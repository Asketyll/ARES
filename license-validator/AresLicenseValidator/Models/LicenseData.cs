using Newtonsoft.Json;

namespace AresLicenseValidator.Models
{
    public class LicenseData
    {
        [JsonProperty("company")]
        public string Company { get; set; }

        [JsonProperty("domain")]
        public string Domain { get; set; }

        [JsonProperty("installed_by")]
        public string InstalledBy { get; set; }

        [JsonProperty("installation_date")]
        public string InstallationDate { get; set; }

        [JsonProperty("license_key")]
        public string LicenseKey { get; set; }

        [JsonProperty("environment_hash")]
        public string EnvironmentHash { get; set; }

        [JsonProperty("authorized_users")]
        public string[] AuthorizedUsers { get; set; }

        [JsonProperty("max_users")]
        public int MaxUsers { get; set; }

        [JsonProperty("signature")]
        public string Signature { get; set; }

        [JsonProperty("version")]
        public string Version { get; set; }
    }
}