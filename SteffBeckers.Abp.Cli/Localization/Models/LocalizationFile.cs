using Newtonsoft.Json;
using System.Collections.Generic;

namespace SteffBeckers.Abp.Cli.Localization.Models
{
    public class LocalizationFile
    {
        [JsonIgnore]
        public string Path { get; set; }

        public string Culture { get; set; }

        public Dictionary<string, string> Texts { get; set; }
    }
}