using Newtonsoft.Json;
using System.Collections.Generic;

namespace SteffBeckers.Abp.Cli.Localization.Models
{
    public class LocalizationFile
    {
        [JsonIgnore]
        public string Path { get; set; }

        [JsonProperty("culture")]
        public string Culture { get; set; }

        [JsonProperty("texts")]
        public Dictionary<string, string> Texts { get; set; }
    }
}