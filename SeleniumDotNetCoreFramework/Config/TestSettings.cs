using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using static SeleniumDotNetCoreFramework.Base.Browser;

namespace SeleniumDotNetCoreFramework.Config
{
    [JsonObject("testSettings")]
    public class TestSettings
    {


        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("aut")]
        public string AUT { get; set; }

        [JsonProperty("env")]
        public string Environment { get; set; }

        [JsonProperty("browser")]
        public BrowserType Browser { get; set; }


        [JsonProperty("testType")]
        public string TestType { get; set; }

        [JsonProperty("isLog")]
        public string IsLog { get; set; }


        [JsonProperty("logPath")]
        public string LogPath { get; set; }

        [JsonProperty("autConnectionString")]
        public string AUTConnectionString { get; set; }
    }
}   

