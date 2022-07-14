using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace MDS
{
    class AuthorizationInfo
    {
        private class Info
        {
            public string service;         // The service the user used last.
            public DateTime date;          // When was the last action.
        }

        private List<string> services;  // List of available services.
        private Info info;              // Information for JSON file.

        public AuthorizationInfo()
        {
            services = new List<string>();


            services.Add("Google");
            services.Add("Yahoo");
            services.Add("MailRu");
            services.Add("Rambler");
        }

        public void CreateAuthFile(string path)
        {
            if (File.Exists(path))
            {
                info = JsonConvert.DeserializeObject<Info>(File.ReadAllText(path));
            }
            else
            {
                info = new Info();
            }
        }

        public void Save(string service)
        {
            info.date = info.date.Date;
            info.service = service;

            // Serialize JSON directly to a file.
            using (StreamWriter file = File.CreateText(@"authorization_info.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, info);
            }
        }

        public string CurrentServise
        {
            get
            {
                return info.service;
            }
        }

        public static class Services
        {
            public const string Google = "Google";
            public const string Yahoo = "Yahoo";
            public const string MailRu = "MailRu";
            public const string Rambler = "Rambler";
        }
    }
}
