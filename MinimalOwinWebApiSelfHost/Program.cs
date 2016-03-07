using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Add reference to:
using Microsoft.Owin.Hosting;
using System.Data.Entity;
using MinimalOwinWebApiSelfHost.Models;

namespace MinimalOwinWebApiSelfHost
{
    class Program
    {
        static string baseUri = "http://localhost:8080";
        static void Main(string[] args)
        {
            if (args.Contains("-url"))
            {
                baseUri = GetUrl(args);
            }
            using (WebApp.Start(baseUri))
            {
                Console.WriteLine("Server running on {0}", baseUri);
                Console.ReadLine();
            }
        }
        private static string GetUrl(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-url" && args.Length > i + 1)
                {
                    return args[i + 1];
                }
            }
            return baseUri;
        }
    }
}
