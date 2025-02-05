// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Identity.Abstractions;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Configuration;
using System.Threading.Tasks;
using Configuration = System.Configuration.Configuration;

namespace daemon_console
{
    /// <summary>
    /// This sample shows how to query the Microsoft Graph from a daemon application
    /// </summary>
    class Program
    {
        public static IHostBuilder CreateHostBuilder(string[] args) =>
       Host.CreateDefaultBuilder(args)
           .ConfigureAppConfiguration((context, config) =>
           {
               config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
           })
           .ConfigureServices((context, services) =>
           {
               services.AddSingleton<ExchangeServiceHelper>();
           });

        static async Task Main(string[] args)
        {
            var host = CreateHostBuilder(args).Build();

            var serviceHelper = host.Services.GetRequiredService<ExchangeServiceHelper>();

            
            // Read and display emails
            var emails = await serviceHelper.ReadEmailsAsync();
            foreach (var email in emails)
            {
                Console.WriteLine($"Subject: {email.Subject}");
                Console.WriteLine($"From: {email.From}");
                Console.WriteLine($"Received: {email.DateTimeReceived}");
                Console.WriteLine($"Body: {email.Body.Text}");
                Console.WriteLine(new string('-', 50));
            }

            host.Run();

        }
    }
}
