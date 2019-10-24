﻿using SharePoint.Modernization.Scanner.Reports;
using SharePoint.Modernization.Scanner.Telemetry;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace SharePoint.Modernization.Scanner
{
    /// <summary>
    /// SharePoint PnP Modernization scanner
    /// </summary>
    class Program
    {
        private static ScannerTelemetry scannerTelemetry;

        /// <summary>
        /// Main method to execute the program
        /// </summary>
        /// <param name="args">Command line arguments</param>
        [STAThread]
        static void Main(string[] args)
        {
            var options = new Options();

            // Show wizard to help the user with filling the needed scan configuration
            if (args.Length == 0)
            {
                var wizard = new Forms.Wizard(options);
                wizard.ShowDialog();

                if (string.IsNullOrEmpty(options.User) && string.IsNullOrEmpty(options.ClientID))
                {
                    // Trigger validation which will show usage information
                    options.ValidateOptions(args);
                }
            }
            else
            {
                // Validate commandline options
                options.ValidateOptions(args);
            }

            if (options.ExportPaths != null && options.ExportPaths.Count > 0)
            {
                Generator generator = new Generator();
                generator.CreateGroupifyReport(options.ExportPaths);
                generator.CreateListReport(options.ExportPaths);
                generator.CreatePageReport(options.ExportPaths);
                generator.CreatePublishingReport(options.ExportPaths);
                generator.CreateWorkflowReport(options.ExportPaths);
                generator.CreateInfoPathReport(options.ExportPaths);
            }
            else
            {
                try
                {
                    DateTime scanStartDateTime = DateTime.Now;

                    // let's catch unhandled exceptions 
                    AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                    //Instantiate scan job
                    ModernizationScanJob job = new ModernizationScanJob(options)
                    {

                        // I'm debugging
                        //UseThreading = false
                    };

                    scannerTelemetry = job.ScannerTelemetry;

                    job.Execute();

                    // Create reports
                    if (!options.SkipReport)
                    {
                        string workingFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                        List<string> paths = new List<string>
                        {
                            Path.Combine(workingFolder, job.OutputFolder)
                        };

                        var generator = new Generator();

                        generator.CreateGroupifyReport(paths);

                        if (Options.IncludeLists(options.Mode))
                        {
                            generator.CreateListReport(paths);
                        }

                        if (Options.IncludePage(options.Mode))
                        {
                            generator.CreatePageReport(paths);
                        }

                        if (Options.IncludePublishing(options.Mode))
                        {
                            generator.CreatePublishingReport(paths);
                        }

                        if (Options.IncludeWorkflow(options.Mode))
                        {
                            generator.CreateWorkflowReport(paths);
                        }

                        if (Options.IncludeInfoPath(options.Mode))
                        {
                            generator.CreateInfoPathReport(paths);
                        }
                    }

                    TimeSpan duration = DateTime.Now.Subtract(scanStartDateTime);
                    if (scannerTelemetry != null)
                    {
                        scannerTelemetry.LogScanDone(duration);
                    }
                }
                finally
                {
                    if (scannerTelemetry != null)
                    {
                        scannerTelemetry.Flush();
                    }
                }
            }            
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (scannerTelemetry != null)
            {
                scannerTelemetry.LogScanCrash(e.ExceptionObject);
            }
        }
    }
}
