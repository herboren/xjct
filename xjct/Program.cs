using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace xjct
{
    internal class Program
    {
        private const string _whitelist = @"^([cC]:\\|\\\\)([^\\/:*? "" <>|\r\n]+\\)*[^\\/:*?""<>|\r\n]*$";

        static public string flag = string.Empty;
        static private ConfigureExcelApplication openExcelConnectivity = null;
        static private OpenExcelWorkBook openExcelWorkBook = null; 

        static private Workbook _workbook = null;
        static private Worksheet _worksheet = null;
        static private Worksheets[] _worksheets = null;
        static private Application _configuredApplication = null;

        /// <summary>
        /// Create asycnhronous task to speed up the process. This will help in instance
        /// Where data exceed nth amount of rows and may take more thean the mimimun amount
        /// of processing power to complete.
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static async Task Main(string[] args)
        {
            // Introduction
            Console.WriteLine("\n\n\tThis tool is designed to simplify the process of converting large\r\n\t" +
                "Excel datasets into JSON format, allowing you to easily convert\r\n\t" +
                "Excel data into JSON, manipulate and automate it as needed.");

            // Get File path`
            var userInput = Console.ReadLine();

            // Validation needed before moving forward.
            // Here we check if the absolute file path meets the requirements
            if (args.Length != 0)
            {
                // If absolute path passes security checks
                if (ValidateFilePath(args[0]))
                {
                    // Checklist:
                    Console.WriteLine($"");

                    // Configure excel application (complete)
                    flag = ConfigureExcelApplication.ConfigureApplication(args[0]) ? "✓" : "✘";
                    Console.WriteLine($"\n\tApplication Configured {flag}");

                    // Open Excel Workbook
                    Console.WriteLine($"\n\tWorkbook Connected {flag}");
                }
            }
            else
            {
                // We can proceed if all is well
                if (ValidateFilePath(userInput))
                {
                    // await 
                }
            }
        }

        /// <summary>
        /// Used for whitelisting or validating file paths in the context of
        /// security validatiion context
        /// <param name="filename"></param>
        /// <returns></returns>
        static public bool ValidateFilePath(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new WhiteListException("File not found or moved.");
            }

            if (!Path.IsPathRooted(filename))
            {
                throw new WhiteListException("Invalid file path, absolute path is required");
            }

            if (string.IsNullOrEmpty(filename))
            {
                throw new WhiteListException("File path cannot be empty.");
            }

            if (!Regex.IsMatch(filename, _whitelist))
            {
                throw new WhiteListException("Invalid characters identified in file path.");
            }

            // If all checks pass, return true to indicate a valid path
            return true;
        }
    }
    /* public async Task ReadExcelDataTable()
    {
        await Task.Run(() =>
        {

        });

        // Task
    }*/
}