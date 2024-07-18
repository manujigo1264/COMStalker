#pragma warning disable CA1416 // Suppress platform compatibility warnings

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;

namespace COMStalker
{
    class Program
    {
        // Define ANSI color codes
        private const string ColorReset = "\x1b[0m";
        private const string ColorRed = "\x1b[31m";
        private const string ColorGreen = "\x1b[32m";
        private const string ColorYellow = "\x1b[33m";
        private const string ColorBlue = "\x1b[34m";

        public static string usage = $"{ColorYellow}Usage: COMStalker.exe <-inproc|-localserver>{ColorReset}";

        public struct COMServer
        {
            public string CLSID;
            public string ServerPath;
            public string Type;
        }

        static void Main(string[] args)
        {
            // Enable ANSI color support on Windows
            EnableAnsiSupport();

            // Display the creator's name
            Console.WriteLine($"{ColorBlue}COMStalker by 0xTron{ColorReset}");
            Console.WriteLine();

            try
            {
                var servers = ProcessArguments(args);

                if (servers == null)
                {
                    Console.WriteLine(usage);
                    return;
                }

                string[] defaultMethods = { "Equals", "GetHashCode", "GetType", "ToString" };

                Console.WriteLine($"{ColorBlue}COM Servers Information{ColorReset}");
                Console.WriteLine($"{ColorGreen}======================={ColorReset}");

                foreach (var server in servers)
                {
                    Console.WriteLine($"{ColorYellow}CLSID: {ColorReset}{server.CLSID}");
                    Console.WriteLine($"{ColorYellow}Path: {ColorReset}{server.ServerPath}");
                    Console.WriteLine($"{ColorYellow}Type: {ColorReset}{server.Type}");
                    Console.WriteLine($"{ColorGreen}-----------------------{ColorReset}");

                    if (server.ServerPath.ToLower().Contains("mscoree.dll"))
                    {
                        PrintDotNetAssemblyMethods(server, defaultMethods);
                    }

                    Console.WriteLine($"{ColorGreen}======================={ColorReset}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ColorRed}[-] An error occurred: {ex.Message}{ColorReset}");
            }
        }

        private static List<COMServer> ProcessArguments(string[] args)
        {
            if (args.Length == 0)
            {
                return WMICollection("InprocServer32").Concat(WMICollection("LocalServer32")).ToList();
            }
            else if (args[0].ToLower() == "-inproc")
            {
                return WMICollection("InprocServer32");
            }
            else if (args[0].ToLower() == "-localserver")
            {
                return WMICollection("LocalServer32");
            }
            return null;
        }

        private static void PrintDotNetAssemblyMethods(COMServer server, string[] defaultMethods)
        {
            var assemblyPath = Registry.GetValue($"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\CLSID\\{server.CLSID}\\InprocServer32\\1.0.0.0", "CodeBase", null) as string;

            if (assemblyPath != null)
            {
                Console.WriteLine($".NET Assembly: {assemblyPath}");

                var assemblyType = Type.GetTypeFromCLSID(Guid.Parse(server.CLSID));
                if (assemblyType != null)
                {
                    foreach (var method in assemblyType.GetMethods())
                    {
                        if (!defaultMethods.Contains(method.Name))
                        {
                            Console.WriteLine($"  Method: {method.Name}");
                        }
                    }
                }
            }
        }

        static List<COMServer> WMICollection(string type)
        {
            var comServers = new List<COMServer>();

            try
            {
                var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_ClassicCOMClassSetting");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    var serverPath = GetServerPath(queryObj, type);
                    if (!string.IsNullOrEmpty(serverPath) && File.Exists(serverPath))
                    {
                        comServers.Add(new COMServer
                        {
                            CLSID = queryObj["ComponentId"]?.ToString() ?? string.Empty,
                            ServerPath = serverPath,
                            Type = type
                        });
                    }
                }
            }
            catch (ManagementException ex)
            {
                Console.WriteLine($"{ColorRed}[-] An error occurred while querying for WMI data: {ex.Message}{ColorReset}");
            }

            return comServers.OrderBy(x => x.ServerPath).ToList();
        }

        private static string? GetServerPath(ManagementObject queryObj, string type)
        {
            var svrObj = Convert.ToString(queryObj[type]);
            if (svrObj == null)
                return null;

            var svr = Environment.ExpandEnvironmentVariables(svrObj).Trim('"');

            if (!svr.ToLower().Contains(@"c:\windows\"))
            {
                return svr;
            }

            return null;
        }

        private static void EnableAnsiSupport()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var handle = GetStdHandle(STD_OUTPUT_HANDLE);
                var mode = 0;
                GetConsoleMode(handle, ref mode);
                SetConsoleMode(handle, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
            }
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll")]
        private static extern bool GetConsoleMode(IntPtr hConsoleHandle, ref int lpMode);

        [DllImport("kernel32.dll")]
        private static extern bool SetConsoleMode(IntPtr hConsoleHandle, int dwMode);

        private const int STD_OUTPUT_HANDLE = -11;
        private const int ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004;
    }
}

#pragma warning restore CA1416 // Restore platform compatibility warnings
