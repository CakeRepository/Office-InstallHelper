using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeInstallHelper
{
    class Program
    {
        // Define a class to receive parsed values
        

        public static readonly string _uninstallC2r = Path.Combine(Environment.CurrentDirectory, "Scripts", "uninstallPrep.vbs");
        public static readonly string _pathToSetup = Path.Combine(Environment.CurrentDirectory, "Install", "setup.exe");
        public static readonly string _pathToConfiguration = Path.Combine(Environment.CurrentDirectory, "Install", "ConfigurationBasicBusiness.xml");

        static void Main(string[] args)
        {
            removeClickToRun();

            installOffice();

            var versions = detectVersion();

            foreach(var version in versions)
            {
                Console.WriteLine(version);
            }

            Console.ReadKey();
            
        }

        public static List<string> detectVersion()
        {
            string officeLocation = @"SOFTWARE\Microsoft\Office\";
            RegistryKey key = Registry.LocalMachine.OpenSubKey(officeLocation, true);

            var officeVersion = new List<string>();

            using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(officeLocation))
            {
                foreach (var subKey in registryKey.GetSubKeyNames())
                {
                    switch (subKey)
                    {
                        case "16.0":
                            officeVersion.Add("Office 365 or 2016");
                            break;
                        case "15.0":
                            officeVersion.Add("Office 2013");
                            break;
                        case "14.0":
                            officeVersion.Add("Office 2010");
                            break;
                        case "12.0":
                            officeVersion.Add("Office 2007");
                            break;
                        default:
                            break;
                    }
                }
            }
            return officeVersion;
        }

        public static void removeClickToRun()
        {
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cscript.exe",
                Arguments = "//B //Nologo " + _uninstallC2r + " ALL /Quiet /NoCancel /Force /OSE",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (process != null)
            {
                process.WaitForExit();
            }
        }
        public static void installOffice()
        {
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = _pathToSetup,
                Arguments = " /Configure " + _pathToConfiguration,
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (process != null)
            {
                process.WaitForExit();
            }
        }
    }
}
