using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using Microsoft.Win32;
using System.Linq;
using System.ServiceProcess;

namespace RepairWindowsUpdate
{
    internal class Program
    {
        // Variables globales
        private static string name = "Repair Windows Update";
        private static string version = "1.0.0";
        private static string windowsVersion = "";
        private static string windowsFamily = "";
        private static bool isCompatible = false;

        static void Main(string[] args)
        {
            // Configuration de la console
            SetupConsole();

            // Détecter la version de Windows
            DetectWindowsVersion();

            // Vérifier si l'application est exécutée en tant qu'administrateur
            CheckAdminPrivileges();

            // Afficher les termes et conditions
            ShowTermsAndConditions();

            // Afficher le menu principal
            ShowMainMenu();
        }

        #region Configuration de la console
        static void SetupConsole()
        {
            Console.Title = "Repair Windows Update";
            Console.BackgroundColor = ConsoleColor.Blue;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Clear();
        }
        #endregion

        #region Affichage
        static void PrintHeader(string message)
        {
            Console.Clear();
            Console.WriteLine();
            Console.WriteLine($"{name} [Version: {version}]");
            Console.WriteLine("Repair Windows Update.");
            Console.WriteLine();
            Console.WriteLine(message);
            Console.WriteLine();
        }
        #endregion

        #region Détection du système
        static void DetectWindowsVersion()
        {
            PrintHeader("Detecting Windows version...");

            // Obtenir la version de Windows
            OperatingSystem os = Environment.OSVersion;
            Version ver = os.Version;

            // Formater la version pour correspondre au format du script batch
            string formattedVersion = $"{ver.Major}.{ver.Minor}.{ver.Build}";

            switch (formattedVersion)
            {
                case "5.1.2600":
                    windowsVersion = "Microsoft Windows XP";
                    windowsFamily = "5";
                    isCompatible = true;
                    break;
                case "5.2.3790":
                    windowsVersion = "Microsoft Windows XP Professional x64 Edition";
                    windowsFamily = "5";
                    isCompatible = true;
                    break;
                case "6.0.6000":
                    windowsVersion = "Microsoft Windows Vista";
                    windowsFamily = "6";
                    isCompatible = true;
                    break;
                case "6.0.6001":
                    windowsVersion = "Microsoft Windows Vista SP1";
                    windowsFamily = "6";
                    isCompatible = true;
                    break;
                case "6.0.6002":
                    windowsVersion = "Microsoft Windows Vista SP2";
                    windowsFamily = "6";
                    isCompatible = true;
                    break;
                case "6.1.7600":
                    windowsVersion = "Microsoft Windows 7";
                    windowsFamily = "7";
                    isCompatible = true;
                    break;
                case "6.1.7601":
                    windowsVersion = "Microsoft Windows 7 SP1";
                    windowsFamily = "7";
                    isCompatible = true;
                    break;
                case "6.2.9200":
                    windowsVersion = "Microsoft Windows 10 / 11";
                    windowsFamily = "8";
                    isCompatible = true;
                    break;
                case "6.3.9200":
                case "6.3.9600":
                    windowsVersion = "Microsoft Windows 8.1";
                    windowsFamily = "8";
                    isCompatible = true;
                    break;
                default:
                    // Vérifier si c'est Windows 10
                    if (ver.Major == 10)
                    {
                        windowsVersion = "Microsoft Windows 10";
                        windowsFamily = "10";
                        isCompatible = true;
                    }
                    else
                    {
                        windowsVersion = "Unknown";
                        isCompatible = false;
                    }
                    break;
            }

            PrintHeader($"{windowsVersion} detected...");

            if (!isCompatible)
            {
                Console.WriteLine("    Sorry, this Operative System is not compatible with this tool.");
                Console.WriteLine();
                Console.WriteLine("    An error occurred while attempting to verify your system.");
                Console.WriteLine("    This might be using a business or test version.");
                Console.WriteLine();
                Console.WriteLine("    If not, verify that your system has the correct security fix.");
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                Environment.Exit(0);
            }

            Thread.Sleep(1000);
        }
        #endregion

        #region Vérification des privilèges administrateur
        static void CheckAdminPrivileges()
        {
            PrintHeader("Checking for Administrator elevation.");

            bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent())
                .IsInRole(WindowsBuiltInRole.Administrator);

            if (!isAdmin)
            {
                Console.WriteLine("    You are not running as Administrator.");
                Console.WriteLine("    This tool cannot do its job without elevation.");
                Console.WriteLine();
                Console.WriteLine("    You need to run this tool as Administrator.");
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                Environment.Exit(0);
            }

            Thread.Sleep(1000);
        }
        #endregion

        #region Termes et conditions
        static void ShowTermsAndConditions()
        {
            PrintHeader("Terms and Conditions of Use.");

            Console.WriteLine("    The methods inside this tool modify files and registry settings.");
            Console.WriteLine("    While they are tested and tend to work, We do not take responsibility for");
            Console.WriteLine("    the use of this tool.");
            Console.WriteLine();
            Console.WriteLine("    This tool is provided without warranty. Any damage caused is your");
            Console.WriteLine("    own responsibility.");
            Console.WriteLine();
            Console.WriteLine("Do you want to continue with this process? (Y/N)");

            ConsoleKeyInfo key = Console.ReadKey();
            if (key.Key != ConsoleKey.Y)
            {
                Environment.Exit(0);
            }
        }
        #endregion

        #region Menu principal
        static void ShowMainMenu()
        {
            bool exit = false;

            while (!exit)
            {
                PrintHeader("This tool resets the Windows Update Components.");

                Console.WriteLine("    1. Opens the system protection.");
                Console.WriteLine("    2. Resets the Windows Update Components.");
                Console.WriteLine("    3. Deletes the temporary files in Windows.");
                Console.WriteLine("    4. Opens the Internet Explorer options.");
                Console.WriteLine("    5. Runs Chkdsk on the Windows partition.");
                Console.WriteLine("    6. Runs the System File Checker tool.");
                Console.WriteLine("    7. Scans the image for component store corruption.");
                Console.WriteLine("    8. Checks whether the image has been flagged as corrupted.");
                Console.WriteLine("    9. Performs repair operations automatically.");
                Console.WriteLine("    10. Cleans up the superseded components.");
                Console.WriteLine("    11. Deletes any incorrect registry values.");
                Console.WriteLine("    12. Repairs/Resets Winsock settings.");
                Console.WriteLine("    13. Forces Group Policy Update.");
                Console.WriteLine("    14. Searches Windows updates.");
                Console.WriteLine("    15. Runs SetupDiag (Require .NET Framework 4.6).");
                Console.WriteLine("    16. Explores other local solutions.");
                Console.WriteLine("    17. Explores other online solutions.");
                Console.WriteLine("    18. Downloads the Diagnostic Tools.");
                Console.WriteLine("    19. Restarts your PC.");
                Console.WriteLine();
                Console.WriteLine("                                            H. Help.    0. Close.");
                Console.WriteLine();

                Console.Write("Select an option: ");
                string option = Console.ReadLine();

                switch (option)
                {
                    case "0":
                        exit = true;
                        break;
                    case "1":
                        OpenSystemProtection();
                        break;
                    case "2":
                        ResetComponents();
                        break;
                    case "3":
                        DeleteTempFiles();
                        break;
                    case "4":
                        OpenInternetOptions();
                        break;
                    case "5":
                        RunChkdsk();
                        break;
                    case "6":
                        RunSFC();
                        break;
                    case "7":
                        RunDISMScanHealth();
                        break;
                    case "8":
                        RunDISMCheckHealth();
                        break;
                    case "9":
                        RunDISMRestoreHealth();
                        break;
                    case "10":
                        RunDISMStartComponentCleanup();
                        break;
                    case "11":
                        FixRegistry();
                        break;
                    case "12":
                        ResetWinsock();
                        break;
                    case "13":
                        ForceGroupPolicyUpdate();
                        break;
                    case "14":
                        SearchUpdates();
                        break;
                    case "15":
                        RunSetupDiag();
                        break;
                    case "16":
                        ExploreLocalSolutions();
                        break;
                    case "17":
                        ExploreOnlineSolutions();
                        break;
                    case "18":
                        ShowDiagnosticTools();
                        break;
                    case "19":
                        RestartPC();
                        break;
                    case "H":
                    case "h":
                    case "?":
                        ShowHelp();
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        Thread.Sleep(1000);
                        break;
                }
            }
        }
        #endregion

        #region Functions
        static void OpenSystemProtection()
        {
            PrintHeader("Opening the system protection.");

            if (windowsFamily != "5")
            {
                Process.Start("systempropertiesprotection");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        static void ResetComponents()
        {
            PrintHeader("Stopping the Windows Update services.");

            // Arrêter les services
            RunCommand("net stop bits");
            RunCommand("net stop wuauserv");
            RunCommand("net stop appidsvc");
            RunCommand("net stop cryptsvc");

            // Tuer les processus
            try
            {
                foreach (Process proc in Process.GetProcessesByName("wuauclt"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error killing process: {ex.Message}");
            }

            // Vérifier l'état des services
            PrintHeader("Checking the services status.");

            if (!IsServiceStopped("bits"))
            {
                Console.WriteLine("    Failed to stop the BITS service.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            if (!IsServiceStopped("wuauserv"))
            {
                Console.WriteLine("    Failed to stop the Windows Update service.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            // Service appidsvc peut ne pas exister sur certains systèmes
            if (!IsServiceStopped("appidsvc") && ServiceExists("appidsvc"))
            {
                Console.WriteLine("    Failed to stop the Application Identity service.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                if (windowsFamily != "6") return;
            }

            if (!IsServiceStopped("cryptsvc"))
            {
                Console.WriteLine("    Failed to stop the Cryptographic Services service.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            // Supprimer les fichiers qmgr*.dat
            PrintHeader("Deleting the qmgr*.dat files.");

            string allUsersProfile = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string[] qmgrPaths = {
                Path.Combine(allUsersProfile, @"Application Data\Microsoft\Network\Downloader"),
                Path.Combine(allUsersProfile, @"Microsoft\Network\Downloader")
            };

            foreach (string path in qmgrPaths)
            {
                if (Directory.Exists(path))
                {
                    try
                    {
                        foreach (string file in Directory.GetFiles(path, "qmgr*.dat"))
                        {
                            File.Delete(file);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error deleting qmgr files: {ex.Message}");
                    }
                }
            }

            // Supprimer les anciennes sauvegardes et créer de nouvelles
            PrintHeader("Deleting the old software distribution backup copies.");

            string systemRoot = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string[] backups = {
                Path.Combine(systemRoot, "winsxs\\pending.xml.bak"),
                Path.Combine(systemRoot, "SoftwareDistribution.bak"),
                Path.Combine(systemRoot, "system32\\Catroot2.bak"),
                Path.Combine(systemRoot, "WindowsUpdate.log.bak")
            };

            foreach (string backup in backups)
            {
                try
                {
                    if (File.Exists(backup))
                    {
                        File.Delete(backup);
                    }
                    else if (Directory.Exists(backup))
                    {
                        Directory.Delete(backup, true);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting backup: {ex.Message}");
                }
            }

            PrintHeader("Renaming the software distribution folders.");

            // Renommer les dossiers
            try
            {
                string pendingXml = Path.Combine(systemRoot, "winsxs\\pending.xml");
                if (File.Exists(pendingXml))
                {
                    RunCommand($"takeown /f \"{pendingXml}\"");
                    RunCommand($"attrib -r -s -h /s /d \"{pendingXml}\"");
                    File.Move(pendingXml, pendingXml + ".bak");
                }

                string softwareDistribution = Path.Combine(systemRoot, "SoftwareDistribution");
                if (Directory.Exists(softwareDistribution))
                {
                    RunCommand($"attrib -r -s -h /s /d \"{softwareDistribution}\"");
                    try
                    {
                        Directory.Move(softwareDistribution, softwareDistribution + ".bak");
                    }
                    catch
                    {
                        Console.WriteLine();
                        Console.WriteLine("    Failed to rename the SoftwareDistribution folder.");
                        Console.WriteLine();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        return;
                    }
                }

                string catroot2 = Path.Combine(systemRoot, "system32\\Catroot2");
                if (Directory.Exists(catroot2))
                {
                    RunCommand($"attrib -r -s -h /s /d \"{catroot2}\"");
                    Directory.Move(catroot2, catroot2 + ".bak");
                }

                string windowsUpdateLog = Path.Combine(systemRoot, "WindowsUpdate.log");
                if (File.Exists(windowsUpdateLog))
                {
                    RunCommand($"attrib -r -s -h /s /d \"{windowsUpdateLog}\"");
                    File.Move(windowsUpdateLog, windowsUpdateLog + ".bak");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming directories: {ex.Message}");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            // Réinitialiser les descripteurs de sécurité
            PrintHeader("Reset the BITS service and the Windows Update service to the default security descriptor.");

            RunCommand("sc.exe sdset wuauserv D:(A;CI;CCLCSWRPLORC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)S:(AU;FA;CCDCLCSWRPWPDTLOSDRCWDWO;;;WD)");
            RunCommand("sc.exe sdset bits D:(A;CI;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;SAFA;WDWO;;;BA)");
            RunCommand("sc.exe sdset cryptsvc D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)(A;;CCLCSWRPWPDTLOCRRC;;;SO)(A;;CCLCSWLORC;;;AC)(A;;CCLCSWLORC;;;S-1-15-3-1024-3203351429-2120443784-2872670797-1918958302-2829055647-4275794519-765664414-2751773334)");
            RunCommand("sc.exe sdset trustedinstaller D:(A;CI;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRRC;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;SAFA;WDWO;;;BA)");

            // Ré-enregistrer les DLLs
            PrintHeader("Reregister the BITS files and the Windows Update files.");

            string[] dlls = {
                "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll", "jscript.dll",
                "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll", "msxml6.dll", "actxprxy.dll",
                "softpub.dll", "wintrust.dll", "dssenh.dll", "rsaenh.dll", "gpkcsp.dll", "sccbase.dll",
                "slbcsp.dll", "cryptdlg.dll", "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll",
                "wuapi.dll", "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll",
                "wuweb.dll", "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
            };

            foreach (string dll in dlls)
            {
                RunCommand($"regsvr32.exe /s {dll}");
            }

            // Réinitialiser Winsock
            PrintHeader("Resetting Winsock.");
            RunCommand("netsh winsock reset");

            // Réinitialiser proxy WinHTTP
            PrintHeader("Resetting WinHTTP Proxy.");

            if (windowsFamily == "5")
            {
                RunCommand("proxycfg.exe -d");
            }
            else
            {
                RunCommand("netsh winhttp reset proxy");
            }

            // Configurer les services en automatique
            PrintHeader("Resetting the services as automatics.");

            RunCommand("sc.exe config wuauserv start= auto");
            RunCommand("sc.exe config bits start= delayed-auto");
            RunCommand("sc.exe config cryptsvc start= auto");
            RunCommand("sc.exe config TrustedInstaller start= demand");
            RunCommand("sc.exe config DcomLaunch start= auto");

            // Démarrer les services
            PrintHeader("Starting the Windows Update services.");

            RunCommand("net start bits");
            RunCommand("net start wuauserv");
            RunCommand("net start appidsvc");
            RunCommand("net start cryptsvc");
            RunCommand("net start DcomLaunch");

            // Fin
            PrintHeader("The operation completed successfully.");

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void DeleteTempFiles()
        {
            PrintHeader("Deleting the temporary files in Windows.");

            try
            {
                // Supprimer les fichiers temporaires utilisateur
                string tempPath = Path.GetTempPath();
                foreach (string file in Directory.GetFiles(tempPath, "*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch (Exception)
                    {
                        // Ignorer les erreurs (fichiers verrouillés, etc.)
                    }
                }

                // Supprimer les fichiers temporaires système
                string winTemp = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
                foreach (string file in Directory.GetFiles(winTemp, "*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch (Exception)
                    {
                        // Ignorer les erreurs
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting temporary files: {ex.Message}");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void OpenInternetOptions()
        {
            PrintHeader("Opening the Internet Explorer options.");
            Process.Start("InetCpl.cpl");
        }

        static void RunChkdsk()
        {
            PrintHeader("Check the file system and file system metadata of a volume for logical and physical errors (CHKDSK.exe).");

            string systemDrive = Path.GetPathRoot(Environment.SystemDirectory);
            int result = RunCommand($"chkdsk {systemDrive} /f /r");

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunSFC()
        {
            PrintHeader("Scan your system files and to repair missing or corrupted system files (SFC.exe).");

            int result = 0;
            if (windowsFamily != "5")
            {
                result = RunCommand("sfc /scannow");
            }
            else
            {
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
            }

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunDISMScanHealth()
        {
            PrintHeader("Scan the image for component store corruption (The DISM /ScanHealth argument).");

            int result = 0;
            if (windowsFamily == "8" || windowsFamily == "10")
            {
                result = RunCommand("Dism.exe /Online /Cleanup-Image /ScanHealth");
            }
            else
            {
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
            }

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunDISMCheckHealth()
        {
            PrintHeader("Check whether the image has been flagged as corrupted by a failed process and whether the corruption can be repaired (The DISM /CheckHealth argument).");

            int result = 0;
            if (windowsFamily == "8" || windowsFamily == "10")
            {
                result = RunCommand("Dism.exe /Online /Cleanup-Image /CheckHealth");
            }
            else
            {
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
            }

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunDISMRestoreHealth()
        {
            PrintHeader("Scan the image for component store corruption, and then perform repair operations automatically (The DISM /RestoreHealth argument).");

            int result = 0;
            if (windowsFamily == "8" || windowsFamily == "10")
            {
                result = RunCommand("Dism.exe /Online /Cleanup-Image /RestoreHealth");
            }
            else
            {
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
            }

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunDISMStartComponentCleanup()
        {
            PrintHeader("Clean up the superseded components and reduce the size of the component store (The DISM /StartComponentCleanup argument).");

            int result = 0;
            if (windowsFamily == "8" || windowsFamily == "10")
            {
                result = RunCommand("Dism.exe /Online /Cleanup-Image /StartComponentCleanup");
            }
            else
            {
                Console.WriteLine("Sorry, this option is not available on this Operative System.");
            }

            if (result == 0)
            {
                Console.WriteLine();
                Console.WriteLine("The operation completed successfully.");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("An error occurred during operation.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void FixRegistry()
        {
            // Créer un dossier de sauvegarde
            string now = DateTime.Now.ToString("yyyyMMddHHmm");
            string backupFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Backup", "Regedit", now);

            PrintHeader($"Making a backup of the Registry in: {backupFolder}");

            if (!Directory.Exists(backupFolder))
            {
                Directory.CreateDirectory(backupFolder);
            }

            // Vérifier si la sauvegarde existe déjà
            if (File.Exists(Path.Combine(backupFolder, "HKLM.reg")))
            {
                Console.WriteLine("An unexpected error has occurred.");
                Console.WriteLine();
                Console.WriteLine("    Changes were not carried out in the registry.");
                Console.WriteLine("    Will try it later.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            // Créer une sauvegarde du registre
            RunCommand($"reg Export HKCR \"{Path.Combine(backupFolder, "HKCR.reg")}\"");
            RunCommand($"reg Export HKCU \"{Path.Combine(backupFolder, "HKCU.reg")}\"");
            RunCommand($"reg Export HKLM \"{Path.Combine(backupFolder, "HKLM.reg")}\"");
            RunCommand($"reg Export HKU \"{Path.Combine(backupFolder, "HKU.reg")}\"");
            RunCommand($"reg Export HKCC \"{Path.Combine(backupFolder, "HKCC.reg")}\"");

            // Vérifier la sauvegarde
            PrintHeader("Checking the backup.");

            if (!File.Exists(Path.Combine(backupFolder, "HKLM.reg")))
            {
                Console.WriteLine("An unexpected error has occurred.");
                Console.WriteLine();
                Console.WriteLine("    Something went wrong.");
                Console.WriteLine("    You manually create a backup of the registry before continuing.");
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("The operation completed successfully.");
                Console.WriteLine();
            }

            // Supprimer les valeurs dans le registre
            PrintHeader("Deleting values in the Registry.");

            // Supprimer les clés du registre
            string[] keysToDelete = {
                @"HKCU\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo",
                @"HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate",
                @"HKCU\Software\Microsoft\WindowsSelfHost",
                @"HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate",
                @"HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate",
                @"HKLM\Software\Microsoft\WindowsSelfHost",
                @"HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate",
                @"HKLM\COMPONENTS\PendingXmlIdentifier",
                @"HKLM\COMPONENTS\NextQueueEntryIndex",
                @"HKLM\COMPONENTS\AdvancedInstallersNeedResolving",
                @"HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
            };

            foreach (string keyPath in keysToDelete)
            {
                try
                {
                    string[] parts = keyPath.Split('\\');
                    string hive = parts[0];
                    string subKey = string.Join("\\", parts.Skip(1));

                    RegistryKey baseKey = null;
                    switch (hive)
                    {
                        case "HKLM":
                            baseKey = Registry.LocalMachine;
                            break;
                        case "HKCU":
                            baseKey = Registry.CurrentUser;
                            break;
                        case "HKCR":
                            baseKey = Registry.ClassesRoot;
                            break;
                        case "HKU":
                            baseKey = Registry.Users;
                            break;
                    }

                    if (baseKey != null)
                    {
                        baseKey.DeleteSubKeyTree(subKey, false);
                    }
                }
                catch (Exception)
                {
                    // Ignorer les erreurs (clé non trouvée, etc.)
                }
            }

            Console.WriteLine("The operation completed successfully.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void ResetWinsock()
        {
            PrintHeader("Resetting Winsock.");

            RunCommand("netsh winsock reset");
            RunCommand("netsh int ip reset");
            RunCommand("ipconfig /release");
            RunCommand("ipconfig /renew");
            RunCommand("ipconfig /flushdns");

            Console.WriteLine();
            Console.WriteLine("The operation completed successfully.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void ForceGroupPolicyUpdate()
        {
            PrintHeader("Forcing Group Policy Update.");

            RunCommand("gpupdate /force");

            Console.WriteLine();
            Console.WriteLine("The operation completed successfully.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void SearchUpdates()
        {
            PrintHeader("Searching Windows updates.");

            if (windowsFamily == "10" || windowsFamily == "8")
            {
                Process.Start("ms-settings:windowsupdate");
            }
            else
            {
                RunCommand("wuauclt /resetauthorization /detectnow");
                Process.Start("control", "update");
            }

            Console.WriteLine();
            Console.WriteLine("The operation completed successfully.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void RunSetupDiag()
        {
            PrintHeader("Running SetupDiag tool.");

            string setupDiagPath = Path.Combine(Environment.CurrentDirectory, "SetupDiag.exe");

            if (!File.Exists(setupDiagPath))
            {
                Console.WriteLine("SetupDiag.exe not found. Downloading...");

                try
                {
                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile("https://download.microsoft.com/download/1/1/1/111c347e-b7de-4510-8e62-a2f046efcc48/SetupDiag.exe", setupDiagPath);
                    }

                    Console.WriteLine("Download completed successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to download SetupDiag.exe: {ex.Message}");
                    Console.WriteLine();
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    return;
                }
            }

            try
            {
                Console.WriteLine("Running SetupDiag...");
                RunCommand(setupDiagPath);

                Console.WriteLine();
                Console.WriteLine("SetupDiag completed. Check the logs in the current directory.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running SetupDiag: {ex.Message}");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void ExploreLocalSolutions()
        {
            bool exit = false;

            while (!exit)
            {
                PrintHeader("Explore other local solutions.");

                Console.WriteLine("    1. Open the Solutions troubleshooter.");
                Console.WriteLine("    2. Start the Windows Update troubleshooter.");
                Console.WriteLine("    3. Run the Windows 10 Upgrade Assistant.");
                Console.WriteLine("    4. Open the Settings app.");
                Console.WriteLine("    5. Open the Control Panel.");
                Console.WriteLine("    6. Open the Event Viewer.");
                Console.WriteLine("    7. Open the services console.");
                Console.WriteLine("    8. Open the registry editor.");
                Console.WriteLine("    9. Open the Microsoft Support and Recovery Assistant.");
                Console.WriteLine();
                Console.WriteLine("                                            0. Back to the Main menu.");
                Console.WriteLine();

                Console.Write("Select an option: ");
                string option = Console.ReadLine();

                switch (option)
                {
                    case "0":
                        exit = true;
                        break;
                    case "1":
                        Process.Start("control.exe", "/name Microsoft.Troubleshooting");
                        break;
                    case "2":
                        if (windowsFamily == "10")
                        {
                            RunCommand("msdt.exe /id WindowsUpdateDiagnostic");
                        }
                        else
                        {
                            Process.Start("control.exe", "/name Microsoft.Troubleshooting");
                        }
                        break;
                    case "3":
                        if (windowsFamily != "10")
                        {
                            Console.WriteLine("This option is only available for Windows 10.");
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            string upgradeAssistantPath = Path.Combine(Environment.CurrentDirectory, "Windows10Upgrade.exe");

                            if (!File.Exists(upgradeAssistantPath))
                            {
                                Console.WriteLine("Windows10Upgrade.exe not found. Downloading...");

                                try
                                {
                                    using (WebClient client = new WebClient())
                                    {
                                        client.DownloadFile("https://go.microsoft.com/fwlink/?LinkID=799445", upgradeAssistantPath);
                                    }

                                    Console.WriteLine("Download completed successfully.");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Failed to download Windows 10 Upgrade Assistant: {ex.Message}");
                                    Console.WriteLine();
                                    Console.WriteLine("Press any key to continue...");
                                    Console.ReadKey();
                                    break;
                                }
                            }

                            Process.Start(upgradeAssistantPath);
                        }
                        break;
                    case "4":
                        if (windowsFamily == "10")
                        {
                            Process.Start("ms-settings:");
                        }
                        else
                        {
                            Console.WriteLine("This option is only available for Windows 10.");
                            Thread.Sleep(2000);
                        }
                        break;
                    case "5":
                        Process.Start("control.exe");
                        break;
                    case "6":
                        Process.Start("eventvwr.msc");
                        break;
                    case "7":
                        Process.Start("services.msc");
                        break;
                    case "8":
                        Process.Start("regedit.exe");
                        break;
                    case "9":
                        string saraPath = Path.Combine(Environment.CurrentDirectory, "SaraSetup.exe");

                        if (!File.Exists(saraPath))
                        {
                            Console.WriteLine("SaraSetup.exe not found. Downloading...");

                            try
                            {
                                using (WebClient client = new WebClient())
                                {
                                    client.DownloadFile("https://aka.ms/SaRA_CommandLineVersionFiles", saraPath);
                                }

                                Console.WriteLine("Download completed successfully.");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to download Microsoft Support and Recovery Assistant: {ex.Message}");
                                Console.WriteLine();
                                Console.WriteLine("Press any key to continue...");
                                Console.ReadKey();
                                break;
                            }
                        }

                        Process.Start(saraPath);
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        Thread.Sleep(1000);
                        break;
                }
            }
        }

        static void ExploreOnlineSolutions()
        {
            bool exit = false;

            while (!exit)
            {
                PrintHeader("Explore other online solutions.");

                Console.WriteLine("    1. Visit the Microsoft Support website.");
                Console.WriteLine("    2. Run the online Windows Update troubleshooter.");
                Console.WriteLine("    3. Check for known issues with Windows updates.");
                Console.WriteLine("    4. Visit the Windows community forums.");
                Console.WriteLine("    5. Check your update history online.");
                Console.WriteLine();
                Console.WriteLine("                                            0. Back to the Main menu.");
                Console.WriteLine();

                Console.Write("Select an option: ");
                string option = Console.ReadLine();

                switch (option)
                {
                    case "0":
                        exit = true;
                        break;
                    case "1":
                        Process.Start("https://support.microsoft.com/en-us/windows");
                        break;
                    case "2":
                        Process.Start("https://support.microsoft.com/en-us/help/4027322/windows-update-troubleshooter");
                        break;
                    case "3":
                        Process.Start("https://docs.microsoft.com/en-us/windows/release-health/");
                        break;
                    case "4":
                        Process.Start("https://answers.microsoft.com/en-us/windows/forum/windows_10-update");
                        break;
                    case "5":
                        Process.Start("https://account.microsoft.com/devices/");
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        Thread.Sleep(1000);
                        break;
                }
            }
        }

        static void ShowDiagnosticTools()
        {
            bool exit = false;

            while (!exit)
            {
                PrintHeader("Download Microsoft diagnostic tools.");

                Console.WriteLine("    1. Windows Update Reset Script");
                Console.WriteLine("    2. SetupDiag Tool");
                Console.WriteLine("    3. Microsoft Support and Recovery Assistant");
                Console.WriteLine("    4. Windows 10 Upgrade Assistant");
                Console.WriteLine("    5. Windows Update Troubleshooter");
                Console.WriteLine("    6. Fix Windows Update issues");
                Console.WriteLine();
                Console.WriteLine("                                            0. Back to the Main menu.");
                Console.WriteLine();

                Console.Write("Select an option: ");
                string option = Console.ReadLine();

                switch (option)
                {
                    case "0":
                        exit = true;
                        break;
                    case "1":
                        Process.Start("https://support.microsoft.com/en-us/topic/use-the-system-file-checker-tool-to-repair-missing-or-corrupted-system-files-79aa86cb-ca52-166a-92a3-966e85d4094e");
                        break;
                    case "2":
                        Process.Start("https://docs.microsoft.com/en-us/windows/deployment/upgrade/setupdiag");
                        break;
                    case "3":
                        Process.Start("https://support.microsoft.com/en-us/office/about-the-microsoft-support-and-recovery-assistant-e90bb691-c2a7-4697-a94f-88836856c72f");
                        break;
                    case "4":
                        Process.Start("https://www.microsoft.com/en-us/software-download/windows10");
                        break;
                    case "5":
                        Process.Start("https://support.microsoft.com/en-us/help/4027322/windows-update-troubleshooter");
                        break;
                    case "6":
                        Process.Start("https://support.microsoft.com/en-us/topic/windows-update-common-errors-and-mitigation-e9b8f1f6-05dd-f045-5588-90f2ad524696");
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        Thread.Sleep(1000);
                        break;
                }
            }
        }

        static void RestartPC()
        {
            PrintHeader("Restarting your PC.");

            Console.WriteLine("Are you sure you want to restart your PC? (Y/N)");
            ConsoleKeyInfo key = Console.ReadKey();

            if (key.Key == ConsoleKey.Y)
            {
                try
                {
                    Process.Start("shutdown.exe", "/r /t 10 /c \"Windows Update Reset Tool - Restart required\"");
                    Console.WriteLine();
                    Console.WriteLine("Your PC will restart in 10 seconds...");
                    Console.WriteLine();
                    Console.WriteLine("Press any key to cancel restart...");

                    if (Console.ReadKey().Key != ConsoleKey.Escape)
                    {
                        Process.Start("shutdown.exe", "/a");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error restarting: {ex.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void ShowHelp()
        {
            PrintHeader("Help and information.");

            Console.WriteLine("This tool helps resolve issues with Windows Update by performing various troubleshooting steps:");
            Console.WriteLine();
            Console.WriteLine("1. System Protection: Create a restore point before making changes");
            Console.WriteLine("2. Reset Components: Reset Windows Update components to default settings");
            Console.WriteLine("3. Delete Temp Files: Clean temporary files that may interfere with updates");
            Console.WriteLine("4. Internet Options: Adjust settings that may affect downloads");
            Console.WriteLine("5-10. System Repair: Various commands to check and repair system files");
            Console.WriteLine("11. Fix Registry: Removes registry entries that may cause update issues");
            Console.WriteLine("12. Reset Winsock: Repairs network configuration issues");
            Console.WriteLine("13. Group Policy Update: Forces group policy refresh");
            Console.WriteLine("14. Search Updates: Checks for new Windows updates");
            Console.WriteLine("15-18. Advanced Tools: Various diagnostic and repair utilities");
            Console.WriteLine("19. Restart PC: Completes repairs that require a restart");
            Console.WriteLine();
            Console.WriteLine("For best results, run the options in order for your specific issue.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
        #endregion

        #region Helper Methods
        static int RunCommand(string command)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c {command}",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = false
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    Console.WriteLine(output);
                    if (!string.IsNullOrEmpty(error))
                    {
                        Console.WriteLine(error);
                    }

                    return process.ExitCode;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running command: {ex.Message}");
                return -1;
            }
        }

        static bool IsServiceStopped(string serviceName)
        {
            try
            {
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    return sc.Status == ServiceControllerStatus.Stopped;
                }
            }
            catch
            {
                return true; // Considering not found as stopped
            }
        }

        static bool ServiceExists(string serviceName)
        {
            try
            {
                ServiceController[] services = ServiceController.GetServices();
                return services.Any(s => s.ServiceName.ToLower() == serviceName.ToLower());
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}