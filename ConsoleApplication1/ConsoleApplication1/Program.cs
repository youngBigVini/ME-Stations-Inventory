using System;
using System.IO;
using System.Management;
using NationalInstruments.VisaNS;
using System.Collections.Generic;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> aliasPowerSupplyLeftSide = new List<string>();
            List<string> aliasPowerSupplyRightSide = new List<string>();
            List<string> addressPowerSupplyLeftSide = new List<string>();
            List<string> addressPowerSupplyRightSide = new List<string>();
            List<string> aliasTestset = new List<string>();
            string directoryLog = "";
            string filePathLog = "";
            string enabledFieldsLog = "";
            string enableLog = "";
            string fileNameLog = "";
            string enableManufacturer = "";
            string enableSystemModel = "";
            string enableSystemName = "";
            string enableWindowsVersion = "";
            string enableBuild = "";
            string enableInstalledRAM = "";
            string enableRAMs = "";
            string enableHDModel = "";
            string enableHDSize = "";
            string enablePowerSupplyLeft = "";
            string enablePowerSupplyRight = "";
            string enableTestSet = "";
            List<string> messageLogList = new List<string>();
            List<string> logList = new List<string>();
            Dictionary<string, string> found = new Dictionary<string, string>();
            // Obtendo o diretório atual da aplicação
            string currentDirectory = Directory.GetCurrentDirectory();
            string newDirectoryName = "ConfigFiles";
            string newDirectoryPath = Path.Combine(currentDirectory, newDirectoryName);

            if (!Directory.Exists(newDirectoryPath))
            {
                // Criando o novo diretório
                Directory.CreateDirectory(newDirectoryPath);               
            }

            string filePath = Path.Combine(newDirectoryPath, "config.cfg");
            if (!File.Exists(filePath))
            {
                
                string[] configLines = {
                "[ALIAS POWER SUPPLY]",
                "aliasPowerSupplyLeftSide=WZ_K6206_1",
                "aliasPowerSupplyRightSide=WZ_K6206_2",
                "[ADDRESS POWER SUPPLY]",
                "addressPowerSupplyLeftSide=GPIB0::4::INSTR",
                "addressPowerSupplyRightSide=GPIB0::6::INSTR",
                "[ALIAS TEST SET]",
                "aliasTestset=RS_CMW500",
                "[DIRECTORY LOG]",
                "directoryLog=C:\\TestSW",
                "fileNameLog=log.csv",
                "[ENABLE LOGGING FIELDS]",
                "Manufacturer=true",
                "ComputerModel=true",
                "ComputerName=true",
                "WindowsVersion=true",
                "Build=false",
                "RAMsize=true",
                "RAMsModel=true",                
                "HDModel=true",
                "HDSize=true",
                "PowerSupplyLeft=true",
                "PowerSupplyRight=true",
                "TestSet=true",
                "enableLog=false"
                };
                File.WriteAllLines(filePath, configLines);
            }
            if (File.Exists(filePath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        if (line.Contains("aliasPowerSupplyLeftSide"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["aliasPowerSupplyLeftSide"] = parts[1].Trim();
                            }
                        }
                        if (line.Contains("aliasPowerSupplyRightSide"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["aliasPowerSupplyRightSide"] = parts[1].Trim();
                            }
                        }
                        if (line.Contains("addressPowerSupplyLeftSide"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["addressPowerSupplyLeftSide"] = parts[1].Trim();
                            }
                        }
                        if (line.Contains("addressPowerSupplyRightSide"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["addressPowerSupplyRightSide"] = parts[1].Trim();
                            }
                        }
                        if (line.Contains("aliasTestset"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["aliasTestset"] = parts[1].Trim();
                                Console.WriteLine($"aliasTest Set: {found["aliasTestset"]}");
                            }
                        }
                        if (line.Contains("directoryLog"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["directoryLog"] = parts[1].Trim();
                                directoryLog = found["directoryLog"];

                            }
                        }
                        if (line.Contains("fileNameLog"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["fileNameLog"] = parts[1].Trim().Replace("\"", "");
                            }
                        }

                        if (line.Contains("Manufacturer"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["Manufacturer"] = parts[1].Trim();
                                enableManufacturer = found["Manufacturer"];
                            }
                        }

                        if (line.Contains("ComputerModel"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["ComputerModel"] = parts[1].Trim();
                                enableSystemModel = found["ComputerModel"];
                            }
                        }
                        if (line.Contains("ComputerName"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["ComputerName"] = parts[1].Trim();
                                enableSystemName = found["ComputerName"];
                            }
                        }
                        if (line.Contains("WindowsVersion"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["WindowsVersion"] = parts[1].Trim();
                                enableWindowsVersion = found["WindowsVersion"];
                            }
                        }
                        if (line.Contains("Build"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["Build"] = parts[1].Trim();
                                enableBuild = found["Build"];
                            }
                        }
                        if (line.Contains("RAMsize"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["InstalledRAM"] = parts[1].Trim();
                                enableInstalledRAM = found["InstalledRAM"];
                            }
                        }
                        if (line.Contains("RAMsModel"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["RAMsModel"] = parts[1].Trim();
                                enableRAMs = found["RAMsModel"];
                            }
                        }
                        
                        
                        if (line.Contains("HDModel"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["HDModel"] = parts[1].Trim();
                                enableHDModel = found["HDModel"];
                            }
                        }
                        if (line.Contains("HDSize"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["HDSize"] = parts[1].Trim();
                                enableHDSize = found["HDSize"];
                            }
                        }
                        if (line.Contains("PowerSupplyLeft"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["PowerSupplyLeft"] = parts[1].Trim();
                                enablePowerSupplyLeft = found["PowerSupplyLeft"];
                            }
                        }
                        if (line.Contains("PowerSupplyRight"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["PowerSupplyRight"] = parts[1].Trim();
                                enablePowerSupplyRight = found["PowerSupplyRight"];
                            }
                        }
                        if (line.Contains("TestSet"))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                found["TestSet"] = parts[1].Trim();
                                enableTestSet = found["TestSet"];
                            }
                        }
                        if (line.Contains("enableLog"))
                        {
                            string[] parts = line.Split('=');
                            Console.WriteLine(parts);
                            if (parts.Length == 2)
                            {
                                found["enableLog"] = parts[1].Trim();
                                enableLog = found["enableLog"];
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao ler o arquivo: {ex.Message}");
                    Console.WriteLine("Press any key to continue the aplication.");
                    Console.ReadKey();
                }
            }
            
            try
            {

                if ((found != null) && (enableLog == "true"))
                {

                    string[] partsLeft = found["aliasPowerSupplyLeftSide"].Split(',');
                    foreach (string str in partsLeft)
                    {
                        aliasPowerSupplyLeftSide.Add(str);
                    }

                    string[] partsRight = found["aliasPowerSupplyRightSide"].Split(',');
                    foreach (string str in partsRight)
                    {
                        aliasPowerSupplyRightSide.Add(str);
                    }
                    string[] partsAddressLeft = found["addressPowerSupplyLeftSide"].Split(',');
                    foreach (string str in partsAddressLeft)
                    {
                        addressPowerSupplyLeftSide.Add(str);
                    }

                    string[] partsAddressRight = found["addressPowerSupplyRightSide"].Split(',');
                    foreach (string str in partsAddressRight)
                    {
                        addressPowerSupplyRightSide.Add(str);
                    }

                    string[] partsAliasTestset = found["aliasTestset"].Split(',');
                    foreach (string str in partsAliasTestset)
                    {
                        aliasTestset.Add(str);
                    }

                    directoryLog = found["directoryLog"];
                    fileNameLog = found["fileNameLog"];
                    filePathLog = Path.Combine(directoryLog, fileNameLog);

                    List<string> enabledFields = new List<string>();
                    enabledFields.Add("DateAndTime");
                    if (enableManufacturer == "true")
                        enabledFields.Add("Manufacturer");
                   
                    if (enableSystemModel == "true")
                        enabledFields.Add("ComputerModel");
                   
                    if (enableSystemName == "true")
                        enabledFields.Add("ComputerName");
                    
                    if (enableWindowsVersion == "true")
                        enabledFields.Add("WindowsVersion");
                    
                    if (enableBuild == "true")
                        enabledFields.Add("BuildComputer");
                    
                    if (enableInstalledRAM == "true")
                        enabledFields.Add("RAMsize");
                    
                    if (enableRAMs == "true")
                        enabledFields.Add("RAMmodels");
                    
                    if (enableHDModel == "true")
                        enabledFields.Add("HDModel");
                  
                    if (enableHDSize == "true")
                        enabledFields.Add("HDsize");
                    
                    if (enablePowerSupplyLeft == "true")
                        enabledFields.Add("PowerSupplyLeft");
                    
                    if (enablePowerSupplyRight == "true")
                        enabledFields.Add("PowerSupplyRight");
                    
                    if (enableTestSet == "true")
                        enabledFields.Add("TestSet");
                    
                    enabledFieldsLog = String.Join(", ", enabledFields);                   
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Error reading the file: {ex.Message}");
                Console.WriteLine("Press any key to continue the aplication.");
                Console.ReadKey();
            }

            string inputFilePath = @"C:\ProgramData\National Instruments\NIvisa\visaconf.ini";

            List<string> addressTS = new List<string>();

            try
            {

                if (File.Exists(inputFilePath) && enableLog == "true")
                {
                    string[] lines = File.ReadAllLines(inputFilePath);
                    Dictionary<string, string> foundAliases = new Dictionary<string, string>();

                    foreach (string line in lines)
                    {
                        foreach (string alias in aliasPowerSupplyLeftSide)
                        {
                            if (line.Contains(alias))
                            {
                                string[] parts = line.Split('=');
                                if (parts.Length == 2)
                                {
                                    foundAliases[alias] = parts[1].Trim().Replace("\"", "");
                                }
                            }
                        }
                    }
                    foreach (string line in lines)
                    {
                        foreach (string alias in aliasPowerSupplyRightSide)
                        {
                            if (line.Contains(alias))
                            {
                                string[] parts = line.Split('=');
                                if (parts.Length == 2)
                                {
                                    foundAliases[alias] = parts[1].Trim().Replace("\"", "");
                                }
                            }
                        }

                        foreach (string alias in aliasTestset)
                        {
                            if (line.Contains(alias))
                            {
                                string[] parts = line.Split('=');
                                if (parts.Length == 2)
                                {
                                    foundAliases[alias] = parts[1].Trim().Replace("\"", "");
                                }
                            }
                        }
                    }
                    foreach (var alias in foundAliases)
                    {                        
                        string[] parts = alias.Value.Split(new string[] { "','" }, StringSplitOptions.None);
                        string parts1 = parts[1].Replace("\'", "");
                        if (alias.Key.Contains(String.Join(",", aliasTestset)))
                        {
                            if (parts.Length > 1)
                            {
                                string tcpIpAddress = parts1;
                                addressTS.Add(tcpIpAddress);
                            }
                            else
                            {
                                Console.WriteLine("Formato de endereço inesperado.");
                                Console.WriteLine();
                            }
                        }

                        if (alias.Key.Contains(String.Join(",", aliasPowerSupplyLeftSide)))

                            if (parts.Length > 1)
                            {
                                string tcpIpAddress = parts1;
                                addressPowerSupplyLeftSide.Add(tcpIpAddress);
                            }
                        if (alias.Key.Contains(String.Join(",", aliasPowerSupplyRightSide)))

                            if (parts.Length > 1)
                            {
                                string tcpIpAddress = parts1;
                                addressPowerSupplyRightSide.Add(tcpIpAddress);
                            }
                    }
                 
                }
               
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Press any key to continue the aplication.");
                Console.ReadKey();
            }
            // Cria o diretório se não existir
            if (!Directory.Exists(directoryLog) && enableLog == "true")
            {
                Directory.CreateDirectory(directoryLog);
            }

            // Verifica se o arquivo existe, se não, cria e adiciona cabeçalhos
            if (!File.Exists(filePathLog) && enableLog == "true")
            {
                using (StreamWriter sw = File.CreateText(filePathLog))
                {
                    sw.WriteLine(enabledFieldsLog);
                }
            }
        
            // Variáveis para armazenar os dados
            string manufacturer = "";
            string model = "";            
            string version = "";
            string build = "";
            double totalVisibleMemory = 0;                                  
            string hdModel = "";
            double hdSize = 0;
            string powerSupplyLeft = "";
            string powerSupplyRight = "";
            string testSet = "";

            if (enableLog == "true")
            {
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        if (enableManufacturer == "true")
                        {
                            manufacturer = obj["Manufacturer"].ToString();
                            messageLogList.Add(manufacturer);
                        }
                        if (enableSystemModel == "true")
                        {
                            model = obj["Model"].ToString();
                            messageLogList.Add(model);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message}");
                    Console.WriteLine("Press any key to continue the aplication.");
                    Console.ReadKey();
                }

                
                if (enableSystemName == "true")
                {
                    try
                    {
                        messageLogList.Add(Environment.MachineName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{ex.Message}");
                        Console.WriteLine("Press any key to continue the aplication.");
                        Console.ReadKey();
                    }
                    
                }
                try
                {
                    ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
                    ManagementObjectCollection results = searcher.Get();
                    ManagementObjectSearcher searcher1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                    foreach (ManagementObject os in searcher1.Get())
                    {
                        if (enableWindowsVersion == "true")
                        {
                            version = os["Caption"].ToString();
                            messageLogList.Add(version);
                        }
                        if (enableBuild == "true")
                        {
                            build = os["BuildNumber"].ToString();
                            messageLogList.Add(build);
                        }


                        if (enableInstalledRAM == "true")
                        {
                            foreach (ManagementObject result in results)
                            {
                                totalVisibleMemory = ((ulong)os["TotalVisibleMemorySize"]);
                                totalVisibleMemory = Math.Round((totalVisibleMemory / 1048576), 0);
                            }
                            messageLogList.Add(totalVisibleMemory.ToString() + " GB");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message}");
                    Console.WriteLine("Press any key to continue the aplication.");
                    Console.ReadKey();
                }
                try
                {
                    ManagementObjectSearcher searcher2 = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
                    List<string> rams = new List<string>();
                    foreach (ManagementObject queryObj in searcher2.Get())
                    {                        
                        string ramModel = queryObj["PartNumber"].ToString();
                        rams.Add(ramModel);                                                                                       
                    }
                    messageLogList.Add(string.Join(" / ", rams));
                }
                catch (ManagementException ex)
                {
                    Console.WriteLine($"{ex.Message}");
                    Console.WriteLine("Press any key to continue the aplication.");
                    Console.ReadKey();
                }

                try
                {
                    ManagementObjectSearcher searcher3 = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                    List<string> hds = new List<string>();
                    foreach (ManagementObject queryObj in searcher3.Get())
                    {
                        if (enableHDModel == "true")
                        {
                            hdModel = queryObj["Model"].ToString();
                            hds.Add(hdModel);                            
                        }
                        messageLogList.Add(string.Join(" / ", hds));
                        if (enableHDSize == "true")
                        {
                            hdSize = (ulong)queryObj["Size"];
                            hdSize = Math.Round((hdSize / 1000000000), 0);
                            messageLogList.Add(hdSize.ToString() + " GB");
                        }
                        
                    }
                }
                catch (ManagementException ex)
                {
                    Console.WriteLine($"{ex.Message}");
                    Console.WriteLine("Press any key to continue the aplication.");
                    Console.ReadKey();
                }

                if (enablePowerSupplyLeft == "true")
                {
                    // The VNA uses a message based session
                    MessageBasedSession[] mbSessions = new MessageBasedSession[addressPowerSupplyLeftSide.Count];
                    Session[] mySessions = new Session[addressPowerSupplyLeftSide.Count];

                    // Response strings
                    string[] responseStrings = new string[addressPowerSupplyLeftSide.Count];

                    for (int i = 0; i < addressPowerSupplyLeftSide.Count; i++)
                    {
                        try
                        {
                            // Open a Session to the VNA
                            mySessions[i] = ResourceManager.GetLocalManager().Open(addressPowerSupplyLeftSide[i]);
                            // Cast this to a message based session
                            mbSessions[i] = (MessageBasedSession)mySessions[i];
                            // Send "*IDN?" command
                            mbSessions[i].Write("*IDN?\n");
                            // Read the response
                            responseStrings[i] = mbSessions[i].ReadString();
                            // Return to Local Control
                            powerSupplyLeft = responseStrings[i];
                            powerSupplyLeft = powerSupplyLeft.Replace(',', '-');
                            mbSessions[i].Write("RTL\n");
                        }
                        catch (VisaException v_exp)
                        {
                            Console.WriteLine($"error find: {addressPowerSupplyLeftSide[i]}");
                            Console.WriteLine($"error VISA: {v_exp.Message}");
                            Console.WriteLine("Press any key to continue the aplication.");
                            Console.ReadKey();
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine($"error find: {addressPowerSupplyLeftSide[i]}");
                            Console.WriteLine($"error: {exp.Message}");
                            Console.WriteLine("Press any key to continue the aplication.");
                            Console.ReadKey();
                        }
                        finally
                        {
                            // Close the session if it was opened
                            mbSessions[i]?.Dispose();
                        }
                    }
                    messageLogList.Add(powerSupplyLeft);
                }
                if (enablePowerSupplyRight == "true")
                {
                    string[] responseStrings = new string[addressPowerSupplyRightSide.Count];
                    MessageBasedSession[] mbSessions = new MessageBasedSession[addressPowerSupplyRightSide.Count];
                    Session[] mySessions = new Session[addressPowerSupplyRightSide.Count];
                    for (int i = 0; i < addressPowerSupplyRightSide.Count; i++)
                    {
                        try
                        {
                            // Open a Session to the VNA
                            mySessions[i] = ResourceManager.GetLocalManager().Open(addressPowerSupplyRightSide[i]);
                            // Cast this to a message based session
                            mbSessions[i] = (MessageBasedSession)mySessions[i];
                            // Send "*IDN?" command
                            mbSessions[i].Write("*IDN?\n");
                            // Read the response
                            responseStrings[i] = mbSessions[i].ReadString();
                            // Return to Local Control
                            powerSupplyRight = responseStrings[i];
                            powerSupplyRight = powerSupplyRight.Replace(',', '-');
                            mbSessions[i].Write("RTL\n");
                        }
                        catch (VisaException v_exp)
                        {
                            Console.WriteLine($"error find: {addressPowerSupplyRightSide[i]}");
                            Console.WriteLine($"error VISA: {v_exp.Message}");
                            Console.WriteLine("Press any key to  the continue aplication.");
                            Console.ReadKey();
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine($"error find: {addressPowerSupplyRightSide[i]}");
                            Console.WriteLine($"error: {exp.Message}");
                            Console.WriteLine("Press any key to continue the aplication.");
                            Console.ReadKey();
                        }
                        finally
                        {
                            // Close the session if it was opened
                            mbSessions[i]?.Dispose();
                        }
                    }
                    messageLogList.Add(powerSupplyRight);
                }
                if (enableTestSet == "true")
                {
                    // The VNA uses a message based session
                    MessageBasedSession[] mbSessionsTestSet = new MessageBasedSession[addressTS.Count];
                    Session[] mySessionsTestSet = new Session[addressTS.Count];

                    // Response strings test set
                    string[] responseStringsTestSet = new string[addressTS.Count];

                    for (int j = 0; j < addressTS.Count; j++)
                    {
                        try
                        {
                            // Open a Session to the VNA
                            mySessionsTestSet[j] = ResourceManager.GetLocalManager().Open(addressTS[j]);

                            // Cast this to a message based session
                            mbSessionsTestSet[j] = (MessageBasedSession)mySessionsTestSet[j];

                            // Send "*IDN?" command
                            mbSessionsTestSet[j].Write("*IDN?\n");

                            // Read the response
                            responseStringsTestSet[j] = mbSessionsTestSet[j].ReadString();
                            testSet = responseStringsTestSet[j];
                            testSet = testSet.Replace(",", "-");

                            messageLogList.Add(testSet);
                            Console.WriteLine($"test set: {testSet}");
                            // Return to Local Control

                            testSet = testSet.Replace(",", "-");

                            mbSessionsTestSet[j].Write("RTL\n");
                        }
                        catch (VisaException v_exp)
                        {
                            Console.WriteLine($"Visa caught an error with {addressTS[j]}!!");
                            Console.WriteLine(v_exp.Message);
                            Console.WriteLine();

                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine($"Something didn't work with {addressTS[j]}!!");
                            Console.WriteLine(exp.Message);
                            Console.WriteLine();
                        }
                        finally
                        {                         
                            mbSessionsTestSet[j]?.Dispose();
                        }
                    }
                    
                }
            }

            if (enableLog == "true")
            {
                string msgLog = String.Join(", ", messageLogList);
                Log(msgLog, filePathLog);
            }
        }
        static void Log( string message, string filePath)
        {
            using (StreamWriter sw = File.AppendText(filePath))
            {
                string logEntry = $"{DateTime.Now},{message}";
                sw.WriteLine(logEntry);
            }
        }
    }
}