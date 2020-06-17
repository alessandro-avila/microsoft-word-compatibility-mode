using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Linq;

namespace WordCompatibilityMode
{
    public class Program
    {
        static void Main(string[] args)
        {
            string path = string.Empty;
            Console.Write("Set directory path (or press ENTER to set current dir):");
            path = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(path))
                path = @".";
            DirectoryInfo di = new DirectoryInfo(path: path);
            FileInfo[] files = di.GetFilesByExtensions(".docx", ".doc").ToArray();

            Application ap = new Application();
            {
                foreach (FileInfo f in files)
                {
                    try
                    {
                        Document wDoc = ap.Documents.Open(f.FullName);
                        string compatibilityMode = Enum.GetName(typeof(WdCompatibilityMode), wDoc.CompatibilityMode);
                        Console.WriteLine($"Document: {f.Name}: Compatibility Mode={compatibilityMode}");
                        wDoc.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Document: {f.Name}: {ex.Message}");
                    }
                }
            }
            Console.WriteLine("Completed.");
            Console.ReadLine();
        }
    }
}
