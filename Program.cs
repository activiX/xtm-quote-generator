using ClosedXML.Excel;
using System.Text.RegularExpressions;
using IOPath = System.IO.Path;

namespace XtmQuoteV1;

public static class Program
{
    
    public static int Main(string[] args)
    {
        Console.Clear();
        Console.WriteLine("=== XTM Quote Generator ===");
        Console.WriteLine();
        Console.WriteLine("Enter the path to folder with Excel analysis files:");
        

        string? inputFolder;
        while(true)
        {
            Console.Write("> ");
            inputFolder = (Console.ReadLine() ?? "").Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(inputFolder) || !Directory.Exists(inputFolder))
            {
                Console.WriteLine("Entered path doesn't exist.");
                continue;
            }
            break;
        }
        var excelFiles = Directory.GetFiles(inputFolder, "*.xlsx")
            .Where(f => !IOPath.GetFileName(f).StartsWith("~$"))
            .ToList();

        if (excelFiles.Count == 0)
        {
            Console.WriteLine("There are no .xlsx files on provided path.");
            Exit();
            return 1;
        }

        Console.WriteLine($"Number of Excel files found: {excelFiles.Count}");
        foreach (var f in excelFiles)
        {
            Console.WriteLine($" - {IOPath.GetFileName(f)}");
        }

        string outputFolder;

        Console.WriteLine("Enter the path where output quote should be placed:");
        while (true)
        {
            Console.Write("> ");
            outputFolder = (Console.ReadLine() ?? "").Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                Console.WriteLine("Entered path doesn't exist.");
                continue;
            }

            break;
        }
        var mapPath = IOPath.Combine(AppContext.BaseDirectory, "language-map.csv");
        if (!File.Exists(mapPath))
        {
            Console.WriteLine($"There is no mapping file \"language-map.csv\", in folder that contains exe.");
            Exit();
            return 1;
        }

        var languageMap = LoadLanguageMap(mapPath);
        Console.WriteLine($"Mapping file loaded");

        int outputRow = 2;
        int processed = 0;
        int skipped = 0;

        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string outPath = IOPath.Combine(outputFolder, $"Quote_{timestamp}.xlsx");

        using var outWb = new XLWorkbook();
        var outWs = outWb.Worksheets.Add("Quote");
        
        outWs.Cell(1, 1).Value = "Target language";
        outWs.Cell(1, 2).Value = "Context Matches";
        outWs.Cell(1, 3).Value = "Repetitions";
        outWs.Cell(1, 4).Value = "100% match";
        outWs.Cell(1, 5).Value = "95-99%";
        outWs.Cell(1, 6).Value = "75-94%";
        outWs.Cell(1, 7).Value = "New Words";
        outWs.Cell(1, 8).Value = "Total words";
        outWs.Row(1).Style.Font.Bold = true;

        foreach (var excel in excelFiles)
        {
            Console.WriteLine($"Processing file: {IOPath.GetFileName(excel)}");
            try
            {
            using var wb = new XLWorkbook(excel);
            var ws = wb.Worksheets.First(); 
            
            string b4 = ws.Cell("B4").GetString();

            if (!b4.Contains("Target", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Skipping file {IOPath.GetFileName(excel)} - invalid template/layout");
                skipped++;
                continue;
            }

            string langCode = ExtractTargetLanguageCode(b4); 

            //Fixed cell positions, since XTM template is always same
            long d9  = ReadWordCount(ws, "D9");
            long d11 = ReadWordCount(ws, "D11");
            long d12 = ReadWordCount(ws, "D12");
            long d13 = ReadWordCount(ws, "D13");
            long d14 = ReadWordCount(ws, "D14");
            long d15 = ReadWordCount(ws, "D15");
            long d16 = ReadWordCount(ws, "D16");
            long d18 = ReadWordCount(ws, "D18");
            long d19 = ReadWordCount(ws, "D19");
            long d20 = ReadWordCount(ws, "D20");
            long d21 = ReadWordCount(ws, "D21");
            long d22 = ReadWordCount(ws, "D22");
            
            long contextMatches = d9 + d11;                       // B
            long repetitions    = d18;                            // C
            long match100       = d12;                            // D
            long match9599      = d13 + d19;                      // E
            long match7594      = d14 + d15 + d20 + d21;          // F
            long newWords       = d16 + d22;                      // G
            long totalWords     = contextMatches + repetitions + match100 + match9599 + match7594 + newWords;             
            
            outWs.Cell(outputRow, 1).Value = MapLanguage(langCode, languageMap);        
            outWs.Cell(outputRow, 2).Value = contextMatches;
            outWs.Cell(outputRow, 3).Value = repetitions;
            outWs.Cell(outputRow, 4).Value = match100;
            outWs.Cell(outputRow, 5).Value = match9599;
            outWs.Cell(outputRow, 6).Value = match7594;
            outWs.Cell(outputRow, 7).Value = newWords;
            outWs.Cell(outputRow, 8).Value = totalWords;

            outputRow++; 
            processed++;
            }
            catch (Exception)
            {
                Console.WriteLine($"Skipping file {IOPath.GetFileName(excel)} - file is damaged, open or not valid .xlsx");
                skipped++;
                continue;
            }
        }
        outWs.Columns().AdjustToContents();
        var dataRange = outWs.Range(2, 1, outputRow - 1, 8);
        dataRange.Sort(1);
        outWb.SaveAs(outPath);
        
        Console.WriteLine($"\nOutput: {outPath}");
        Console.WriteLine($"\nProcessed: {processed}, skipped: {skipped}");
        Exit();
        return 0;
    }

//Method to ensure that values for calaculation are in number format
    private static long ReadWordCount(IXLWorksheet ws, string address)
    {
        var cell = ws.Cell(address);

        if (cell.TryGetValue<double>(out var d))
            return (long)Math.Round(d);

        var s = cell.GetString().Trim();
        if (string.IsNullOrWhiteSpace(s))
            return 0;

        s = s.Replace(" ", "").Replace(",", "").Replace("\u00A0", "");
        return long.TryParse(s, out var v) ? v : 0;
    }

//Method containing Regex to extract language code from .xlxs file
    private static string ExtractTargetLanguageCode(string b4)
    {
        if (string.IsNullOrWhiteSpace(b4)) return "";

        var m = Regex.Match(b4, @"Target\s*Language\s*:\s*([A-Za-z]{2,3}(?:[_-][A-Za-z0-9]+)*)");
        if (m.Success) return m.Groups[1].Value.Replace("-", "_");
        return "";
    }

//Method to compare values from mapping file to codes from .xlxs and assign correct values
    static string MapLanguage(string langCode, Dictionary<string, string> languageMap)
    {
        if (string.IsNullOrWhiteSpace(langCode))
            return "";

        if (languageMap.TryGetValue(langCode, out var mapped))
            return mapped;

        Console.WriteLine($"Language code not supported: {langCode}");
        return langCode;
    }
//Method to load mapping CSV file
    static Dictionary<string, string> LoadLanguageMap(string csvPath)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in File.ReadLines(csvPath))
        {
            var s = line.Trim();
            if (s.Length == 0 || s.StartsWith("#")) continue;

            var parts = s.Split(',', 2);
            if (parts.Length < 2) continue;

            var key = parts[0].Trim();
            var val = parts[1].Trim();

            if (key.Equals("langCode", StringComparison.OrdinalIgnoreCase)) continue; // header

            if (key.Length == 0 || val.Length == 0) continue;

            map[key] = val; // last wins
        }

        return map;
    }
    
    //Method to hold exiting program in case there is unexpected result
    static void Exit(string message = "Press Enter to exit app.")
    {
        Console.WriteLine(message);
        Console.ReadLine();
    }
}
