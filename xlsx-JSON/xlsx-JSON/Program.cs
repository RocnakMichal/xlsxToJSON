using Newtonsoft.Json;
using OfficeOpenXml;
using Path = System.IO.Path;


namespace xlsx_JSON;

public class Program
{
    public static void Main(string[] args)
    {
        string? xlsxFile;

        //fronta pro více vstupů
        var fileQueue = new Queue<string>();

        // tabulky jsou ve stejném souboru jako program
        var baseDir = AppContext.BaseDirectory;
        if (baseDir == null) throw new InvalidOperationException("Složka neexistuje");

        Console.WriteLine(baseDir);
        var projectDir = GetProjectRoot(baseDir);
        Console.WriteLine(projectDir);



        while (true)
        {
            // vstup z konzole
            Console.WriteLine("Zadejte název souboru, pokud chceš převést více souborů, odděl je čárkou  ");
            xlsxFile = Console.ReadLine();
            if (string.Equals(xlsxFile, "konec", StringComparison.OrdinalIgnoreCase) || xlsxFile == null)
            {
                Console.WriteLine("Program ukončen.");
                break;
            }

            if (string.Equals(xlsxFile, "napoveda", StringComparison.OrdinalIgnoreCase))
            {
                WriteColorLine(ConsoleColor.Yellow, "Zadej názvy souborů, ktere chces prevest do formatu .json." +
                                                    "\n" +
                                                    "Pro převod všech souborů ve složce napiš klíčové slovo VSE" +
                                                    "\n" +
                                                    "Pro ukončení napiš klíčové slovo KONEC");
                continue;
            }

            if (string.Equals(xlsxFile, "vse", StringComparison.OrdinalIgnoreCase))
            {
                //hledá vše, co končí .xlsx, nemusí se nic dolňovat, protože soubory jsou takto uloženy
                var allFiles = Directory.GetFiles(projectDir, "*.xlsx", SearchOption.TopDirectoryOnly);

                if (allFiles.Length == 0)
                {
                    WriteColorLine(ConsoleColor.Red, "Nebyly nalezeny žádné .xlsx soubory.");
                    continue;
                }

                foreach (var file in allFiles) fileQueue.Enqueue(Path.GetFileName(file));
            }
            else
            {
                // rozdělení vstupu, delimiter-"," 
                //string[] fileNames = xlsxFile.Split(',', StringSplitOptions.RemoveEmptyEntries);
              
                var fileNames = SplitInput(xlsxFile);
                
                // naplnění fronty
                foreach (var file in fileNames)
                    fileQueue.Enqueue(file);

                foreach (var a in fileNames) Console.WriteLine(a);
            }

            while (fileQueue.Count > 0)
            {
                xlsxFile = fileQueue.Dequeue();
                // tabulky jsou ve stejném souboru jako program

                var filePath = Path.Combine(projectDir, xlsxFile);


                // soubor existuje
                if (File.Exists(filePath))
                {
                    ProcessExcelFile(filePath);
                }
                // uzivatel zadal nazev souboru bez přípony
                else
                {
                    xlsxFile = AddXlsx(xlsxFile);
                    filePath = Path.Combine(projectDir, xlsxFile);

                    if (File.Exists(filePath))
                        ProcessExcelFile(filePath);
                    else
                        WriteColorLine(ConsoleColor.Red, $"Soubor {filePath} neexistuje!");
                }
            }
        }
    }
    
    //TODO nacteni do froty jiz pri zpracovani
    // Zpracovani vstupu na jednotlive nazvy souboru
    public static List<string> SplitInput(string input)
    {
        var result = new List<string>();
        // nejdříve rozdělíme podle "" a sekundárně poté podle ,
        var parts = input.Split('"', StringSplitOptions.TrimEntries);
        for (int i = 0; i < parts.Length; i++)
        {
            // vstup v uvozovkách je oddělen sudým počtem
            if (i % 2 == 1)
            {
                if (!string.IsNullOrWhiteSpace(parts[i]))
                    result.Add(parts[i]);
            }
            else
            {
                // sekundární oddělování
                var subparts = parts[i]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim());

                result.AddRange(subparts);
            }
        }

        return result;
    }



    //pridani pripony k souboru
    public static string AddXlsx(string fileName)
    {
        if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) return fileName + ".xlsx";

        return fileName;
    }
    public static string GetProjectRoot(string currentDir)
    {
        // Iterujeme dokud se nedostaneme na požadovanou složku /xlsx-JSON je složka, kde se nachází program
        var dir = new DirectoryInfo(currentDir);
        while (dir != null)
        {
            if (Directory.Exists(Path.Combine(dir.FullName, "xlsx-JSON")))
            {
                // Pracovní složka
                return Path.Combine(dir.FullName, "xlsx-JSON");
            }
            // Posun o jednu úroveň nahoru
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Nelze najít kořenový adresář projektu.");
    }

    public static void WriteColorLine(ConsoleColor color, string text)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(text);
        Console.ResetColor();
    }

    public static void ProcessExcelFile(string fileName)
    {
        try
        {
            using (var package = new ExcelPackage(new FileInfo(fileName)))
            {
                var workbook = package.Workbook;
                // ošetření prázdného sešitu
                if (workbook == null || workbook.Worksheets.Count == 0)
                {
                    WriteColorLine(ConsoleColor.Red, $"Soubor '{Path.GetFileName(fileName)}' je prázdný");
                    return;
                }

                // Použití prvního listu
                var worksheet = workbook.Worksheets[0];


                if (worksheet.Dimension == null)
                {
                    WriteColorLine(ConsoleColor.Red, $"Soubor '{Path.GetFileName(fileName)}' obsahuje prázdný list.");
                    return;
                }

                var rows = worksheet.Dimension.Rows;
                // int columns = worksheet.Dimension.Columns;

                // Seznam pro uložení všech řádků, jako JSON objekty
                var resultList = new List<Dictionary<string, object>>();

                // Načítání dat z Excelu, 1. řádek-hlavičky 
                for (var row = 2; row <= rows; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    // Zpracování sloupců A, B, C (hlavičky jako klíče)
                    for (var col = 1; col <= 3; col++)
                    {
                        // Název sloupce z prvního řádku
                        var columnName = worksheet.Cells[1, col].Text;
                        // Načtení dat
                        rowData[columnName] = worksheet.Cells[row, col].Text;
                    }


                    var dataCollection = new List<Dictionary<string, string>>();

                    // Zpracování sloupců D (4) až AA (27)
                    for (var col = 4; col <= 27; col++)
                    {
                        // Název sloupce D-AA
                        var columnName = worksheet.Cells[1, col].Text;
                        // Hodnota buňky
                        var value = worksheet.Cells[row, col].Text;

                        // Přidání do seznamu jako klíč-hodnota slovník
                        dataCollection.Add(new Dictionary<string, string>
                        {
                            // Název sloupce
                            { "Období", columnName },
                            // Hodnota buňky
                            { "Počet kusů", value } 
                        });
                    }

                    // Uložení kolekce do objektu
                    rowData["dataCollection"] = dataCollection;

                    // Přidání zpracovaného řádku do výsledného seznamu
                    resultList.Add(rowData);
                }

                // převod do JSON
                var jsonResult = JsonConvert.SerializeObject(resultList, Formatting.Indented);

                // Uložení souboru
                var jsonPath = Path.ChangeExtension(fileName, ".json");
                File.WriteAllText(jsonPath, jsonResult);
                WriteColorLine(ConsoleColor.Green, $"JSON je uložen zde: {jsonPath}");
            }
        }

        catch (Exception ex)
        {
            WriteColorLine(ConsoleColor.Red, $"Došlo k chybě: {ex.Message}");
        }
    }
}