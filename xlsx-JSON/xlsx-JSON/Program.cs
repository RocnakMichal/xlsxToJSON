using Newtonsoft.Json;
using OfficeOpenXml;
using Path = System.IO.Path;


namespace xlsx_JSON;

public abstract class Program
{
    private static bool _shortJson;
    public static void Main(string[] args)
    {
        //fronta pro více vstupů
        var fileQueue = new Queue<string>();

        // tabulky jsou ve stejném souboru jako program
        var baseDir = AppContext.BaseDirectory;
        if (baseDir == null) throw new InvalidOperationException("Složka neexistuje");
        
        var projectDir = GetProjectRoot(baseDir);
        
        while (true)
        {
            // vstup z konzole
            Console.WriteLine("Zadejte název souboru, pokud chceš převést více souborů, odděl je čárkou  ");
            var xlsxFile = Console.ReadLine();
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
                                                    "Pro kratší JSON soubor napiš kratsi zapis" +
                                                    "\n" +
                                                    "Pro delší JSON soubor napiš delsi zapis" +
                                                    "\n" +
                                                    "Pro ukončení napiš klíčové slovo KONEC");
                continue;
            }

           
            if (string.Equals(xlsxFile, "kratky vypis", StringComparison.OrdinalIgnoreCase))
            {
                _shortJson = true;
                WriteColorLine(ConsoleColor.Yellow,"Výpis je nyní krátký");
                continue;
            }
            if (string.Equals(xlsxFile, "dlouhy vypis", StringComparison.OrdinalIgnoreCase))
            {
                _shortJson = false;
                WriteColorLine(ConsoleColor.Yellow,"Výpis je nyní dlouhý");
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
                fileQueue = SplitInput(xlsxFile);
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
 /** <summary>
  Nejdříve se oddělí vstupy v "", aby mohl uživatel zpracovat soubory, které obsahují čárku. Poté se zpracují ostatní soubory, jako delimiter se používá čárka
  </summary>
  <param name="input">Vstup od uživatele z konzole, názvy souborů jsou oddělené</param>
  <returns>Fronta všech souborů pro zpracování</returns> */
    public static Queue<string> SplitInput(string input)
    {
        var result = new Queue<string>();
        // nejdříve rozdělíme podle "" a sekundárně poté podle ,
        var parts = input.Split('"', StringSplitOptions.TrimEntries);
        for (var i = 0; i < parts.Length; i++)
            // vstup v uvozovkách je oddělen sudým počtem uvozovek
            if (i % 2 == 1)
            {
                if (!string.IsNullOrWhiteSpace(parts[i]))
                    result.Enqueue(parts[i]);
            }
            else
            {
                // sekundární oddělování
                var subparts = parts[i]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim());

                foreach (var item in subparts)
                    result.Enqueue(item);
            }

        return result;
    }
    
    /** <summary>
     Funkce pro přidání .xlsx přípony k souboru, pokud tak uživatel nedělá.
     Pouze pro pohodlnější prácí s programem
     Pokud uživatel zadá název souboru s příponou nic se nepřidá
    </summary>*/
    /// <param name="fileName">Vstupní soubor, který uživatel zadá do konzole</param>
    /// <returns>Vstupní soubor, který postrádal příponu .xlsx ji nyní má.</returns>
    public static string AddXlsx(string fileName)
    {
        if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) return fileName + ".xlsx";

        return fileName;
    }

    /** <summary> Pro použitelnost se vynořujeme z adresáře, kde jsem vytvoří program až do adresáře, kde budeme dále pracovat
     </summary>
     <param name="currentDir">Zde se vytvoří program</param>
     <returns>Adresář, kde se nachází spustitelný kód a soubory na převod</returns>
     <exception cref="InvalidOperationException">Nelze najít adresář</exception>*/
    public static string GetProjectRoot(string currentDir)
    {
        // Iterujeme dokud se nedostaneme na požadovanou složku /xlsx-JSON je složka, kde se nachází program
        var dir = new DirectoryInfo(currentDir);
        while (dir != null)
        {
            if (Directory.Exists(Path.Combine(dir.FullName, "xlsx-JSON")))
                // Pracovní složka
                return Path.Combine(dir.FullName, "xlsx-JSON");
            // Posun o jednu úroveň nahoru
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Nelze najít kořenový adresář projektu.");
    }
    
    /** <summary>
     Vypíše text barevně
     </summary>
     <param name="color">Barva textu</param>
     <param name="text">Vlastní text</param>*/
    public static void WriteColorLine(ConsoleColor color, string text)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(text);
        Console.ResetColor();
    }

    /** <summary>
    Funkce na převod .xlsx souborů do formátu .JSON
    Postupuje po jednotilivých buńkách
    První řádek definuje názvy v kolekci
     </summary>
     <param name="fileName">Název souboru, který chceme převést</param>*/
    public static void ProcessExcelFile(string fileName)
    {
        try
        {
            using var package = new ExcelPackage(new FileInfo(fileName));
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

                        // pokud chce u6ivatel pouze krátký výpis neuloži se hodnoty buněk, kde je hodnota null, hodnota 0 se uloží
                        if (value != "" &&  _shortJson )
                        {
                            // Přidání do seznamu jako klíč-hodnota slovník
                            dataCollection.Add(new Dictionary<string, string>
                            {
                                // Název sloupce
                                { "Období", columnName },
                                // Hodnota buňky
                                { "Počet kusů", value }
                            });
                        }

                        else if(! _shortJson || value != "")
                        {
                            // Přidání do seznamu jako klíč-hodnota slovník
                            dataCollection.Add(new Dictionary<string, string>
                            {
                                // Název sloupce
                                { "Období", columnName },
                                // Hodnota buňky
                                { "Počet kusů", value }
                            });
                        }
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

        catch (Exception ex)
        {
            WriteColorLine(ConsoleColor.Red, $"Došlo k chybě: {ex.Message}");
        }
    }
}