using Newtonsoft.Json;
using OfficeOpenXml;




string xlsxFile = @"C:\Users\Michal Ročňák\Desktop\Příklad - zakázky do JSON.xlsx";


while (true)
{

// vstup z konzole
    Console.WriteLine("Zadejte název souboru ");
    xlsxFile = Console.ReadLine();
    if (string.Equals(xlsxFile, "konec", StringComparison.OrdinalIgnoreCase) || xlsxFile == null)
    {
        Console.WriteLine("Program ukončen.");
        break;
    }


// tabulky jsou ve stejném souboru jako program
    string baseDir = AppContext.BaseDirectory;
    string projectDir = Directory.GetParent(baseDir).Parent.Parent.Parent.FullName;
    string filePath = Path.Combine(projectDir, xlsxFile);


//Console.WriteLine($"Soubor: {filePath}");

// soubor existuje
    if (File.Exists(filePath))
    {
        ProcessExcelFile(filePath);

    }
// uzivatel zadal nazev souboru bez přípony
    else
    {
        if (!xlsxFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            xlsxFile += ".xlsx";
            filePath = Path.Combine(projectDir, xlsxFile);
            Console.WriteLine($"Soubor: {filePath}");
        }

        if (File.Exists(filePath))
        {
            ProcessExcelFile(filePath);
        }
        else
        {
            Console.WriteLine("Soubor neexistuje!");
        }

    }

}

static void ProcessExcelFile(string fileName)
{
    try
    {
        using (var package = new ExcelPackage(new FileInfo(fileName)))
        {
            var workbook = package.Workbook;
            // Použití prvního listu
            var worksheet = workbook.Worksheets[0];

            int rows = worksheet.Dimension.Rows;
            // int columns = worksheet.Dimension.Columns;

            // Seznam pro uložení všech řádků, jako JSON objekty
            var resultList = new List<Dictionary<string, object>>();

            // Načítání dat z Excelu, 1. řádek-hlavičky 
            for (int row = 2; row <= rows; row++)
            {
                var rowData = new Dictionary<string, object>();

                // Zpracování sloupců A, B, C (hlavičky jako klíče)
                for (int col = 1; col <= 3; col++)
                {
                    // Název sloupce z prvního řádku
                    string columnName = worksheet.Cells[1, col].Text;
                    // Načtení dat
                    rowData[columnName] = worksheet.Cells[row, col].Text;
                }


                var dataCollection = new List<Dictionary<string, string>>();

                // Zpracování sloupců D (4) až AA (27)
                for (int col = 4; col <= 27; col++)
                {
                    // Název sloupce D-AA
                    string columnName = worksheet.Cells[1, col].Text;
                    // Hodnota buňky
                    string value = worksheet.Cells[row, col].Text;

                    // Přidání do seznamu jako klíč-hodnota slovník
                    dataCollection.Add(new Dictionary<string, string>
                    {
                        { "Období", columnName }, // Název sloupce
                        { "Počet kusů", value } // Hodnota buňky
                    });
                }


                // Uložení kolekce do objektu
                rowData["dataCollection"] = dataCollection;

                // Přidání zpracovaného řádku do výsledného seznamu
                resultList.Add(rowData);
            }

            // převod do JSON
            string jsonResult = JsonConvert.SerializeObject(resultList, Formatting.Indented);

            // Uložení souboru
            string jsonPath = Path.ChangeExtension(fileName, ".json");
            File.WriteAllText(jsonPath, jsonResult);

            Console.WriteLine($"JSON uložen zde: {jsonPath}");
        }
    }

    catch (Exception ex)
    {
        Console.WriteLine($"Došlo k chybě: {ex.Message}");
    }
}


