using xlsx_JSON;
using OfficeOpenXml;
using Newtonsoft.Json;
namespace Test;

[TestClass]
public class UnitTest1
{
    [TestMethod]
    public void Add_Xlsx()
    {
        var inputFileName = "soubor_bez_pripony";
        var expectedFileName = "soubor_bez_pripony.xlsx";

        var result = Program.AddXlsx(inputFileName);

        Assert.AreEqual(expectedFileName, result);
    }

    [TestMethod]
    public void Dont_Add_Xlsx()
    {
        var inputFileName = "soubor_s_priponou.xlsx";
        var expectedFileName = "soubor_s_priponou.xlsx";

        var result = Program.AddXlsx(inputFileName);

        Assert.AreEqual(expectedFileName, result);
    }

    [TestMethod]
    public void EmptyFileName_AddXlsx()
    {
        var inputFileName = "";
        var expectedFileName = ".xlsx";

        var result = Program.AddXlsx(inputFileName);

        Assert.AreEqual(expectedFileName, result);
    }
    [TestMethod]
    public void ConvertExcelFile()
    {
       // Tvorba souboru
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        var excelPath = Path.Combine(tempDir, "test.xlsx");
        var jsonPath = Path.Combine(tempDir, "test.json");

      

        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");

            // Hlavičky
            ws.Cells[1, 1].Value = "Klient";
            ws.Cells[1, 2].Value = "IČ";
            ws.Cells[1, 3].Value = "Zakázka";
            
            // Období zjednodušeno na čísla
            for (int c = 4; c <= 27; c++)
            {
                ws.Cells[1, c].Value = c.ToString();
            }
            
            // První záznam
            ws.Cells[2, 1].Value = "ROSSMANN, spol. s r.o.";
            ws.Cells[2, 2].Value = "61246093";
            ws.Cells[2, 3].Value = "RO0001";
            
            for (int col = 4; col <= 27; col++)
            {
                ws.Cells[2, col].Value = col;
            }

            package.SaveAs(new FileInfo(excelPath));
        }
        
        Program.ProcessExcelFile(excelPath);
        
        var json = File.ReadAllText(jsonPath);

        var data = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);

        Assert.IsNotNull(data);

        var firstRow = data[0];

        Assert.AreEqual("ROSSMANN, spol. s r.o.", firstRow["Klient"].ToString());
        Assert.AreEqual("61246093",firstRow["IČ"].ToString());
        Assert.AreEqual("RO0001", firstRow["Zakázka"].ToString());
        
        var dataCollection = firstRow["dataCollection"] as Newtonsoft.Json.Linq.JArray;
        Assert.AreEqual(24, dataCollection.Count);
        
        var secondRow = dataCollection[0];
        Assert.AreEqual("4",secondRow["Období"].ToString());
        Assert.AreEqual("4", secondRow["Počet kusů"].ToString());
        
        
        var secondRowSecondCell = dataCollection[1];
        Assert.AreEqual("5",secondRowSecondCell["Období"].ToString());
        Assert.AreEqual("5", secondRowSecondCell["Počet kusů"].ToString());
    }
}