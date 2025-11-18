using Microsoft.VisualStudio.TestPlatform.TestHost;
using System;
namespace xlsxToJSON.UnitTests;


public class Tests
{
    private string _testDirectory;

    [SetUp]
    public void Setup()
    {
        // Vytvoření dočasného adresáře pro testování
        _testDirectory = Path.Combine(Path.GetTempPath(), "xlsxToJSONTests");
        Directory.CreateDirectory(_testDirectory);
    }

    [TearDown]
    public void Teardown()
    {
        // Vyčištění dočasných testovacích souborů
        if (Directory.Exists(_testDirectory))
        {
            Directory.Delete(_testDirectory, true);
        }
    }

    
    //test pro naplnění fronty
    [Test]
    public void FileQueueShouldFillCorrectlyWhenValidInput()
    {
        string input = "soubor1.xlsx,soubor2.xlsx,soubor3.xlsx";
        var fileQueue = new Queue<string>();
        
        string[] fileNames = input.Split(',', System.StringSplitOptions.RemoveEmptyEntries);
        foreach (var file in fileNames)
        {
            fileQueue.Enqueue(file);
        }
        
        Assert.AreEqual(3, fileQueue.Count);
        Assert.AreEqual("soubor1.xlsx", fileQueue.Dequeue());
        Assert.AreEqual("soubor2.xlsx", fileQueue.Dequeue());
        Assert.AreEqual("soubor3.xlsx", fileQueue.Dequeue());
    }

   
    [Test]
    public void EnsureXlsxExtensionShouldAddExtensionWhenMissing()
    {
        
        string inputFileName = "soubor_bez_pripony";
        string expectedFileName = "soubor_bez_pripony.xlsx";
        
        string resultFileName = Tests.Add(inputFileName);
        
        Assert.AreEqual(expectedFileName, resultFileName);
    }

    [Test]
    public void EnsureXlsxExtensionShouldNotChangeFileNameWhenExtensionExists()
    {
        // Arrange
        string inputFileName = "soubor_s_priponou.xlsx";
        string expectedFileName = "soubor_s_priponou.xlsx";

        // Act
        
        var x = new ExcelProcessor();
        
        x.AddXlsx(inputFileName);
        string resultFileName = EnsureXlsxExtension(inputFileName);

        // Assert
        Assert.AreEqual(expectedFileName, resultFileName);
    }
}
