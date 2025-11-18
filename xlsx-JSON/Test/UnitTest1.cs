namespace Test;

[TestClass]
public class UnitTest1
{
    [TestMethod]
    public void Add_Xlsx()
    {
        
        var inputFileName = "soubor_bez_pripony";
        var expectedFileName = "soubor_bez_pripony.xlsx";
        
        var result=Program.AddXlsx(inputFileName);
        
        Assert.AreEqual(expectedFileName, result);
    }
    
    
    
    [TestMethod]
    public void Dont_Add_Xlsx()
    {
        
        var inputFileName = "soubor_s_priponou.xlsx";
        var expectedFileName = "soubor_s_priponou.xlsx";
        
        var result=Program.AddXlsx(inputFileName);
        
        Assert.AreEqual(expectedFileName, result);
    }
    
    [TestMethod]
    public void EmptyFileName_AddXlsx()
    {
        string inputFileName = "";
        string expectedFileName = ".xlsx";
        
        string result = Program.AddXlsx(inputFileName);
        
        Assert.AreEqual(expectedFileName, result);
    }

    
}