using ConsoleApp1;

Console.WriteLine("Hello, World!");
try
{
    var List = IExcelReport.SeedData();
    Stream temp = ExportExcel.CreateExcel(List);
    var fileStream = new FileStream("C:\\Users\\omid\\Desktop\\res.xlsx", FileMode.Create, FileAccess.Write);
    temp.CopyTo(fileStream);
    fileStream.Dispose();
}
catch (Exception e)
{
    Console.WriteLine(e);
    throw;
}

