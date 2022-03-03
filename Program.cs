using System.Drawing;
using System.Reflection.Metadata;
using ExelDemo;using OfficeOpenXml;
using OfficeOpenXml.Style;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var file = new FileInfo(@"/home/payrav/Test.xlsx");
var user = GetSetupData();
await SaveExelFile(user,file);
List<Person> per = await LoadExcesFiles(file);
foreach (var p in user)
{
    Console.WriteLine($"{p.Id} {p.FirsName} {p.LAstName}");
}

static async Task<List<Person>> LoadExcesFiles(FileInfo info)
{
    List<Person> output = new();
    using var package = new ExcelPackage(info);
    await package.LoadAsync(info);

    var ws = package.Workbook.Worksheets[0];

    int row = 3;
    int col = 1;

    while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString())==false)
    {
        Person p = new();
        p.Id = int.Parse((ws.Cells[row, col].Value.ToString() ?? string.Empty));
        p.FirsName = ws.Cells[row, col + 1].Value.ToString();
        p.LAstName = ws.Cells[row, col + 2].Value.ToString();
        output.Add(p);
        row += 1;
    };
    return output;
}

static List<Person> GetSetupData()
{
    List<Person> output = new List<Person>()
     {
          new() {Id = 1, FirsName = "Test", LAstName = "testov"},
          new() {Id = 2, FirsName = "Test1", LAstName = "testov1"},
          new() {Id = 3, FirsName = "Test3", LAstName = "testov2"},
          new() {Id = 4, FirsName = "Test3", LAstName = "testov3"}
     };
     return output;
}

static async Task SaveExelFile(List<Person> persons, FileInfo fileInfo)
{
    DeleteIfExists(fileInfo);
    using var package = new ExcelPackage(fileInfo);
    var ws = package.Workbook.Worksheets.Add("MainReport");
    var range = ws.Cells["A1"].LoadFromCollection(persons, true);

    //Format the header
    ws.Cells["A1"].Value = "Our Cool Report";
    ws.Cells["A1:C1"].Merge = true;
    ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    ws.Row(1).Style.Font.Size = 24;
    ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

    ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    ws.Row(2).Style.Font.Bold = true;
    ws.Column(3).Width = 20;
    
    
    await package.SaveAsync();

}

static void DeleteIfExists(FileInfo file)
{
    if(file.Exists)
        file.Delete();
}