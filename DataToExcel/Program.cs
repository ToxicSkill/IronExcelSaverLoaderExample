using DataToExcel;

(List<string> headers, List<List<string>> data) data = new ExcelLoader().LoadExcel("test.xlsx");
var dataToSave = new List<List<string>>();
dataToSave.Add(data.headers);
foreach (var item in data.data)
{
    dataToSave.Add(item);
}
var settings = new ExcelSettings();
settings.Colors = new()
            {
                Color.Aquamarine,
                Color.Blue,
                Color.Orange,
                Color.Orchid,
                Color.Red,
                Color.Green,
                Color.DarkBlue,
                Color.Gold,
                Color.Magenta
            };
settings.HasHeader = true;
settings.Rotation = 90;
settings.SheetName = "New";
settings.Format = IronXL.ExcelFileFormat.XLSX; 
new ExcelSaver().SaveExcel(dataToSave, 
    "saveTest.xlsx",
    settings);