using IronXL;

namespace DataToExcel
{
    public class ExcelSaver
    {
        public bool SaveExcel(List<List<string>> data, 
            string path,
            ExcelSettings settings)
        {
            var xlsWorkbook = WorkBook.Create(settings.Format); 
            var xlsSheet = xlsWorkbook.CreateWorkSheet(settings.SheetName);
            var rowCounter = 1;
            var alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            foreach (var item in data)
            {
                for (var i = 0; i < item.Count; i++)
                {
                    xlsSheet[$"{alpha[i]}{rowCounter}"].Value = item[i].ToString();
                }
                if (settings.HasHeader && rowCounter == 1)
                {
                    for (var i = 0; i < item.Count; i++)
                    {
                        xlsSheet[$"{alpha[i]}{rowCounter}"].Style.Rotation = settings.Rotation;
                        xlsSheet[$"{alpha[i]}{rowCounter}"].Style.SetBackgroundColor(settings.Colors[i]);
                    }
                }
                rowCounter++;
            }
            xlsWorkbook.SaveAs(path);
            return true;
        }
    }
}
