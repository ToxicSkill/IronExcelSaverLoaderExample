using IronXL;

namespace DataToExcel;

public class ExcelLoader
{
    public (List<string> headers, List<List<string>> data) LoadExcel(string path)
    {
        WorkBook workBook = WorkBook.Load(path);
        
        var data = new List<List<string>>();
        var headers = new List<string>();
        var workSheet = workBook.WorkSheets[0];

        foreach (var item in workSheet.Columns)
        {
            headers.Add(item.Rows[0].ToString());
        }
        foreach (var item in workSheet.Rows.Skip(1))
        {
            var internalList = new List<string>();  
            foreach (var colItem in item.Columns)
            {
                internalList.Add(colItem.ToString());
            }
            data.Add(internalList);
        }
        return new(headers, data);
    }
}