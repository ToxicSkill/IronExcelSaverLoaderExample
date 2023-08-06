using IronXL;

namespace DataToExcel
{
    public class ExcelSettings
    {
        public short Rotation { get; set; }

        public string SheetName { get; set; }

        public ExcelFileFormat Format { get; set; }

        public List<Color> Colors { get; set; }

        public bool HasHeader { get; set; }
    }
}
