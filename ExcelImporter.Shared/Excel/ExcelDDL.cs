
namespace ExcelImporter.Shared.Excel
{
    public class ExcelDDL
    {
        public ExcelDropDownListType ListType { get; set; }
        public IList<ExcelDDLItem> Collection { get; set; }
    }

    public class ExcelDDLItem
    {
        public object Value { get; set; }
        public string DisplayName { get; set; }
        public object ParentValue { get; set; }
    }
}
