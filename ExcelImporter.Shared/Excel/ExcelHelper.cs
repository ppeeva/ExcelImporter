using OfficeOpenXml;

namespace ExcelImporter.Shared.Excel
{
    public static class ExcelHelper
    {
        public static T GetModelFromRow<T>(ExcelWorksheet workSheet, int startColumn, int endColumn, int dataRow, int rowWithProperty, IEnumerable<ExcelDDL> ddlLists)
        {
            bool isDataRow = false;
            T obj = (T)Activator.CreateInstance(typeof(T), null);
            for (int col = startColumn; col <= endColumn; col++)
            {
                object rawValue = workSheet.Cells[dataRow, col].Value;

                // do not change it any more once we find a non-null value 
                isDataRow = !isDataRow ? (rawValue != null && !String.IsNullOrWhiteSpace(rawValue.ToString())) : isDataRow;

                // first row contains entity properties names
                string propName = workSheet.Cells[rowWithProperty, col].Value != null ? workSheet.Cells[rowWithProperty, col].Value.ToString() : String.Empty;
                ExcelDropDownListType listType = ExportModelHelper.GetPropFKListType<T>(propName);
                string listStringProp = ExportModelHelper.GetPropFKStringProp<T>(propName);

                if (listType != ExcelDropDownListType.None && rawValue != null)
                {
                    ExcelDDL ddl = ddlLists.FirstOrDefault(x => x.ListType == listType);
                    if (ddl != null)
                    {
                        ExcelDDLItem ddlItem = ddl.Collection.Where(x => x.DisplayName.ToLower() == rawValue.ToString().ToLower()).FirstOrDefault();
                        if (ddlItem != null)
                        {
                            rawValue = ddlItem.Value;
                        }

                        if (!String.IsNullOrEmpty(listStringProp))
                        {
                            ExportModelHelper.SetPropertyValue<T>(listStringProp, rawValue, obj);
                        }
                    }
                }

                ExportModelHelper.SetPropertyValue<T>(propName, rawValue, obj);
            }

            return isDataRow ? obj : default;
        }

    }
}
