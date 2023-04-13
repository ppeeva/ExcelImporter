
namespace ExcelImporter.Shared.Attributes
{
    [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = false)]
    public class ExportTemplateAttribute : Attribute
    {
        public string DisplayName { get; }
        public bool Ignore { get; }
        public ExcelDropDownListType FKSourceName { get; }
        public string FKStringProp { get; }
        public string ParentDDLProperty { get; }
        public bool AllowNonFKValues { get; }
        public bool Lock { get; }
        public bool FKMultiValue;
        public string FKMultiValuesPropName;

        public ExportTemplateAttribute(string displayName)
        {
            this.DisplayName = displayName;
            this.Ignore = false;
            this.FKSourceName = ExcelDropDownListType.None;
            this.FKStringProp = null;
            this.ParentDDLProperty = null;
            this.AllowNonFKValues = false;
            this.Lock = false;
        }

        public ExportTemplateAttribute(string displayName, bool lockColumn)
            : this(displayName)
        {
            this.Lock = lockColumn;
        }

        public ExportTemplateAttribute(string displayName, ExcelDropDownListType fkSourceName, bool fkMultiValue = false, string fkMultiValuesPropName = null, string fkStringProp = null, string parentDDLProperty = null, bool allowNonFKValues = false)
            : this(displayName)
        {
            this.FKSourceName = fkSourceName;
            this.FKMultiValue = fkMultiValue;
            this.FKMultiValuesPropName = fkMultiValuesPropName;
            this.FKStringProp = fkStringProp;
            this.ParentDDLProperty = parentDDLProperty;
            this.AllowNonFKValues = allowNonFKValues;
        }

        public ExportTemplateAttribute(bool ignore)
            : this(String.Empty)
        {
            this.Ignore = ignore;
        }
    }
}
