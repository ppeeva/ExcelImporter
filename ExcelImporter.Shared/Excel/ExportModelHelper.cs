using ExcelImporter.Shared.Attributes;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ExcelImporter.Shared.Excel
{
    public static class ExportModelHelper
    {
        public static ExcelDropDownListType GetPropFKListType<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).FKSourceName;
            }

            return ExcelDropDownListType.None;
        }

        public static string GetPropFKStringProp<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).FKStringProp;
            }

            return null;
        }

        public static void SetPropertyValue<T>(string propName, object rawValue, T model)
        {
            if (string.IsNullOrWhiteSpace(propName) || model == null)
                return;

            PropertyInfo pInfo = typeof(T).GetProperty(propName);
            if (pInfo != null)
            {
                var propType = pInfo.PropertyType;
                if (propType == typeof(string))
                {
                    string propValue = rawValue?.ToString();
                    typeof(T).GetProperty(propName).SetValue(model, propValue);
                }
                else if (propType == typeof(DateTime))
                {
                    DateTime propValue = Constants.DefaultDate;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else if (DateTime.TryParseExact(rawValue.ToString(), "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParseExact(rawValue.ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParseExact(rawValue.ToString(), "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParse(rawValue.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else
                    {
                        if (double.TryParse(rawValue.ToString(), out double dateNum))
                        {
                            propValue = DateTime.FromOADate(dateNum);
                            typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                        }
                        else if (DateTime.TryParse(rawValue.ToString(), out propValue))
                        {
                            typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                        }
                        else
                            typeof(T).GetProperty(propName).SetValue(model, Constants.DefaultDate);
                    }

                }
                else if (propType == typeof(DateTime?))
                {
                    DateTime propValue = Constants.DefaultDate;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, null);
                    else if (DateTime.TryParseExact(rawValue.ToString(), "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParseExact(rawValue.ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParseExact(rawValue.ToString(), "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else if (DateTime.TryParse(rawValue.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                    else
                    {
                        if (double.TryParse(rawValue.ToString(), out double dateNum))
                        {
                            propValue = DateTime.FromOADate(dateNum);
                            typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                        }
                        else if (DateTime.TryParse(rawValue.ToString(), out propValue))
                        {
                            typeof(T).GetProperty(propName).SetValue(model, propValue.ToUniversalTime());
                        }
                        else
                        {
                            typeof(T).GetProperty(propName).SetValue(model, Constants.DefaultDate);
                        }
                    }

                }
                else if (propType == typeof(int))
                {
                    int propValue = 0;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else if (Int32.TryParse(rawValue.ToString(), out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, 0);
                }
                else if (propType == typeof(int?))
                {
                    int propValue = 0;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, null);
                    else if (Int32.TryParse(rawValue.ToString(), out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, null);
                }
                else if (propType == typeof(decimal))
                {
                    decimal propValue = 0;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else if (decimal.TryParse(rawValue.ToString().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, 0);
                }
                else if (propType == typeof(decimal?))
                {
                    decimal propValue = 0;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, null);
                    else if (decimal.TryParse(rawValue.ToString().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, null);
                }
                else if (propType == typeof(bool))
                {
                    bool propValue = false;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, false);
                    else if (Boolean.TryParse(rawValue.ToString(), out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, false);
                }
                else if (propType == typeof(bool?))
                {
                    bool propValue = false;
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, null);
                    else if (Boolean.TryParse(rawValue.ToString(), out propValue))
                        typeof(T).GetProperty(propName).SetValue(model, propValue);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, null);
                }
                else if (propType == typeof(IEnumerable<string>))
                {
                    if (rawValue == null)
                        typeof(T).GetProperty(propName).SetValue(model, null);
                    else
                        typeof(T).GetProperty(propName).SetValue(model, (List<string>)rawValue);

                }
            }
        }

        public static bool AreDropDownValuesValid<T>(T obj, IEnumerable<ExcelDDL> ddlLists, out string error)
        {
            error = "";
            StringBuilder allErrors = new StringBuilder();
            PropertyInfo[] membersToValidate = typeof(T)
                .GetProperties()
                .Where(p => p.GetCustomAttributes(typeof(ExportTemplateAttribute), true)
                        .Where(ca => ((ExportTemplateAttribute)ca).FKSourceName != ExcelDropDownListType.None)
                        .Any())
                .ToArray();

            foreach (PropertyInfo prop in membersToValidate)
            {
                ExcelDropDownListType listType = GetPropFKListType<T>(prop.Name);
                ExcelDDL list = ddlLists.FirstOrDefault(x => x.ListType == listType);
                if (list == null)
                {
                    continue;
                }

                object propValue = prop.GetValue(obj);
                if (propValue == null)
                {
                    continue;
                }

                bool isMultiValue = IsFKMultiValue<T>(prop.Name);
                if (isMultiValue)
                {
                    string collectionPropName = GetPropFKMultiValuesPropName<T>(prop.Name);
                    List<string> multiValuesCodes = new List<string>();
                    string[] multiValues = ((string)propValue).Split(";", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    foreach (string val in multiValues)
                    {
                        var found = list.Collection.Where(x => (string)x.Value == val || x.DisplayName == val).FirstOrDefault();
                        if (found == null)
                        {
                            allErrors.Append($"({prop.Name} - {val}) ");
                            continue;
                        }
                        else
                        {
                            multiValuesCodes.Add(found.Value.ToString());
                        }
                    }

                    if (allErrors.Length == 0 && !String.IsNullOrEmpty(collectionPropName))
                    {
                        SetPropertyValue(collectionPropName, multiValuesCodes, obj);
                    }
                }
                else
                {
                    bool isBoolValue = list.Collection.Any(x => x.Value.GetType() == typeof(bool));
                    bool isIntValue = list.Collection.Any(x => x.Value.GetType() == typeof(int));
                    bool isInvalid = false;

                    if (isBoolValue)
                    {
                        isInvalid = !list.Collection.Any(x => (bool)x.Value == (bool)propValue);
                    }
                    else if (isIntValue)
                    {
                        isInvalid = !list.Collection.Any(x => (int)x.Value == (int)propValue);
                    }
                    else
                    {
                        isInvalid = !list.Collection.Any(x => (string)x.Value == (string)propValue);
                    }

                    if (isInvalid)
                    {
                        allErrors.Append($"({prop.Name} - {propValue}) ");
                        continue;
                    }
                }
            }

            error = allErrors.ToString();

            return String.IsNullOrWhiteSpace(error);
        }

        public static bool IsFKMultiValue<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).FKMultiValue;
            }

            return false;
        }

        public static string GetPropFKMultiValuesPropName<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).FKMultiValuesPropName;
            }

            return null;
        }

        public static MemberInfo[] GetMembersToExport<T>()
        {
            MemberInfo[] membersToInclude = typeof(T)
                .GetProperties()
                .Where(p => p.GetCustomAttributes(typeof(ExportTemplateAttribute), true)
                        .Where(ca => !((ExportTemplateAttribute)ca).Ignore)
                        .Any())
                .ToArray();

            return membersToInclude;
        }

        public static string GetDisplayName<T>(string propName)
        {
            object[] attrInfo = typeof(T)
                            .GetProperty(propName)
                            .GetCustomAttributes(typeof(ExportTemplateAttribute), false);

            return (attrInfo != null && attrInfo.Length > 0) ? (attrInfo[0] as ExportTemplateAttribute).DisplayName : propName;
        }

        public static bool IsLocked<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).Lock;
            }

            return false;
        }

        public static bool AllowFKValues<T>(string propName)
        {
            PropertyInfo propInfo = typeof(T).GetProperty(propName);
            if (propInfo != null)
            {
                object[] attrInfo = propInfo.GetCustomAttributes(typeof(ExportTemplateAttribute), false);
                if (attrInfo != null && attrInfo.Length > 0)
                    return (attrInfo[0] as ExportTemplateAttribute).AllowNonFKValues;
            }

            return false;
        }


    }
}
