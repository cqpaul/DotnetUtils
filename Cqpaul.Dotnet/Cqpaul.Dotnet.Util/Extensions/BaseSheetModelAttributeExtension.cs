using Cqpaul.Dotnet.Util.Attributes;
using Cqpaul.Dotnet.Util.Enums;
using Cqpaul.Dotnet.Util.Helpers;
using Cqpaul.Dotnet.Util.Model;
using NPOI.SS.UserModel;
using System.ComponentModel;

namespace Cqpaul.Dotnet.Util.Extensions
{
    public static class BaseSheetModelAttributeExtension
    {
        /// <summary>
        /// 获取PropertyExcelColumnIndexAttribute的值，已转成INT Index
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static int GetPropertyExcelColumnIndexValue(this BaseSheetModel obj, string propertyName)
        {
            var referencedType = obj.GetType();
            int mappingColumnIndex = -1;
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(referencedType);
            if (properties != null)
            {
                PropertyDescriptor property = properties[propertyName];
                if (property != null)
                {
                    AttributeCollection attributeCollection = property.Attributes;
                    PropertyExcelColumnIndexAttribute mappingAttribute = (PropertyExcelColumnIndexAttribute)attributeCollection[typeof(PropertyExcelColumnIndexAttribute)];
                    if (mappingAttribute == null)
                    {
                        return mappingColumnIndex;
                    }
                    mappingColumnIndex = ExcelFunction.ToIndex(mappingAttribute.ExcelColumnName);
                }
            }
            if (mappingColumnIndex < 0)
            {
                throw new Exception($"没有找到{propertyName}的Excel映射位置。");
            }
            return mappingColumnIndex;
        }

        /// <summary>
        /// 获取PropertyExcelColumnIndexAttribute的值, 字母，非数字
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static string GetPropertyExcelColumnIndexStr(this BaseSheetModel obj, string propertyName)
        {
            var referencedType = obj.GetType();
            string mappingColumnStr = string.Empty;
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(referencedType);
            if (properties != null)
            {
                PropertyDescriptor property = properties[propertyName];
                if (property != null)
                {
                    AttributeCollection attributeCollection = property.Attributes;
                    PropertyExcelColumnIndexAttribute mappingAttribute = (PropertyExcelColumnIndexAttribute)attributeCollection[typeof(PropertyExcelColumnIndexAttribute)];
                    mappingColumnStr = mappingAttribute.ExcelColumnName;
                }
            }
            if (string.IsNullOrEmpty(mappingColumnStr))
            {
                throw new Exception($"没有找到{propertyName}的Excel映射位置。");
            }
            return mappingColumnStr;
        }

        /// <summary>
        /// 获取PropertyExcelLocationAttribute的值，对象属性在Excel中的绝对地址
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static (int rowIndex, int columnIndex) GetPropertyExcelLocationValue(this BaseSheetModel obj, string propertyName)
        {
            var referencedType = obj.GetType();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(referencedType);
            if (properties != null)
            {
                PropertyDescriptor property = properties[propertyName];
                // 只对普通类型的Property取绝对定位。
                if (property != null)
                {
                    AttributeCollection attributeCollection = property.Attributes;
                    PropertyExcelLocationAttribute locationAttribute = (PropertyExcelLocationAttribute)attributeCollection[typeof(PropertyExcelLocationAttribute)];
                    // 参数类似Cell Address, 代码中是从0开 -->> 所以locationAttribute.ExcelLocation.rowIndex - 1。
                    return (locationAttribute.ExcelLocation.rowIndex - 1, ExcelFunction.ToIndex(locationAttribute.ExcelLocation.columnIndex));
                }
            }
            throw new Exception($"没有找到{propertyName}的Excel映射位置。");
        }

        /// <summary>
        /// 通过获取Property PropertyValidateTypeAttribute, 来获取该字段的校验结果
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string GetPropertyValidateResult(this BaseSheetModel obj, ICell cell, string propertyName)
        {
            var referencedType = obj.GetType();
            string resultValidateMessage = string.Empty;
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(referencedType);
            if (properties != null)
            {
                PropertyDescriptor property = properties[propertyName];
                if (property != null)
                {
                    AttributeCollection attributeCollection = property.Attributes;
                    PropertyValidateTypeAttribute validateTypeAttribute = (PropertyValidateTypeAttribute)attributeCollection[typeof(PropertyValidateTypeAttribute)];
                    if (validateTypeAttribute != null)
                    {
                        ExcelDataValidateType validateType = validateTypeAttribute.ValidateType;
                        switch (validateType)
                        {
                            case ExcelDataValidateType.DoNotNeedValidate:
                                return string.Empty;
                            case ExcelDataValidateType.MustHave:
                                return NumberFunction.CheckCellValueMustHave(propertyName, cell, validateTypeAttribute.IsNeedShowColumnName);
                            case ExcelDataValidateType.MustNumber:
                                return NumberFunction.CheckCellValueMustNumber(propertyName, cell, validateTypeAttribute.IsNeedShowColumnName);
                            case ExcelDataValidateType.MustDate:
                                return NumberFunction.CheckCellValueMustDateFormat(propertyName, cell);
                            case ExcelDataValidateType.MustInOptions:
                                return NumberFunction.CheckCellValueMustWithOptions(propertyName, cell, validateTypeAttribute.DataOptions, validateTypeAttribute.IsNeedShowColumnName);
                            case ExcelDataValidateType.IfNotNullMustNumber:
                                return NumberFunction.CheckCellValueIfNotNullMustNumber(propertyName, cell, validateTypeAttribute.IsNeedShowColumnName);
                            case ExcelDataValidateType.IfNotNullMustDate:
                                return NumberFunction.CheckCellValueIfNotNullDateFormat(propertyName, cell);
                            case ExcelDataValidateType.IfNotNullMustInOptions:
                                return NumberFunction.CheckCellValueIfNotNullMustWithOptions(propertyName, cell, validateTypeAttribute.DataOptions);
                            case ExcelDataValidateType.MustDateOrString:
                                return NumberFunction.CheckCellValueMustDateOrStrFormat(propertyName, cell);
                            case ExcelDataValidateType.MustDateOrWString:
                                return NumberFunction.CheckCellValueMustDateOrNoStrFormat(propertyName, cell);
                            default:
                                return resultValidateMessage;
                        }
                    }
                }
            }
            return resultValidateMessage;
        }

        /// <summary>
        /// 获取类型的所有Property名
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static List<string> GetClassAllPropertyNameList(this BaseSheetModel obj)
        {
            var objectType = obj.GetType();
            List<string> allPropertyNameList = objectType.GetProperties().Where(a => !a.PropertyType.Name.Contains("List")).Select(a => a.Name).ToList();
            return allPropertyNameList;
        }

        /// <summary>
        /// 获取类型 为用户输入项的所有Property名
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static List<string> GetClassUserInputPropertyNameList(this BaseSheetModel obj)
        {
            var objectType = obj.GetType();
            List<string> userInputPropertyNames = new List<string>();
            var allProperties = objectType.GetProperties();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(objectType);
            foreach (var property in allProperties)
            {
                PropertyDescriptor propertyDescriptor = properties[property.Name];
                if (propertyDescriptor != null)
                {
                    AttributeCollection attributeCollection = propertyDescriptor.Attributes;
                    PropertyIsUserInputAttribute isUserInputAttribute = (PropertyIsUserInputAttribute)attributeCollection[typeof(PropertyIsUserInputAttribute)];
                    if (isUserInputAttribute != null)
                    {
                        if (isUserInputAttribute.IsUserInput)
                        {
                            userInputPropertyNames.Add(property.Name);
                        }
                    }
                }
            }
            return userInputPropertyNames;
        }

        /// <summary>
        /// 获取类型为 需要隐藏列的所有Property名
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static List<int> GetClassNeedHidngColumnList(this BaseSheetModel obj)
        {
            List<int> result = new List<int>();
            var objectType = obj.GetType();
            var allProperties = objectType.GetProperties();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(objectType);
            foreach (var property in allProperties)
            {
                PropertyDescriptor propertyDescriptor = properties[property.Name];
                if (propertyDescriptor != null)
                {
                    AttributeCollection attributeCollection = propertyDescriptor.Attributes;
                    PropertyExcelColumnHideAttribute ishideAttribute = (PropertyExcelColumnHideAttribute)attributeCollection[typeof(PropertyExcelColumnHideAttribute)];
                    if (ishideAttribute != null)
                    {
                        if (ishideAttribute.IsHiding)
                        {
                            PropertyExcelColumnIndexAttribute mappingAttribute = (PropertyExcelColumnIndexAttribute)attributeCollection[typeof(PropertyExcelColumnIndexAttribute)];
                            var mappingColumnIndex = ExcelFunction.ToIndex(mappingAttribute.ExcelColumnName);
                            result.Add(mappingColumnIndex);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 隐藏类型为 需要隐藏列的所有Property
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="sheet"></param>
        public static void SetExcelSheetHidden(this BaseSheetModel obj, ISheet sheet)
        {
            var columnIndex = GetClassNeedHidngColumnList(obj);
            ExcelFunction.HideColumnForSheet(sheet, columnIndex);
        }

        /// <summary>
        /// 依据 校验的 Options, 补充Excel下拉选项
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="sheet"></param>
        /// <param name="propertyName"></param>
        /// <param name="rowStartIndex"></param>
        /// <param name="rowEndIndex"></param>
        public static void SupplementDropdownListForCellsByValidationOptions(this BaseSheetModel obj, ISheet sheet, string propertyName, int rowStartIndex, int rowEndIndex)
        {
            var referencedType = obj.GetType();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(referencedType);
            if (properties != null)
            {
                PropertyDescriptor property = properties[propertyName];
                if (property != null)
                {
                    AttributeCollection attributeCollection = property.Attributes;
                    PropertyValidateTypeAttribute validateTypeAttribute = (PropertyValidateTypeAttribute)attributeCollection[typeof(PropertyValidateTypeAttribute)];
                    if (validateTypeAttribute != null)
                    {
                        ExcelDataValidateType validateType = validateTypeAttribute.ValidateType;
                        if (validateType == ExcelDataValidateType.MustInOptions || validateType == ExcelDataValidateType.IfNotNullMustInOptions)
                        {
                            PropertyExcelColumnIndexAttribute propertyExcelColumnIndexAttribute = (PropertyExcelColumnIndexAttribute)attributeCollection[typeof(PropertyExcelColumnIndexAttribute)];
                            ExcelFunction.SupplementDropdownListForCells(sheet, validateTypeAttribute.DataOptions, rowStartIndex, rowEndIndex, propertyExcelColumnIndexAttribute.ExcelColumnName);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 通过反射给属性赋值，给某对象，赋值某行
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dataRow"></param>
        /// <param name="isOnlyUserInputInfo">是否只取值用户填写的信息</param>
        public static void SetObjectValueByUploadDataRow(this BaseSheetModel obj, IRow dataRow, bool isOnlyUserInputInfo = true)
        {
            // 2022-03-28： 赋值给某行时候，用User input的属性，不然会覆盖 Template 信息
            List<string> itemProperties = obj.GetClassUserInputPropertyNameList();
            if (!isOnlyUserInputInfo)
            {
                itemProperties = obj.GetClassAllPropertyNameList();
            }
            obj.SetObjectValueByUploadDataRow(dataRow, itemProperties);
        }

        public static void SetObjectValueByUploadDataRow(this BaseSheetModel obj, IRow dataRow, List<string> propertyNames)
        {
            foreach (var propertyName in propertyNames)
            {
                var property = obj.GetType().GetProperty(propertyName);
                var propertyType = property.PropertyType;

                // 通过反射给属性赋值
                var index = obj.GetPropertyExcelColumnIndexValue(propertyName);
                if (index < 0) continue;
                var cell = dataRow.GetCell(index);
                var propertyValue = ExcelFunction.GetValueFromCell(cell);
                if (string.IsNullOrEmpty(propertyValue.ToString()))
                {
                    continue;
                }
                if (propertyType == typeof(DateTime))
                {
                    var updateValue = DataValueFormatter.StringToDateWithCn(propertyValue.ToString());
                    obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                }
                else if (propertyType == typeof(DateTime?))
                {
                    var updateValue = DataValueFormatter.StringToNullDateWithCn(propertyValue.ToString());
                    obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                }
                else if (propertyType == typeof(double) || propertyType == typeof(double?))
                {
                    var updateValue = DataValueFormatter.CellDataValueToDouble(propertyValue);
                    obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                }
                else if (propertyType == typeof(decimal) || propertyType == typeof(decimal?))
                {
                    var updateValue = DataValueFormatter.StringToDecimail(propertyValue.ToString());
                    obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                }
                else if (propertyType == typeof(int) || propertyType == typeof(int?))
                {
                    var updateValue = DataValueFormatter.StringToInt(propertyValue.ToString());
                    obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                }
                else
                {
                    obj.GetType().GetProperty(propertyName).SetValue(obj, propertyValue.ToString().Trim());
                }
            }
        }

        /// <summary>
        /// 通过反射给属性赋值，给某对象表Model，赋值某Excel表定位好的数据
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dataSheet"></param>
        /// <param name="isOnlyUserInputInfo">是否只取值用户填写的信息</param>
        public static void SetObjectValueByUploadDataSheet(this BaseSheetModel obj, ISheet dataSheet, bool isOnlyUserInputInfo = true)
        {
            // 2022-03-28： 赋值给某行时候，用User input的属性，不然会覆盖 Template 信息
            List<string> itemProperties = obj.GetClassUserInputPropertyNameList();
            if (!isOnlyUserInputInfo)
            {
                itemProperties = obj.GetClassAllPropertyNameList();
            }
            foreach (var propertyName in itemProperties)
            {
                // 通过反射给属性赋值
                var projectTypeLocation = obj.GetPropertyExcelLocationValue(propertyName);
                var cell = dataSheet.GetRow(projectTypeLocation.rowIndex)?.GetCell(projectTypeLocation.columnIndex);
                if (cell != null)
                {
                    var propertyValue = ExcelFunction.GetValueFromCell(cell);

                    var property = obj.GetType().GetProperty(propertyName);
                    var propertyType = property.PropertyType;
                    if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
                    {
                        var updateValue = DataValueFormatter.StringToDateWithCn(propertyValue.ToString());
                        obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                    }
                    else if (propertyType == typeof(double))
                    {
                        var updateValue = DataValueFormatter.CellDataValueToDouble(propertyValue);
                        obj.GetType().GetProperty(propertyName).SetValue(obj, updateValue);
                    }
                    else
                    {
                        obj.GetType().GetProperty(propertyName).SetValue(obj, propertyValue);
                    }
                }
            }
        }

        /// <summary>
        /// 通过反射将对象的属性值，填入对应的行；已过滤掉Template中，不为用户填写的属性。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dataRow"></param>
        public static void SetExcelRowValueByDataClass(this BaseSheetModel obj, IRow dataRow, bool isNeedToFilterZero = false)
        {
            if (obj == null) return;
            List<string> allDetailedItemUserInputPropertyNameList = obj.GetClassUserInputPropertyNameList();
            obj.SetExcelRowValueByDataClass(dataRow, allDetailedItemUserInputPropertyNameList, isNeedToFilterZero);
        }

        /// <summary>
        /// 指定行属性进行赋值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyNames"></param>
        /// <param name="dataRow"></param>
        /// <param name="isNeedToFilterZero"></param>
        public static void SetExcelRowValueByDataClass(this BaseSheetModel obj, IRow dataRow, List<string> propertyNames, bool isNeedToFilterZero = false)
        {
            if (dataRow == null) return;
            if (obj == null) return;
            foreach (string propertyName in propertyNames)
            {
                var property = obj.GetType().GetProperty(propertyName);
                var propertyType = property.PropertyType;
                var propertyValue = obj.GetType().GetProperty(propertyName).GetValue(obj);
                if (propertyValue != null)
                {
                    if ((propertyType == typeof(double) || propertyType == typeof(double?)) && !isNeedToFilterZero)
                    {
                        dataRow.GetCell(obj.GetPropertyExcelColumnIndexValue(propertyName)).SetCellDoubleValueIfNotNull(propertyValue.ToString());
                    }
                    else if ((propertyType == typeof(double) || propertyType == typeof(double?)) && isNeedToFilterZero)
                    {
                        if (Convert.ToDouble(propertyValue) != 0)
                        {
                            dataRow.GetCell(obj.GetPropertyExcelColumnIndexValue(propertyName)).SetCellDoubleValueIfNotNull(propertyValue.ToString());
                        }
                    }
                    else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
                    {
                        dataRow.GetCell(obj.GetPropertyExcelColumnIndexValue(propertyName)).SetCellDatetimeValueIfNotNull(propertyValue.ToString());
                    }
                    else
                    {
                        dataRow.GetCell(obj.GetPropertyExcelColumnIndexValue(propertyName))?.SetCellValue(propertyValue.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// 通过反射将对象的属性值，填入对应的Sheet；已过滤掉Template中，不为用户填写的属性。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="sheet"></param>
        public static void SetExcelSheetValueByDataClass(this BaseSheetModel obj, ISheet sheet)
        {
            List<string> allDetailedItemUserInputPropertyNameList = obj.GetClassUserInputPropertyNameList();
            foreach (string propertyName in allDetailedItemUserInputPropertyNameList)
            {
                var property = obj.GetType().GetProperty(propertyName);
                var propertyType = property.PropertyType;
                var propertyValue = obj.GetType().GetProperty(propertyName).GetValue(obj);
                var excelLocation = obj.GetPropertyExcelLocationValue(propertyName);
                if (propertyType == typeof(double))
                {
                    sheet.GetRow(excelLocation.rowIndex).GetCell(excelLocation.columnIndex).SetCellDoubleValueIfNotNull(propertyValue.ToString());
                }
                else if (propertyType == typeof(DateTime))
                {
                    sheet.GetRow(excelLocation.rowIndex).GetCell(excelLocation.columnIndex).SetCellDatetimeValueIfNotNull(propertyValue.ToString());
                }
                else
                {
                    sheet.GetRow(excelLocation.rowIndex).GetCell(excelLocation.columnIndex).SetCellValue(propertyValue.ToString());
                }
            }
        }

    }
}
