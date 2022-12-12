using System;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 类型字段与Excel Column名的映射关系
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class PropertyExcelColumnIndexAttribute : Attribute
    {
        public string ExcelColumnName;

        public PropertyExcelColumnIndexAttribute(string excelColumnName)
        {
            this.ExcelColumnName = excelColumnName;
        }

    }
}
