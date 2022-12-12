using System;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 类型字段，在Excel中的绝对定位
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class PropertyExcelLocationAttribute : Attribute
    {
        public (int rowIndex, string columnIndex) ExcelLocation;

        public PropertyExcelLocationAttribute(string columnIndex, int rowIndex)
        {
            ExcelLocation.columnIndex = columnIndex;
            ExcelLocation.rowIndex = rowIndex; 
        }
    }
}
