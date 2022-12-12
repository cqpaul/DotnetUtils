using Cqpaul.Dotnet.Util.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 类型字段 需要的校验 类型 Attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class PropertyValidateTypeAttribute : Attribute
    {
        public ExcelDataValidateType ValidateType;
        public List<string> DataOptions;
        public bool IsNeedShowColumnName;

        public PropertyValidateTypeAttribute(ExcelDataValidateType validateType, bool IsNeedShowColumnName = true, string[] dataOptions = null)
        {
            this.ValidateType = validateType;
            if (validateType == ExcelDataValidateType.IfNotNullMustInOptions || validateType == ExcelDataValidateType.MustInOptions)
            {
                if (dataOptions == null || dataOptions.Length == 0)
                {
                    throw new ArgumentNullException(nameof(dataOptions));
                }
                this.DataOptions = dataOptions.ToList();
            }
            this.IsNeedShowColumnName = IsNeedShowColumnName;
        }
    }
}
