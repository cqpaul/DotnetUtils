using Cqpaul.Dotnet.Util.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 必填校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MandatoryAttribute : ValidationAttribute
    {
        public override bool IsValid(object value)
        {
            return string.IsNullOrWhiteSpace(value.ToString());
        }
    }

    /// <summary>
    ///  文本最大长度校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MaxCharactorsAttribute : ValidationAttribute
    {
        private int max;
        public MaxCharactorsAttribute(int maxCharactorsLimit)
        {
            max = maxCharactorsLimit;
        }

        public override bool IsValid(object value)
        {
            return value.ToString().Length <= max;
        }
    }

    /// <summary>
    /// 数值最大值校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MaxValueAttribute : ValidationAttribute
    {
        private decimal max;
        public MaxValueAttribute(int maxValue)
        {
            max = maxValue;
        }

        public override bool IsValid(object value)
        {
            decimal decimalValue;
            bool valid = decimal.TryParse(value.ToString(), out decimalValue);
            return valid && decimalValue <= max;
        }
    }

    /// <summary>
    /// 数值范围校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ValueRangeAttribute : ValidationAttribute
    {
        private decimal min;
        private decimal max;
        public ValueRangeAttribute(decimal minValue, decimal maxValue)
        {
            min = minValue;
            max = maxValue;
        }

        public override bool IsValid(object value)
        {
            decimal decimalValue;
            bool valid = decimal.TryParse(value.ToString(), out decimalValue);
            return valid && decimalValue >= min && decimalValue <= max;
        }
    }

    /// <summary>
    /// 数值格式校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ExcelDataTypeAttribute : ValidationAttribute
    {
        private ExcelDataType type;
        public ExcelDataTypeAttribute(ExcelDataType dataType)
        {
            type = dataType;
        }

        public override bool IsValid(object value)
        {
            bool valid = true;
            int intValue;
            decimal decimalValue;
            DateTime dtValue;
            switch (type)
            {
                case ExcelDataType.Int:
                    valid = int.TryParse(value.ToString(), out intValue);
                    break;
                case ExcelDataType.Decimal:
                    valid = decimal.TryParse(value.ToString(), out decimalValue);
                    break;
                case ExcelDataType.PossitiveInt:
                    bool check = int.TryParse(value.ToString(), out intValue);
                    valid = check && intValue >= 0;
                    break;
                case ExcelDataType.PossitiveDecimal:
                    bool check1 = decimal.TryParse(value.ToString(), out decimalValue);
                    valid = check1 && decimalValue >= 0;
                    break;
                case ExcelDataType.Date:
                    bool formatResult1 = DateTime.TryParseExact(value.ToString(), "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    bool formatResult2 = DateTime.TryParseExact(value.ToString(), "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    bool formatResult3 = DateTime.TryParseExact(value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    bool formatResult4 = DateTime.TryParseExact(value.ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    bool formatResult5 = DateTime.TryParseExact(value.ToString(), "MM-dd-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    bool formatResult6 = DateTime.TryParseExact(value.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtValue);
                    valid = formatResult1 || formatResult2 || formatResult3 || formatResult4 || formatResult5 || formatResult6;
                    break;
                default:
                    value = true;
                    break;
            }
            return valid;
        }
    }

    /// <summary>
    /// 列表类型属性展开子元素校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ListItemValidationAttribute : ValidationAttribute
    {
        public override bool IsValid(object value)
        {
            bool valid = true;
            List<object> valueAsList = (value as IEnumerable<object>).Cast<object>().ToList();

            foreach (object obj in valueAsList)
            {
                ValidationContext validationContext = new ValidationContext(obj);
                List<ValidationResult> result = new List<ValidationResult>();
                valid &= Validator.TryValidateObject(obj, validationContext, result, true);
            }

            return valid;
        }
    }
}
