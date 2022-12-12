using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Helpers
{
    internal class NumberFunction
    {
        public static bool IsNumberic(string text)
        {
            try
            {
                bool result = int.TryParse(text, out int value);
                return result;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsDecimalNumberic(string text)
        {
            try
            {
                Convert.ToDecimal(text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Check Cell 必填项
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueMustHave(string columnName, ICell cell, bool isNeedShowColumnName = true)
        {
            if (cell == null)
            {
                return $"{columnName}列，都必填，不能有空值。";
            }

            string cellValue = null;
            if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
            {
                cellValue = cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.String)
            {
                cellValue = cell.StringCellValue;
            }
            else
            {
                cellValue = cell.ToString();
            }
            if (string.IsNullOrEmpty(cellValue))
            {
                if (isNeedShowColumnName)
                {
                    return $"（{cell.Address}）:{columnName}列，必填项，不能为空值。";
                }
                else
                {
                    return $"（{cell.Address}）：必填项，不能为空值。";
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Check Cell 必填且必须为Number
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueMustNumber(string columnName, ICell cell, bool isNeedShowColumnName = true)
        {
            if (cell == null)
            {
                return $"{columnName}列，都必填，不能有空值。";
            }
            string cellValue = null;
            if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
            {
                cellValue = cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.String)
            {
                cellValue = cell.StringCellValue;
            }
            else
            {
                cellValue = cell.ToString();
            }
            string mustHaveValue = CheckCellValueMustHave(columnName, cell);
            if (!string.IsNullOrEmpty(mustHaveValue))
            {
                return mustHaveValue;
            }
            var isNumber = decimal.TryParse(cellValue, out decimal numberValue);
            if (!isNumber)
            {
                if (isNeedShowColumnName)
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}）必须为数字。";
                }
                else
                {
                    return $"（{cell.Address}）：当前值（{cellValue}）必须为数字。";
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Check Cell 必填且必须为日期格式
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueMustDateFormat(string columnName, ICell cell)
        {
            if (cell == null)
            {
                return $"{columnName}列，都必填，不能有空值。";
            }

            var cellValue = cell.ToString().Trim();
            string mustHaveValue = CheckCellValueMustHave(columnName, cell);
            if (!string.IsNullOrEmpty(mustHaveValue))
            {
                return mustHaveValue;
            }
            if (cellValue.Length < 8)
            {
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为合理的日期格式。";
            }
            if (cellValue.Contains("年") && cellValue.Contains("月") && !cellValue.Contains("日"))
            {
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为合理的日期格式。";
            }
            var isDate = DateTime.TryParse(cellValue, out DateTime dateValue);
            if (isDate)
            {
                return string.Empty;
            }
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "dd-MM-yyyy";
            var dataValue = cellValue.ToString().Replace("月", "").Replace("年", "").Replace("日", "");
            var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
            if (isDateFormat)
            {
                return string.Empty;
            }
            return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为合理的日期格式。";
        }

        /// <summary>
        /// Check Cell 必填且值为‘未支付’或合理日期格式
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueMustDateOrStrFormat(string columnName, ICell cell)
        {
            var cellValue = cell.ToString().Trim();
            string mustHaveValue = CheckCellValueMustHave(columnName, cell);
            if (!string.IsNullOrEmpty(mustHaveValue))
            {
                return mustHaveValue;
            }
            if (cellValue.Trim() != "未支付")
            {
                if (cellValue.Length < 8)
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未支付’或合理的日期格式。";
                }
                if (cellValue.Contains("年") && cellValue.Contains("月") && !cellValue.Contains("日"))
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未支付’或合理的日期格式。";
                }
                var isDate = DateTime.TryParse(cellValue, out DateTime dateValue);
                if (isDate)
                {
                    return string.Empty;
                }
                DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
                dtFormat.ShortDatePattern = "dd-MM-yyyy";
                if (cellValue.ToString().Contains("月"))
                {
                    var dataValue = cellValue.ToString().Replace("月", "");
                    var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
                    if (isDateFormat)
                    {
                        return string.Empty;
                    }
                }
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未支付’或合理的日期格式。";
            }
            return string.Empty;
        }

        /// <summary>
        /// Check Cell 必填且值为‘未支付’或合理日期格式
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueMustDateOrNoStrFormat(string columnName, ICell cell)
        {
            var cellValue = cell.ToString().Trim();
            string mustHaveValue = CheckCellValueMustHave(columnName, cell);
            if (!string.IsNullOrEmpty(mustHaveValue))
            {
                return mustHaveValue;
            }
            if (cellValue.Trim() != "未收到")
            {
                if (cellValue.Length < 8)
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未收到’或合理的日期格式。";
                }
                if (cellValue.Contains("年") && cellValue.Contains("月") && !cellValue.Contains("日"))
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未收到’或合理的日期格式。";
                }
                var isDate = DateTime.TryParse(cellValue, out DateTime dateValue);
                if (isDate)
                {
                    return string.Empty;
                }
                DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
                dtFormat.ShortDatePattern = "dd-MM-yyyy";
                if (cellValue.ToString().Contains("月"))
                {
                    var dataValue = cellValue.ToString().Replace("月", "");
                    var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
                    if (isDateFormat)
                    {
                        return string.Empty;
                    }
                }
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），必须为‘未收到’或合理的日期格式。";
            }
            return string.Empty;
        }

        /// <summary>
        /// Check Cell 必填且必须在选项范围内
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static string CheckCellValueMustWithOptions(string columnName, ICell cell, List<string> options, bool isNeedShowColumnName = true)
        {
            if (cell == null)
            {
                return $"{columnName}列，都必填，不能有空值。";
            }

            var cellValue = cell.ToString().Trim();
            string mustHaveValue = CheckCellValueMustHave(columnName, cell);
            if (!string.IsNullOrEmpty(mustHaveValue))
            {
                return mustHaveValue;
            }
            if (!options.Contains(cellValue))
            {
                if (isNeedShowColumnName)
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}）必须在选项范围（{string.Join(',', options)}）内。";
                }
                else
                {
                    return $"（{cell.Address}）：当前值（{cellValue}）必须在选项范围（{string.Join(',', options)}）内。";
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Check Cell 非必填且必须在选项范围内
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static string CheckCellValueIfNotNullMustWithOptions(string columnName, ICell cell, List<string> options)
        {
            if (cell == null) //非必填可以为空
            {
                return string.Empty;
            }

            var cellValue = cell.ToString().Trim();
            if (string.IsNullOrEmpty(cellValue))
            {
                return string.Empty;
            }
            else if (!options.Contains(cellValue))
            {
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}）必须在选项范围（{string.Join(',', options)}）内。";
            }
            return string.Empty;
        }

        /// <summary>
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static string CheckCellValueIfNotNullMustEqualTo(string columnName, ICell cell, string value)
        {
            var cellValue = cell.ToString().Trim();
            if (string.IsNullOrEmpty(cellValue))
            {
                return string.Empty;
            }

            if (!value.Equals(cellValue))
            {
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}）必须为（{value}）。";
            }

            return string.Empty;
        }


        /// <summary>
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static string CheckCellValueMustEqualTo(string columnName, ICell cell, string value)
        {
            var cellValue = cell.ToString().Trim();

            if (!value.Equals(cellValue))
            {
                return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}）必须为（{value}）。";
            }

            return string.Empty;
        }

        /// <summary>
        /// Check Cell 非必填日期
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueIfNotNullDateFormat(string columnName, ICell cell)
        {
            if (cell == null) //非必填可以为空
            {
                return string.Empty;
            }

            var cellValue = cell.ToString().Trim();
            if (string.IsNullOrEmpty(cellValue))
            {
                return string.Empty;
            }
            var isDate = DateTime.TryParse(cellValue, out var dateValue);
            if (isDate)
            {
                return string.Empty;

            }
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "dd-MM-yyyy";
            if (cellValue.ToString().Contains("月"))
            {
                var dataValue = cellValue.ToString().Replace("月", "");
                var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
                if (isDateFormat)
                {
                    return string.Empty;
                }
            }
            return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），如果有值，则必须为合理的日期格式。";
        }

        /// <summary>
        /// Check Cell 非必填数字
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CheckCellValueIfNotNullMustNumber(string columnName, ICell cell, bool isNeedShowColumnName = true)
        {
            if (cell == null) //非必填可以为空
            {
                return string.Empty;
            }

            string cellValue = null;
            if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
            {
                cellValue = cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.String)
            {
                cellValue = cell.StringCellValue;
            }
            else
            {
                cellValue = cell.ToString();
            }

            if (string.IsNullOrEmpty(cellValue))
            {
                return string.Empty;
            }
            var isNumber = decimal.TryParse(cellValue, out decimal numberValue);
            if (!isNumber)
            {
                if (isNeedShowColumnName)
                {
                    return $"（{cell.Address}）：{columnName}列，当前值（{cellValue}），如果有值， 则必须为数字。";
                }
                else
                {
                    return $"（{cell.Address}）：当前值（{cellValue}），如果有值， 则必须为数字。";
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// 中国版的四舍五入。  注意.net自带的Math.Round(value,digial)不是四舍五入，而是四色六入五取偶。
        /// </summary>
        /// <param name="value"></param>
        /// <param name="decimals"></param>
        /// <returns></returns>
        public static decimal ChineseRound(decimal value, int digit)
        {
            double vt = Math.Pow(10, digit);
            //1.乘以倍数 + 0.5
            decimal vx = value * (decimal)vt + 0.5M;
            //2.向下取整
            decimal temp = Math.Floor(vx);
            //3.再除以倍数
            return (temp / (decimal)vt);
        }

        /// <summary>
        /// 得到两个数相比较后增长率. 
        /// </summary>
        /// <param name="firstNumber"></param>
        /// <param name="secondNumber"></param>
        /// <param name="multiplicationOneHundred">是否乘以100，这样前端就不需要乘100了，直接加百分号
        /// 比如54%，我们会返回54，而不是0.54.前端自己加百分号，不用乘以100
        /// </param>
        /// <returns></returns>
        public static decimal GetIncreaseRate(decimal? firstNumber, decimal? secondNumber, bool needMultipOneHundred = true)
        {
            if (firstNumber == null || secondNumber == null || firstNumber == 0)
            {
                return 0;  //目前这些异常都返回0.不报错
            }
            else
            {
                decimal rate = (secondNumber.Value - firstNumber.Value) / firstNumber.Value;
                if (needMultipOneHundred)
                {
                    rate = rate * 100;
                }

                return rate;


            }
        }

        /// <summary>
        /// 得到两个数相除后的值，
        /// </summary>
        /// <param name="firstNumber"></param>
        /// <param name="secondNumber"></param>
        /// <param name="multiplicationOneHundred">是否乘以100，这样前端就不需要乘100了，直接加百分号
        /// 比如54%，我们会返回54，而不是0.54.前端自己加百分号，不用乘以100
        /// </param>
        /// <returns></returns>
        public static string Division(decimal? fenzi, decimal? fenmu, bool needMultipOneHundred = true, int digit = 2)
        {
            //强制保留N位小数，防止小数末位数是0的时候，会被去掉。拼接成这种格式:ToString("0.00");
            var digitFormat = "0.";
            for (int i = 0; i < digit; i++)
            {
                digitFormat = digitFormat + "0";
            }

            string resultString = "";
            if (fenzi == null || fenmu == null || fenmu == 0)
            {
                resultString = "无数据";
            }
            else
            {
                decimal result = fenzi.Value / fenmu.Value;

                if (needMultipOneHundred)//返回带有百分号的形式的字符串
                {
                    result = result * 100;
                    result = ChineseRound(result, digit);
                    resultString = result.ToString(digitFormat) + "%";// eg:保留两位小数。  ToString("0.00");
                }
                else
                {
                    result = ChineseRound(result, digit);
                    resultString = result.ToString(digitFormat);
                }


            }

            return resultString;
        }

        /// <summary>
        /// 获取某cell的INT值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static int GetSheetIntNumber(ICell cell)
        {
            if (cell == null)
            {
                throw new Exception($"（{cell.Address}）：数据Cell不存在。");
            }
            var value = cell.ToString();
            if (string.IsNullOrEmpty(value))
            {
                throw new Exception($"（{cell.Address}）：数据值不能为空");
            }
            bool isIntData = int.TryParse(value, out int dataValue);
            if (isIntData)
            {
                return dataValue;
            }
            else
            {
                throw new Exception($"（{cell.Address}）：数据值（{value}）不是合理数字");
            }
        }
    }
}
