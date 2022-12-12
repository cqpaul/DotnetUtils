using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class DataValueFormatter
    {
        // Check Cell Value
        public static decimal DecimalFromCell(ICell sheetCell)
        {
            if (sheetCell == null)
            {
                return 0;
            }
            if (sheetCell.CellType == CellType.Formula)
            {
                return Convert.ToDecimal(sheetCell.NumericCellValue);
            }
            var value = sheetCell.ToString().Trim();

            return StringToDecimail(value);
        }

        public static double CellDataValueToDouble(string dataValue)
        {
            // 针对千分位的String
            if (dataValue.Contains(","))
            {
                dataValue = dataValue.Replace(",", "");
            }
            var isDouble = double.TryParse(dataValue, out double value);
            if (isDouble)
            {
                return Math.Round(value, 4);
            }
            return 0;
        }

        public static double CellDataValueToDouble(object dataValue)
        {
            return CellDataValueToDouble(dataValue.ToString());
        }

        public static decimal StringToDecimail(string dataValue)
        {
            var isDouble = decimal.TryParse(dataValue, out decimal value);
            if (isDouble)
            {
                return Math.Round(value, 4);
            }
            return 0;
        }

        public static int StringToInt(string dataValue)
        {
            var isInt = int.TryParse(dataValue, out int value);
            if (isInt)
            {
                return value;
            }
            return 0;
        }

        public static DateTime? StringToDatetime(string dataValue)
        {
            var isDouble = DateTime.TryParse(dataValue, out DateTime value);
            if (isDouble)
            {
                return value;
            }
            return null;
        }

        public static DateTime StringToExactDate(string dataValue)
        {
            // TODO：Need to check this funciton.[在用户上传日期数据的时候，需要加强校验。]
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();

            dtFormat.ShortDatePattern = "dd/MM/yyyy";

            var isDate = DateTime.TryParse(dataValue.Split("")[0], out DateTime date);
            if (isDate)
            {
                return date.Date;
            }
            else
            {
                return DateTime.Today.Date;
            }
        }

        public static DateTime StringToDateWithCn(string cellValue)
        {
            var isDate = DateTime.TryParse(cellValue, out DateTime dateValue);
            if (isDate)
            {
                return dateValue;
            }
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "dd-MM-yyyy";
            if (cellValue.Contains("月"))
            {
                var dataValue = cellValue.Replace("月", "");
                var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
                if (isDateFormat)
                {
                    return dateTime;
                }
            }
            else
            {
                DateTimeFormatInfo dtFormatII = new DateTimeFormatInfo();
                dtFormatII.ShortDatePattern = "dd/MM/yyyy";
                var isDateFormat = DateTime.TryParse(cellValue, dtFormatII, DateTimeStyles.None, out var dateTime);
                if (isDateFormat)
                {
                    return dateTime;
                }
            }
            bool isDouble = Double.TryParse(cellValue, out double doubleValue);
            if (isDouble)
            {
                var dateTimeValue = DateTime.FromOADate(doubleValue);
                return dateTimeValue;
            }


            throw new Exception($"日期（{cellValue}）格式不对。");
        }

        public static DateTime? StringToNullDateWithCn(string cellValue)
        {
            var isDate = DateTime.TryParse(cellValue, out DateTime dateValue);
            if (isDate)
            {
                return dateValue;
            }
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "dd-MM-yyyy";
            if (cellValue.Contains("月"))
            {
                var dataValue = cellValue.Replace("月", "");
                var isDateFormat = DateTime.TryParse(dataValue, dtFormat, DateTimeStyles.None, out var dateTime);
                if (isDateFormat)
                {
                    return dateTime;
                }
            }
            return null;
        }

    }
}
