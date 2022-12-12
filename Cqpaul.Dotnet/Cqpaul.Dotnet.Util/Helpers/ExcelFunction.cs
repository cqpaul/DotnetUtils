using Cqpaul.Dotnet.Util.Extensions;
using Cqpaul.Dotnet.Util.Model;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class ExcelFunction
    {
        /// <summary>
        /// 添加合并的单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        public static void AddMergedRegion(ISheet sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            //先删除已有合并单元格，再添加合并单元格
            RemoveMergedRegion(sheet, startRow, endRow, startColumn, endColumn);
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startColumn, endColumn));
        }

        /// <summary>
        /// 获取Sheet中填写有效数据的行（即序号为数字的行）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sheetRowStartIndex"></param>
        /// <returns></returns>
        public static List<IRow> GetSheetRowList(ISheet sheet, int startRowIndex = 1, int startColumnIndex = 0)
        {
            List<IRow> rows = new List<IRow>();
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;
                for (int i = startRowIndex; i <= rowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        ICell cell = sheet.GetRow(i).GetCell(startColumnIndex);
                        if (cell != null && !string.IsNullOrEmpty(GetValueFromCell(cell).ToString()))
                        {
                            rows.Add(sheet.GetRow(i));
                        }
                    }
                }
            }
            return rows;
        }

        public static List<string> ValidateExcelSheetData<T>(ISheet sheet, int startRowIndex = 1) where T : BaseSheetModel, new()
        {
            List<string> errors = new List<string>();
            List<IRow> rows = ExcelFunction.GetSheetRowList(sheet);
            T dataModel = new T();
            var userInputProperties = dataModel.GetClassUserInputPropertyNameList();
            foreach (IRow row in rows)
            {
                foreach (var property in userInputProperties)
                {
                    string error = dataModel.GetPropertyValidateResult(row.GetCell(dataModel.GetPropertyExcelColumnIndexValue(property)), property);
                    if (!string.IsNullOrEmpty(error))
                    {
                        errors.Add(error);
                    }

                }
            }

            return errors;
        }


        /// <summary>
        /// 移除某个合并的单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        private static void RemoveMergedRegion(ISheet sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            int MergedCount = sheet.NumMergedRegions;
            for (int i = MergedCount - 1; i >= 0; i--)
            {
                /**
                CellRangeAddress对象属性有：FirstColumn，FirstRow，LastColumn，LastRow 进行操作 取消合并单元格
                **/
                var temp = sheet.GetMergedRegion(i);
                if (temp.FirstRow == startRow && temp.LastRow == endRow && temp.FirstColumn == startColumn && temp.LastColumn == endColumn)
                {
                    sheet.RemoveMergedRegion(i);
                }
            }
        }

        /// <summary>
        /// 得到某个cell的数字值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="needFormulaCellValue"></param>
        /// <returns></returns>
        public static double? GetNumberValueFromCell(ICell cell, XSSFFormulaEvaluator evalor = null)
        {
            double? value = null;
            if (cell.CellType == CellType.Formula)
            {
                //先尝试解析表达式取值，如果去报错，则取NumericCellValue得值
                try
                {
                    var cellValue = GetValueFromCell(cell, true, evalor);
                    if (NumberFunction.IsDecimalNumberic(cellValue.ToString()))
                    {
                        value = Convert.ToDouble(cellValue);
                    }
                }
                catch
                {
                    value = cell.NumericCellValue;
                }
            }
            else
            {
                //先尝试取NumericCellValue，如果报错则取formula的值
                try
                {
                    value = cell.NumericCellValue;
                    if (cell.CellType == CellType.Blank && value == 0)
                    {
                        value = null; // 有时候用户填了0，有时候是把blank类型的cell错误的解析为了0.这个时候，我们认为value是空的，而不是0
                    }
                }
                catch
                {
                    var cellValue = GetValueFromCell(cell, true);
                    if (NumberFunction.IsNumberic(cellValue.ToString()))
                    {
                        value = Convert.ToDouble(cellValue);
                    }
                }
            }



            return value;
        }

        /// <summary>
        /// 是否需要解析表达式cell的值，如果不需要，则表达值cell格子的值解析为空字符串（有时候我们需要解析速度，而表达式解析有点慢，如果不需要解析表达式的话，就可以禁用）
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="needFormulaCellValue"></param>
        /// <returns></returns>
        public static object GetValueFromCell(ICell cell, bool needFormulaCellValue = true, XSSFFormulaEvaluator evalor = null)
        {
            object value = "";
            if (cell == null)
            {
                return value;
            }

            switch (cell.CellType)
            {
                case CellType.Blank:
                    value = "";
                    break;

                case CellType.Numeric:
                    short format = cell.CellStyle.DataFormat;
                    if (format == 14 || format == 31 || format == 57 || format == 58 || format == 165 || cell.ToString().Contains("月"))
                    {
                        DateTime dateValue;
                        try
                        {
                            dateValue = cell.DateCellValue;
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                dateValue = DateTime.FromOADate(cell.NumericCellValue);
                            }
                            catch
                            {
                                //如果date确实解析不出来，则认为就是一个数字，直接返回Numeric值
                                value = cell.NumericCellValue.ToString();
                                return value;
                            }
                        }

                        if (dateValue.Year == 1900)
                        {
                            //有时候0会被解析为1900年这样的日期，这时候我们认为该cell不是一个有效的日期，而是一个数字
                            try
                            {
                                value = cell.NumericCellValue.ToString();
                            }
                            catch (Exception ee)
                            {
                                //try中确实不能解析为数字，我们再返回日期。
                                value = cell.DateCellValue;
                            }
                        }
                        else
                        {
                            value = dateValue;
                        }

                    }
                    else
                    {
                        try
                        {
                            value = cell.NumericCellValue.ToString();
                        }
                        catch (Exception ee)
                        {
                            value = cell.DateCellValue;
                        }
                    }

                    //short format = cell.CellStyle.DataFormat;
                    //if (format == 14 || format == 31 || format == 57 || format == 58 || format == 165 || cell.ToString().Contains("月"))
                    //{
                    //    try
                    //    {
                    //        value = cell.DateCellValue;
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        value = DateTime.FromOADate(cell.NumericCellValue);
                    //    }

                    //}
                    //else
                    //{
                    //    try
                    //    {
                    //        value = cell.NumericCellValue.ToString();
                    //    }
                    //    catch (Exception ee)
                    //    {
                    //        value = cell.DateCellValue;
                    //    }
                    //}
                    break;

                case CellType.String:
                    value = cell.ToString().Replace('\u00A0', ' ').Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");//.StringCellValue;
                    break;

                case CellType.Formula:
                    if (needFormulaCellValue)
                    {
                        try
                        {
                            //if (evalor == null && cell.CachedFormulaResultType == CellType.Numeric)
                            //{
                            //    value = cell.NumericCellValue;
                            //    break;
                            //}
                            if (evalor == null)
                            {
                                evalor = new XSSFFormulaEvaluator(cell.Sheet.Workbook);//这种方式得到的Workbook，有时候计算会有问题，最好是通过参数把evalor传进来
                            }
                            //如果提供了evalor参数，则用evalor计算表达式
                            //cell = evalor.EvaluateInCell(cell); //这种方式虽然计算出来了结果，但是会破坏原cell的celltype。甚至让其formula丢失
                            evalor.EvaluateFormulaCellEnum(cell);//这种方式只会计算结果，不会破坏celltype

                            if (cell.CachedFormulaResultType == CellType.Numeric)
                            {
                                value = cell.NumericCellValue;
                            }
                            else if (cell.CachedFormulaResultType == CellType.String)
                            {
                                value = cell.RichStringCellValue;
                            }
                            else
                            {
                                value = "";//处理formula计算结果为空的情况。
                            }
                            //if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                            //{
                            //    value = cell.NumericCellValue;
                            //}
                            //else if (cell.CellType == CellType.String)
                            //{
                            //    value = cell.StringCellValue;
                            //}
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                //value = cell.StringCellValue.Replace('\u00A0', ' ').Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");
                                value = cell.NumericCellValue;
                            }
                            catch
                            {
                                value = cell.ErrorCellValue;//直接把错误展示出来
                            }
                        }
                    }
                    else
                    {
                        value = cell.ToString();// 不解析表达式的值。因为解析formula速度有点慢，有时候不需要解析，直接返回空。
                        break;
                    }
                    break;

                default:
                    value = cell.ToString().Trim();//.StringCellValue;
                    break;
            }

            return value;
        }

        /// <summary>
        /// 如果cell是空的话，返回异常提醒,并指明cell的位置。如果不为空，则返回空
        /// </summary>
        public static string GetErrorIfCellIsEmpty(ICell cell)
        {
            if (CellIsEmpty(cell))
            {
                return "Sheet(" + cell.Sheet.SheetName + ")的Cell(" + cell.Address + ")不能为空.";
            }

            return "";
        }

        /// <summary>
        /// 如果cell不是有效数字的话，返回异常提醒,并指明cell的位置。否则返回空
        /// </summary>
        public static string GetErrorIfCellValueIsNotValidNumber(ICell cell)
        {
            var value = GetNumberValueFromCell(cell);

            if (value == null)
            {
                return "Sheet(" + cell.Sheet.SheetName + ")的Cell(" + cell.Address + ")不是一个有效数字.";
            }

            return "";
        }

        /// <summary>
        ///  如果cell不是有效日期的话，返回异常提醒,并指明cell的位置。否则返回空
        /// </summary>
        public static string GetErrorIfCellValueIsNotValidDate(ICell cell)
        {
            try
            {
                var date = cell.DateCellValue;
            }
            catch
            {
                return "Sheet(" + cell.Sheet.SheetName + ")的Cell(" + cell.Address + ")不是一个有效日期.";
            }

            if (cell.DateCellValue == null)
            {
                //如果DateCellValue没值，则尝试解析StringCellValue
                try
                {
                    DateTime.Parse(cell.StringCellValue);
                }
                catch
                {
                    return "Sheet(" + cell.Sheet.SheetName + ")的Cell(" + cell.Address + ")不是一个有效日期.";
                }
            }
            else if (cell.DateCellValue < new DateTime(1901, 1, 1)) //输入数字的时候，会被解析为1900年
            {
                return "Sheet(" + cell.Sheet.SheetName + ")的Cell(" + cell.Address + ")不是一个有效日期.";
            }

            return "";
        }

        /// <summary>
        /// 判断某个cell是否为空
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool CellIsEmpty(ICell cell)
        {
            if (cell == null)
            {
                return true;
            }

            if (cell.CellType == CellType.String && string.IsNullOrEmpty(cell.StringCellValue))
            {
                return true;
            }
            else if (cell.CellType == CellType.Blank)
            {
                return true;
            }
            //else if (cell.CellType == CellType.Numeric && cell.NumericCellValue == 0)
            //{
            //    return true;
            //}
            else if (cell.CellType == CellType.Formula)
            {
                try
                {
                    var value = cell.CellNumberValue();
                    if (value == null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch
                {
                    var value = cell.StringCellValue;
                    if (string.IsNullOrEmpty(value))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            else
            {
                return false;
            }
        }

        ///// <summary>
        ///// 得到所有的sheet name和index的信息
        ///// </summary>
        ///// <param name="fs"></param>
        ///// <returns></returns>
        //public static List<ExcelInfoVO> GetExcelSheetNameList(Stream fs)
        //{
        //    List<ExcelInfoVO> excelInfoVoList = new List<ExcelInfoVO>();
        //    var workbook = new XSSFWorkbook(fs);
        //    workbook.SetForceFormulaRecalculation(true);
        //    var sheetNumber = workbook.NumberOfSheets;
        //    for (int i = 0; i < sheetNumber; i++)
        //    {
        //        ExcelInfoVO excelInfoVo = new ExcelInfoVO
        //        {
        //            SheetIndex = i,
        //            SheetName = workbook.GetSheetName(i).Trim(),
        //            Sheet = workbook.GetSheetAt(i)
        //        };
        //        // 两种隐藏
        //        excelInfoVo.IsSheetHidden = workbook.IsSheetHidden(i) || workbook.IsSheetVeryHidden(i);
        //        excelInfoVoList.Add(excelInfoVo);
        //    }
        //    return excelInfoVoList;
        //}

        ///// <summary>
        ///// 得到所有的sheet name和index的信息
        ///// </summary>
        ///// <param name="workbook"></param>
        ///// <returns></returns>
        //public static List<ExcelInfoVO> GetExcelSheetNameList(IWorkbook workbook)
        //{
        //    List<ExcelInfoVO> excelInfoVoList = new List<ExcelInfoVO>();
        //    var sheetNumber = workbook.NumberOfSheets;
        //    for (int i = 0; i < sheetNumber; i++)
        //    {
        //        ExcelInfoVO excelInfoVo = new ExcelInfoVO
        //        {
        //            SheetIndex = i,
        //            SheetName = workbook.GetSheetName(i).Trim(),
        //            Sheet = workbook.GetSheetAt(i)
        //        };
        //        // 两种隐藏
        //        excelInfoVo.IsSheetHidden = workbook.IsSheetHidden(i) || workbook.IsSheetVeryHidden(i);
        //        excelInfoVoList.Add(excelInfoVo);
        //    }
        //    return excelInfoVoList;
        //}

        ///// <summary>
        ///// 得到所有的sheet name和index的信息， 及Datatable
        ///// </summary>
        ///// <param name="fs"></param>
        ///// <param name="startRowIdx"></param>
        ///// <returns></returns>
        //public static List<ExcelInfoVO> GetExcelInfoVOWithDataTables(Stream fs, int startRowIdx = 5)
        //{
        //    List<ExcelInfoVO> excelInfoVoList = new List<ExcelInfoVO>();
        //    var workbook = new XSSFWorkbook(fs);
        //    workbook.SetForceFormulaRecalculation(true);
        //    var sheetNumber = workbook.NumberOfSheets;
        //    for (int i = 0; i < sheetNumber; i++)
        //    {
        //        ExcelInfoVO excelInfoVo = new ExcelInfoVO
        //        {
        //            SheetIndex = i,
        //            SheetName = workbook.GetSheetName(i).Trim(),
        //            Sheet = workbook.GetSheetAt(i),
        //            SheetDataTable = GetDatatableFromSheet(workbook.GetSheetAt(i), startRowIdx, false, true)
        //        };
        //        excelInfoVo.IsSheetHidden = workbook.IsSheetHidden(i) || workbook.IsSheetVeryHidden(i);
        //        excelInfoVoList.Add(excelInfoVo);
        //    }
        //    return excelInfoVoList;
        //}

        /// <summary>
        /// 得到sheet的数量
        /// </summary>
        /// <param name="fs"></param>
        /// <returns></returns>
        public static int GetExcelSheetNumber(Stream fs)
        {
            var workbook = new XSSFWorkbook(fs);
            return workbook.NumberOfSheets;
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="fs">stream</param>
        /// <param name="sheetIndex">读取excel中的第几个sheet</param>
        /// <param name="dataStartRow">有效数据从第几行开始(有时候前面几行是注释之类的无用数据)</param>
        /// <param name="isColumnNameExist">列名行是否存在</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(Stream fs, int sheetIndex = 0, int dataStartRow = 0, bool isColumnNameExist = true)
        {
            DataTable dataTable = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            try
            {
                workbook = new XSSFWorkbook(fs);

                if (workbook != null)
                {
                    sheet = workbook.GetSheetAt(sheetIndex);//读取第一个sheet，当然也可以循环读取每个sheet
                    dataTable = GetDatatableFromSheet(sheet, dataStartRow, isColumnNameExist);
                }

                return dataTable;
            }
            catch (Exception e)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="fs">stream</param>
        /// <param name="sheetName">excel中sheet名</param>
        /// <param name="dataStartRow">有效数据（或者说列名行）从第几行开始(有时候前面几行是注释之类的无用数据)</param>
        /// <param name="isColumnNameExist">列名行是否存在</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(Stream fs, string sheetName, int dataStartRow = 0, bool isColumnNameExist = true)
        {
            DataTable dataTable = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            try
            {
                workbook = new XSSFWorkbook(fs);

                if (workbook != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        throw new Exception("Cannot find sheet with sheet name " + sheetName);
                    }
                    dataTable = GetDatatableFromSheet(sheet, dataStartRow, isColumnNameExist);
                }

                return dataTable;
            }
            catch (Exception e)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="fs">stream</param>
        /// <param name="sheetName">excel中sheet名</param>
        /// <returns>返回datatable</returns>
        public static ISheet ExcelToSheet(Stream fs, string sheetName)
        {
            DataTable dataTable = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            try
            {
                workbook = new XSSFWorkbook(fs);

                if (workbook != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        throw new Exception("Cannot find sheet with sheet name " + sheetName);
                    }

                    return sheet;
                }
                else
                {
                    throw new Exception("Cannot find workbook ");
                }
            }
            catch (Exception e)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 解析某个sheet为DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataStartRow">有效数据从第几行开始(有时候前面几行是注释之类的无用数据)</param>
        /// <param name="isColumnNameExist">列名行是否存在</param>
        /// <param name="needFormulaCellValue">是否需要解析表达式cell的值，如果不需要，则表达值cell格子的值解析为空字符串（有时候我们需要解析速度，而表达式解析有点慢，如果不需要解析表达式的话，就可以禁用）</param>
        /// <returns></returns>
        public static DataTable GetDatatableFromSheet(ISheet sheet, int dataStartRow = 0, bool isColumnNameExist = true, bool needFormulaCellValue = true)
        {
            DataTable dataTable = new DataTable();
            if (sheet != null)
            {
                if (sheet.SheetName.Contains("其他无形资产"))
                {

                }
                sheet.ForceFormulaRecalculation = true;
                IWorkbook wb = sheet.Workbook;
                XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(wb);
                int rowCount = sheet.LastRowNum;//总行数
                if (rowCount > 0)
                {
                    IRow firstRow = sheet.GetRow(dataStartRow);//从第dataStartRow行开始读取有效数据
                    int cellCount = firstRow.LastCellNum;//列数

                    ICell cell;
                    DataColumn column;
                    //构建datatable的列
                    if (isColumnNameExist)
                    {
                        for (int i = 0; i < cellCount; ++i)
                        {
                            cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                if (cell.StringCellValue != null)
                                {
                                    //取值并去掉换行符,制表符，特殊空格等符号
                                    column = new DataColumn(cell.StringCellValue.Trim().Replace('\u00A0', ' ').Replace("\n", " ").Replace("\t", " ").Replace("\r", " "));
                                    dataTable.Columns.Add(column);
                                }
                            }
                            else
                            {
                                //空的列名,也要加进去占位，不然后面的数据会错位
                                dataTable.Columns.Add("");
                            }
                        }
                    }
                    else
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            column = new DataColumn("column" + (i + 1));
                            dataTable.Columns.Add(column);
                        }
                    }
                    //填充行  (第dataStartRow行是列名，则数据从（dataStartRow + 1）行读取)
                    for (int i = (dataStartRow + 1); i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        DataRow dataRow = dataTable.NewRow();
                        for (int j = 0; j < cellCount; ++j)
                        {
                            cell = row.GetCell(j);
                            if (cell == null)
                            {
                                dataRow[j] = "";
                            }
                            else
                            {
                                if (needFormulaCellValue)
                                {
                                    dataRow[j] = GetValueFromCell(cell, true, evaluator);
                                }
                                else
                                {
                                    dataRow[j] = GetValueFromCell(cell, false);
                                }
                            }
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }
            return dataTable;
        }

        /// <summary>
        /// 将Sheet拷贝到制定的Excel文件
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <param name="sheets"></param>
        public static IWorkbook CopySheetsToWorkbookFile(byte[] fileBytes, List<ISheet> sheets)
        {
            var workbook = new XSSFWorkbook(StreamHelper.BytesToStream(fileBytes));
            foreach (var sheet in sheets)
            {
                sheet.CopyTo(workbook, sheet.SheetName, true, true, true);
            }
            return workbook;
            //using (var outputStream = new ByteArrayOutputStream())
            //{
            //    workbook.Write(outputStream);
            //    var byteArray = outputStream.ToByteArray();
            //    outputStream.Close();
            //    return byteArray;
            //}
        }

        /// <summary>
        /// 将文件的指定Sheet转换问HTML内容
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <param name="sheetName"></param>
        public static string ExcelToHtml(byte[] fileBytes, string sheetName)
        {
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();
            var workbook = new XSSFWorkbook(StreamHelper.BytesToStream(fileBytes));
            // 设置输出参数
            // excelToHtmlConverter.OutputColumnHeaders = false;
            // excelToHtmlConverter.OutputHiddenColumns = false;
            // excelToHtmlConverter.OutputHiddenRows = false;
            // excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = false;
            // excelToHtmlConverter.OutputRowNumbers = false;
            // excelToHtmlConverter.UseDivsToSpan = true;

            // 处理的Excel文件
            excelToHtmlConverter.ProcessWorkbook(workbook);
            //添加表格样式
            excelToHtmlConverter.Document.InnerXml =
                excelToHtmlConverter.Document.InnerXml.Insert(
                    excelToHtmlConverter.Document.InnerXml.IndexOf("<head>", 0) + 6,
                    @"<style>table, td, th{border:1px solid green;}th{background-color:green;color:white;}</style>"
                );

            string htmlContent = excelToHtmlConverter.Document.InnerXml;
            return htmlContent;
        }

        /// <summary>
        /// 取Sheet的Datatable，并求总和
        /// </summary>
        /// <param name="sheetDataTable"></param>
        /// <param name="sumValue"></param>
        /// <param name="sumValueIndex"></param>
        /// <returns></returns>
        public static byte[] GetSheetSerializeDataTableAndSumValue(DataTable sheetDataTable, out decimal sumValue, int sumValueIndex = 0)
        {
            sumValue = 0;
            // SumValue 需要减掉 减记 的部分 -->> 需统一 减记的名称。[2022-07-01]: 看起来都是“减”打头。
            List<string> 减记StrList = new List<string>() {
            "减：账面坏账准备",
            "减：减值准备",
            "减：跌价准备",
            "减：账面坏账准备",
            };
            if (sumValueIndex != 0)
            {
                foreach (DataRow row in sheetDataTable.Rows)
                {
                    bool isNumberLine = int.TryParse(row[0].ToString(), out int numberIdx);
                    if (isNumberLine)
                    {
                        bool isSumValueLine = decimal.TryParse(row[sumValueIndex].ToString(), out decimal rowSumValue);
                        if (isSumValueLine)
                        {
                            sumValue = sumValue + rowSumValue;
                        }
                    }
                    if (row[0].ToString().Trim().StartsWith("减"))
                    {
                        var 减记obj = row[sumValueIndex];
                        if (减记obj != null)
                        {
                            bool is减记Value = decimal.TryParse(减记obj.ToString(), out decimal 减记Value);
                            if (is减记Value)
                            {
                                sumValue = sumValue - 减记Value;
                            }
                        }
                    }
                }
            }

            byte[] byteDate = SerializeHelper.SerializeToBinary(sheetDataTable);
            return byteDate;
        }


        /// <summary>
        /// Excel列字母转数字
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static int ToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index - 1;
        }

        /// <summary>
        /// Excel数字转列字母
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static string ToName(int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }

        /// <summary>
        /// 往已有的sheet中插入一些行，
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="newAddRowsCount">新加的行数</param>
        /// <param name="newRowStartIndex">从哪一行开始插入新行</param>
        /// <param name="locker"></param>
        public static void AddNewRowsIntoSheet(ISheet sheet, int newAddRowsCount, int newRowStartIndex, object locker = null)
        {
            //选择新行插入的前一行作为复制样式style和formula的对象行（注意这里必须是新行前面的行(newRowStartIndex - 1)，不然伴随着新行的插入，后面有些顺序会不准确）
            var sourceRowIndex = newRowStartIndex - 1;
            var sourceRow = sheet.GetRow(sourceRowIndex);

            if (locker != null)
            {
                //需要锁
                lock (locker)
                {
                    sheet.ShiftRows(newRowStartIndex, sheet.LastRowNum, newAddRowsCount, true, false);
                }
            }
            else
            {
                //不需要锁
                sheet.ShiftRows(newRowStartIndex, sheet.LastRowNum, newAddRowsCount, true, false);
            }

            //存储sourceRow各个cell的style信息，后面复制各行的时候，直接从这里拿，速度更快
            Dictionary<int, ICellStyle> sourceRowCellsStyleDic = new Dictionary<int, ICellStyle>();
            Dictionary<int, string> sourceRowCellsFormulaDic = new Dictionary<int, string>();

            for (int colIndex = 0; colIndex < sourceRow.LastCellNum; colIndex++)
            {
                var sourceCell = sourceRow.GetCell(colIndex);

                if (sourceCell == null)
                {
                    continue;
                }

                if (sourceCell.CellStyle != null)
                {
                    sourceRowCellsStyleDic.Add(colIndex, sourceCell.CellStyle);
                }

                if (sourceCell.CellType == CellType.Formula && !string.IsNullOrEmpty(sourceCell.CellFormula))
                {
                    sourceRowCellsFormulaDic.Add(colIndex, sourceCell.CellFormula);
                }
            }

            //逐cell替换其style和formula等属性
            for (int i = newRowStartIndex; i < newRowStartIndex + newAddRowsCount; i++)
            {
                IRow newRow = null;
                if (locker != null)
                {
                    //需要锁
                    lock (locker)
                    {
                        newRow = sheet.CreateRow(i);
                    }
                }
                else
                {
                    //不需要锁
                    newRow = sheet.CreateRow(i);
                }

                if (sourceRow.RowStyle != null)
                {
                    newRow.RowStyle = sourceRow.RowStyle;
                }

                newRow.Height = sourceRow.Height;

                for (int colIndex = 0; colIndex < sourceRow.LastCellNum; colIndex++)
                {
                    var newCell = newRow.CreateCell(colIndex);
                    //设置单元格样式　　　　
                    if (sourceRowCellsStyleDic.ContainsKey(colIndex))
                    {
                        newCell.CellStyle = sourceRowCellsStyleDic.GetValueOrDefault(colIndex);
                    }

                    if (sourceRowCellsFormulaDic.ContainsKey(colIndex))
                    {
                        //这段关于formula的代码有点耗时
                        newCell.CellFormula = sourceRowCellsFormulaDic.GetValueOrDefault(colIndex).Replace((sourceRowIndex + 1).ToString(), (i + 1).ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Excel打印分页处理
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="dataRows">数据总行数</param>
        /// <param name="pageDataNum">每页总行数</param>
        /// <param name="templateDataRows">每页数据行数</param>
        /// <param name="titleRowsNum">列名占用行数</param>
        public static void ExcelPaging(XSSFWorkbook workbook, ISheet sheet, int dataRows, int pageDataNum, int templateDataRows, int titleRowsNum, int footerMergeCount = 4, int assessorMergeCount = 5)
        {
            //页数
            int pageNum = (int)Math.Ceiling(((double)dataRows / (double)templateDataRows));

            if (pageNum > 1)
            {
                //整数页
                for (var i = 1; i < pageNum; i++)
                {
                    sheet.ShiftRows(i * pageDataNum - 4, sheet.LastRowNum, 9 + titleRowsNum, true, false);
                    //插入表尾
                    for (int j = 0; j < 3; j++)
                    {
                        sheet.CopyRow(sheet.LastRowNum - j, i * pageDataNum - j - 2);
                        if (j == 2)
                        {
                            AddMergedRegion(sheet, i * pageDataNum - j - 2, i * pageDataNum - j - 2, 0, footerMergeCount);
                            AddMergedRegion(sheet, i * pageDataNum - j - 2, i * pageDataNum - j - 2, sheet.GetRow(i * pageDataNum - j - 2).LastCellNum - assessorMergeCount, sheet.GetRow(i * pageDataNum - j - 2).LastCellNum - 1);
                            if (sheet.GetRow(i * pageDataNum - j - 2).GetCell(sheet.GetRow(i * pageDataNum - j - 2).LastCellNum - assessorMergeCount) != null)
                            {
                                sheet.GetRow(i * pageDataNum - j - 2).GetCell(sheet.GetRow(i * pageDataNum - j - 2).LastCellNum - assessorMergeCount).CellStyle.Alignment = HorizontalAlignment.Right;
                            }
                        }
                        else
                        {
                            AddMergedRegion(sheet, i * pageDataNum - j - 2, i * pageDataNum - j - 2, 0, sheet.GetRow(i * pageDataNum - j - 2).LastCellNum - 1);
                        }
                    }
                    //插入表头
                    for (int j = 0; j < 5 + titleRowsNum; j++)
                    {
                        sheet.CopyRow(j, i * pageDataNum + j);
                    }
                }
                // 插入页数
                for (var i = 0; i < pageNum; i++)
                {
                    var cellIndex = pageDataNum * i + 3;
                    sheet.GetRow(cellIndex)?.GetCell(ToIndex("A"))?.SetCellValue($"共 {pageNum} 页第 {i + 1} 页");
                }

                AddMergedRegion(sheet, sheet.LastRowNum - 2, sheet.LastRowNum - 2, sheet.GetRow(sheet.LastRowNum - 2).LastCellNum - assessorMergeCount, sheet.GetRow(sheet.LastRowNum - 2).LastCellNum - 1);
                if (sheet.GetRow(sheet.LastRowNum - 2).GetCell(sheet.GetRow(sheet.LastRowNum - 2).LastCellNum - assessorMergeCount) != null)
                {
                    sheet.GetRow(sheet.LastRowNum - 2).GetCell(sheet.GetRow(sheet.LastRowNum - 2).LastCellNum - assessorMergeCount).CellStyle.Alignment = HorizontalAlignment.Right;
                }
            }
        }

        //已copyindex开始复制Sheet行
        public static void CopySheetRows(ISheet sheet, int copyIndex, int copyRows)
        {
            for (int i = 0; i < copyRows; i++)
            {
                sheet.CopyRow(copyIndex, copyIndex + 1);
            }
        }


        /// <summary>
        /// 获取某行Excel数据的值
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private static List<Tuple<CellType, object>> GetSheetRowDataValue(IRow row)
        {
            List<Tuple<CellType, object>> result = new List<Tuple<CellType, object>>();
            int columns = row.LastCellNum;
            for (int i = 0; i < columns; i++)
            {
                if (row.GetCell(i) == null)
                {
                    continue;
                }
                Tuple<CellType, object> cell = new Tuple<CellType, object>(row.GetCell(i).CellType, row.GetCell(i)?.ToString());
                result.Add(cell);
            }
            return result;
        }


        /// <summary>
        /// 复制Sheet的前copyRows条，中间间隔intervalNumberRows行，共复制copyCount次
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="copyRows"></param>
        /// <param name="intervalNumberRows"></param>
        /// <param name="copyCount"></param>
        public static void ExcelCopyTableRows(ISheet sheet, int copyRows, int intervalNumberRows, int copyCount)
        {
            for (int i = 1; i <= copyCount; i++)
            {
                for (int j = 0; j < copyRows; j++)
                {
                    int shiftCount = (copyRows + intervalNumberRows) * i + j;
                    sheet.CopyRow(j, shiftCount);
                }
            }
        }

        /// <summary>
        /// 获取Sheet某行某列的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataStartRowIndex">查找数据的开始行</param>
        /// <param name="findTypeName">需要查找的数据</param>
        /// <param name="findColumnIndex">查找数据的列</param>
        /// <param name="columnIndex">取具体某列的值</param>
        /// <returns></returns>
        public static string GetColumnCellValueByFindRowValue(ISheet sheet, int dataStartRowIndex, string findTypeName, int findColumnIndex, int columnIndex, XSSFFormulaEvaluator evalor = null)
        {
            for (int i = dataStartRowIndex; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(columnIndex);
                    if (cell == null) continue;
                    if (row.GetCell(findColumnIndex).ToString().Trim().Replace(" ", "").StartsWith(findTypeName))
                    {
                        if (evalor != null)
                        {
                            var dataValue = GetValueFromCell(cell, true, evalor);
                            return dataValue.ToString();
                        }
                        if (cell.CellType == CellType.Formula)
                        {
                            return GetValueFromCell(cell, true).ToString();
                        }
                        return row.GetCell(columnIndex).ToString();
                    }
                }
            }
            return null;
            // throw new Exception("没有找到对应的数据。");
        }

        /// <summary>
        /// 查找某Sheet的某行某值的列index
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="findValue"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static string GetColumnIndexByFindRowValue(ISheet sheet, int rowIndex, string findValue, XSSFFormulaEvaluator evalor = null)
        {
            string columnIndex = String.Empty;
            sheet.ForceFormulaRecalculation = true;
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                throw new Exception($"未能找到{sheet.SheetName}的{rowIndex}行数据。");
            }
            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell cell = row.GetCell(i);
                if (cell.ToString().Trim() == findValue)
                {
                    columnIndex = ToName(i);
                }
                if (cell.CellType == CellType.Formula)
                {
                    if (cell.NumericCellValue.ToString() == findValue)
                    {
                        columnIndex = ToName(i);
                    }
                }
                if (evalor != null)
                {
                    var dataValue = GetValueFromCell(cell, true, evalor);
                    if (dataValue.ToString() == findValue)
                    {
                        columnIndex = ToName(i);
                    }
                }
            }
            if (string.IsNullOrEmpty(columnIndex))
            {
                throw new Exception($"未能找到{sheet.SheetName}的{rowIndex}行数据：{findValue}。");
            }
            return columnIndex;
        }

        /// <summary>
        /// 按序号获取行
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="findIndex"></param>
        /// <param name="findValue"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IRow GetDataRowBySerialNumber(List<IRow> rows, int findIndex, string findValue)
        {
            foreach (IRow row in rows)
            {
                if (row.GetCell(findIndex).ToString().Trim() == findValue)
                {
                    return row;
                }
            }
            return null;
            // throw new Exception("没有找到对应的行");
        }

        /// <summary>
        /// 在某列查找某数据的行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="findIndex"></param>
        /// <param name="findValue"></param>
        /// <returns></returns>
        public static IRow GetDataRowByFindValue(ISheet sheet, int findIndex, string findValue)
        {
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(findIndex);
                    if (cell != null)
                    {
                        if (cell.ToString().Trim() == findValue)
                        {
                            return row;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 在sheet中多条件查询值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startFindRowIndex"></param>
        /// <param name="findQuerys"></param>
        /// <param name="valueColumnIndex"></param>
        /// <returns></returns>
        public static string GetSheetCellValueWithMultiQuery(ISheet sheet, int startFindRowIndex, List<(int columnIndex, string findValue)> findQuerys, int valueColumnIndex)
        {
            string findValue = string.Empty;
            for (int i = startFindRowIndex; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                int trueCount = 0;
                foreach (var query in findQuerys)
                {
                    if (row.GetCell(query.columnIndex)?.ToString() == query.findValue)
                    {
                        trueCount++;
                    }
                }
                if (trueCount == findQuerys.Count())
                {
                    ICell cell = row.GetCell(valueColumnIndex);
                    if (cell.CellType == CellType.Formula)
                    {
                        // findValue = cell.NumericCellValue.ToString();
                        findValue = ExcelFunction.GetValueFromCell(cell).ToString();
                    }
                    else
                    {
                        findValue = cell.ToString();
                    }
                }
            }
            if (findValue == string.Empty)
            {
                var strList = findQuerys.Select(query => query.findValue.ToString()).ToList();
                var result = string.Join("：", strList);
                throw new Exception($"没有在{sheet.SheetName}中，找到符合条件（{result}）的值。");
            }
            return findValue;
        }

        /// <summary>
        /// 在多个行组Item中，查找某Cell的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="keySerialNumber">作为条件查找的序号值</param>
        /// <param name="serialNumberStartRowIndex">序号的起始行</param>
        /// <param name="serialNumberColumnIndex">Sheet中序号值所在的列</param>
        /// <param name="gapNumberWithSerialNumberIndex">取值行与序号行之间的GAP</param>
        /// <param name="groupItemRowsCount">每组行的条数</param>
        /// <param name="valueColumnIndex">查找值所在的列</param>
        /// <returns></returns>
        public static string GetCellValueBySerialNumberWithMultiItem(ISheet sheet, string keySerialNumber, int serialNumberStartRowIndex, string serialNumberColumnIndex, int gapNumberWithSerialNumberIndex, int groupItemRowsCount, string valueColumnIndex)
        {
            string cellValue = string.Empty;
            for (int i = serialNumberStartRowIndex; i < sheet.LastRowNum;)
            {
                string cellSerialNumber = sheet.GetRow(i).GetCell(ExcelFunction.ToIndex(serialNumberColumnIndex)).ToString();
                if (cellSerialNumber == keySerialNumber)
                {
                    ICell cell = sheet.GetRow(i + gapNumberWithSerialNumberIndex).GetCell(ExcelFunction.ToIndex(valueColumnIndex));
                    if (cell.CellType == CellType.Formula)
                    {
                        // cellValue = cell.NumericCellValue.ToString();
                        cellValue = ExcelFunction.GetValueFromCell(cell).ToString();
                    }
                    else
                    {
                        cellValue = cell.ToString();
                    }
                }
                i += groupItemRowsCount;
            }
            if (cellValue == string.Empty)
            {
                throw new Exception($"没有在{sheet.SheetName}中，找到序号为{keySerialNumber}的值。");
            }
            return cellValue;
        }

        /// <summary>
        ///遍历sheet某一列的所有值，找出其值等于某个给定值的行。返回row number
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="conditionString"></param>
        ///  <param name="isFuzzyMatching">匹配conditionString的时候，是否是模糊匹配</param>
        /// <returns></returns>
        public static int GetRowIndexWithColumnIndexAndTargetSearchStringList(ISheet sheet, int conditionColumnIndex, List<string> conditionStringList, bool isFuzzyMatching = false, bool throwExceptionWhenNotFound = true)
        {
            foreach (var conditionString in conditionStringList)
            {
                var index = GetRowIndexWithColumnIndexAndTargetSearchString(sheet, conditionColumnIndex, conditionString, isFuzzyMatching, throwExceptionWhenNotFound: false);
                if (index > 0) //找到了
                {
                    return index;
                }
            }

            if (throwExceptionWhenNotFound)
            {
                throw new Exception("找不到条件行。" + String.Join(",", conditionStringList));
            }
            else
            {
                return -1; //找不到返回-1
            }
        }

        /// <summary>
        ///遍历sheet某一列的所有值，找出其值等于某个给定值的行。返回row number
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="conditionString"></param>
        ///  <param name="isFuzzyMatching">匹配conditionString的时候，是否是模糊匹配</param>
        /// <returns></returns>
        public static int GetRowIndexWithColumnIndexAndTargetSearchString(ISheet sheet, int conditionColumnIndex, string conditionString, bool isFuzzyMatching = false, bool throwExceptionWhenNotFound = true)
        {
            conditionString = conditionString.Trim();
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                var cell = row.GetCell(conditionColumnIndex);
                if (cell == null)
                {
                    continue;
                }
                // 这里的needFormulaCellValue必须为false,两个原因
                //1. 通过属性找行数，属性一定是一个特定的名称，而不会是表达式，所以没有必要计算表达式
                //2. 这里会遍历某列的所有行，如果强制解析表达式，会把该列中的某些表达式的cell破坏。得不偿失(后来修改了内部实现，应该不会破坏cell了)
                var cellValue = GetValueFromCell(cell, needFormulaCellValue: true);
                if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString()))
                {
                    if (isFuzzyMatching)
                    {
                        if (cellValue.ToString().Contains(conditionString))
                        {
                            return i;
                        }
                    }
                    else if (cellValue.ToString().Equals(conditionString))
                    {
                        return i;
                    }

                }
            }

            if (throwExceptionWhenNotFound)
            {
                throw new Exception("找不到条件行。" + conditionString);
            }
            else
            {
                return -1; //找不到返回-1
            }
        }

        /// <summary>
        /// 遍历sheet某一列的所有值，找出其值等于某个给定值的行，进而得到该行的某个cell的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="targetString"></param>
        /// <param name="isFuzzyMatching">匹配conditionString的时候，是不是模糊匹配</param>
        /// <returns></returns>
        public static ICell GetCellWithColumnIndexAndTargetSearchString(ISheet sheet, int conditionColumnIndex, string conditionString, int targetCellColumnIndex, bool isFuzzyMatching = false)
        {
            //先找到"相应板块" 的row number
            int rowIndex = GetRowIndexWithColumnIndexAndTargetSearchString(sheet, conditionColumnIndex, conditionString, isFuzzyMatching);
            return sheet.GetRow(rowIndex).GetCell(targetCellColumnIndex);
        }

        /// <summary>
        /// 隐藏所选的sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="hideSheets"></param>
        public static void HideSelectSheets(IWorkbook workbook, List<string> hideSheets)
        {
            if (hideSheets.Any())
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet selectSheet = workbook.GetSheetAt(i);
                    if (hideSheets.Contains(selectSheet.SheetName))
                    {
                        workbook.SetSheetHidden(i, SheetState.VeryHidden);
                    }
                }
            }
        }

        /// <summary>
        /// Sheet中补充下拉选项
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dropdownList"></param>
        /// <param name="rowStartIndex"></param>
        /// <param name="rowEndIndex"></param>
        /// <param name="columnIndex"></param>
        public static void SupplementDropdownListForCells(ISheet sheet, List<string> dropdownList, int rowStartIndex, int rowEndIndex, string columnIndex)
        {
            if (rowStartIndex <= rowEndIndex)
            {
                XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet);
                XSSFDataValidationConstraint constraint = (XSSFDataValidationConstraint)dvHelper.CreateExplicitListConstraint(dropdownList.ToArray());
                CellRangeAddressList cellList = new CellRangeAddressList(rowStartIndex, rowEndIndex, ExcelFunction.ToIndex(columnIndex), ExcelFunction.ToIndex(columnIndex));
                XSSFDataValidation validation = (XSSFDataValidation)dvHelper.CreateValidation(constraint, cellList);
                validation.ShowErrorBox = true;
                sheet.AddValidationData(validation);
            }
        }

        /// <summary>
        /// 通过公式设置下拉选项
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="formula"></param>
        /// <param name="rowStartIndex"></param>
        /// <param name="rowEndIndex"></param>
        /// <param name="columnIndex"></param>
        public static void SupplementDropdownListForCells(ISheet sheet, string formula, int rowStartIndex, int rowEndIndex, string columnIndex)
        {
            if (rowStartIndex <= rowEndIndex)
            {
                XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet);
                XSSFDataValidationConstraint constraint = (XSSFDataValidationConstraint)dvHelper.CreateFormulaListConstraint(formula);
                CellRangeAddressList cellList = new CellRangeAddressList(rowStartIndex, rowEndIndex, ToIndex(columnIndex), ToIndex(columnIndex));
                XSSFDataValidation validation = (XSSFDataValidation)dvHelper.CreateValidation(constraint, cellList);
                validation.ShowErrorBox = true;
                sheet.AddValidationData(validation);
            }
        }

        /// <summary>
        /// 获取组表格的数据Rows
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="groupRowCount">从开始行到结尾行的行数</param>
        /// <param name="findStartStr">启始查找的行名称</param>
        /// <returns></returns>
        public static List<List<IRow>> GetSheetRowsGroups(ISheet sheet, int groupRowCount, string findStartStr)
        {
            List<List<IRow>> groupRows = new List<List<IRow>>();

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                ICell cell = sheet.GetRow(i)?.GetCell(0);
                if (cell != null)
                {
                    if (cell.ToString().Trim() == findStartStr)
                    {
                        List<IRow> groupRow = new List<IRow>();
                        for (int j = i; j < i + groupRowCount; j++)
                        {
                            groupRow.Add(sheet.GetRow(j));
                        }
                        i = i + groupRowCount;
                        groupRows.Add(groupRow);
                    }
                }
            }
            return groupRows;
        }

        /// <summary>
        /// 返回对应列的英文序号。
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static string GetColumnEnlishIndexByColumnIndex(int columnIndex)
        {
            string columnName = string.Empty;
            string k = string.Empty;
            if (columnIndex <= 25)//如果是26以内的就读取单个数据(因为index从0开始)
            {
                columnName = ColumnEnlishIndexLetter(columnIndex);
            }
            else//如果不是就递归计算
            {
                var leftNum = columnIndex - 26;
                k = ColumnEnlishIndexLetter(leftNum);//先获得组合的最后一个字母
                columnName = GetColumnEnlishIndexByColumnIndex(columnIndex / 26 - 1) + k;//计算组合字母
            }
            return columnName;
        }

        /// <summary>
               /// 获取单个列号
               /// </summary>
               /// <param name="number"></param>
               /// <returns></returns>
        public static string ColumnEnlishIndexLetter(int number)
        {
            string letter = string.Empty;
            switch (number)
            {
                case 0: letter = "A"; break;
                case 1: letter = "B"; break;
                case 2: letter = "C"; break;
                case 3: letter = "D"; break;
                case 4: letter = "E"; break;
                case 5: letter = "F"; break;
                case 6: letter = "G"; break;
                case 7: letter = "H"; break;
                case 8: letter = "I"; break;
                case 9: letter = "J"; break;
                case 10: letter = "K"; break;
                case 11: letter = "L"; break;
                case 12: letter = "M"; break;
                case 13: letter = "N"; break;
                case 14: letter = "O"; break;
                case 15: letter = "P"; break;
                case 16: letter = "Q"; break;
                case 17: letter = "R"; break;
                case 18: letter = "S"; break;
                case 19: letter = "T"; break;
                case 20: letter = "U"; break;
                case 21: letter = "V"; break;
                case 22: letter = "W"; break;
                case 23: letter = "X"; break;
                case 24: letter = "Y"; break;
                case 25: letter = "Z"; break;
                default: return "";
            }
            return letter;
        }

        /// <summary>
        /// 获取Sheet指定Cell的值，参数类似A2，方法内部对row index减了1.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static double GetSheetCellDoubleValue(ISheet sheet, string columnIndex, int rowIndex)
        {
            var cell = sheet.GetRow(rowIndex - 1).GetCell(ToIndex(columnIndex));
            if (cell != null)
            {
                var objectValue = GetValueFromCell(cell, true);
                return Convert.ToDouble(objectValue);
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 查找某列某值最后一次出现的行Index
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="finndValue"></param>
        /// <param name="findColumnIndex"></param>
        /// <returns></returns>
        public static int GetFindValueLastRowIndex(ISheet sheet, string finndValue, string findColumnIndex)
        {
            int rowIndex = 0;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var cell = row.GetCell(ToIndex(findColumnIndex));
                    if (cell != null)
                    {
                        string cellValue = cell.ToString();
                        if (cellValue.Trim().StartsWith(finndValue.Trim()))
                        {
                            rowIndex = i;
                        }
                    }
                }
            }
            return rowIndex;
        }



        /// <summary>
        /// sheet隐藏列
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndexList">需要隐藏的列的下标</param>
        public static void HideColumnForSheet(ISheet sheet, List<int> columnIndexList)
        {
            if (columnIndexList.Any())
            {
                foreach (var index in columnIndexList)
                {
                    sheet.SetColumnHidden(index, true);
                }
            }
        }

        /// <summary>
        /// 通过sheet名称获取sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetNameList"></param>
        /// <returns></returns>
        public static List<ISheet> GetSheetBySubject(XSSFWorkbook workbook, List<string> sheetNameList)
        {
            List<ISheet> sheets = new List<ISheet>();
            if (sheetNameList != null && sheetNameList.Any())
            {
                foreach (var sheetName in sheetNameList)
                {
                    ISheet sheet = null;
                    sheet = workbook.GetSheet($"{sheetName}评估明细表");
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheet(sheetName);
                    }
                    if (sheet == null)
                    {
                        continue;
                    }
                    else
                    {
                        sheets.Add(sheet);
                    }
                }
            }

            return sheets;
        }

        /// <summary>
        /// 检查Excel Sheet是否存在后，重新生成新名称
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        public static string GenSheetName(IWorkbook workbook, string sheetName, int i = 1)
        {
            ISheet existSheet = workbook.GetSheet(sheetName);
            if (existSheet != null)
            {
                sheetName = $"{sheetName}_{i}";
                sheetName = GenSheetName(workbook, sheetName, i++);
            }
            return sheetName;
        }

        /// <summary>
        /// 隐藏行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndexList"></param>
        public static void SetRowHidden(ISheet sheet, List<int> rowIndexList)
        {
            if (sheet != null && rowIndexList.Any())
            {
                foreach (var index in rowIndexList)
                {
                    IRow row = sheet.GetRow(index);
                    if (row != null)
                    {
                        row.Hidden = true;
                    }
                }
            }
        }

        public static XSSFCellStyle GetHightCellStyle(IWorkbook workbook)
        {
            #region Style
            XSSFColor redColor = new XSSFColor();
            byte[] colorRgb = { (byte)255, (byte)0, (byte)0 };
            redColor.SetRgb(colorRgb);

            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = "Arial";

            // 数字样式
            XSSFCellStyle numberCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            numberCellStyle.FillForegroundColorColor = redColor;
            numberCellStyle.FillPattern = FillPattern.SolidForeground;

            numberCellStyle.BorderTop = BorderStyle.Thin;
            numberCellStyle.BorderBottom = BorderStyle.Thin;
            numberCellStyle.BorderLeft = BorderStyle.Thin;
            numberCellStyle.BorderRight = BorderStyle.Thin;
            numberCellStyle.VerticalAlignment = VerticalAlignment.Center;
            numberCellStyle.Alignment = HorizontalAlignment.Right;
            numberCellStyle.SetFont(font);

            IDataFormat dataformat = workbook.CreateDataFormat();
            numberCellStyle.DataFormat = dataformat.GetFormat("#,##0");
            #endregion

            return numberCellStyle;
        }



        /// <summary>
        /// 比较两个Cell，如果不同，则更新背景色
        /// </summary>
        /// <param name="uploadCell"></param>
        /// <param name="templateCell"></param>
        public static void CompareTwoSheetCell(ICell uploadCell, ICell templateCell, ICellStyle highLightStyle)
        {
            bool needHighLight = false;

            if (templateCell.CellType != uploadCell.CellType)
            {
                needHighLight = true;
            }
            else
            {
                // 对每种类型进行比较
                switch (templateCell.CellType)
                {
                    case CellType.Numeric:
                        if (templateCell.NumericCellValue != uploadCell.NumericCellValue)
                        {
                            needHighLight = true;
                        }
                        break;
                    case CellType.String:
                        if (templateCell.StringCellValue.Trim() != uploadCell.StringCellValue.Trim())
                        {
                            needHighLight = true;
                        }
                        break;
                    case CellType.Formula:
                        if (templateCell.CellFormula.ToString().Trim() != uploadCell.CellFormula.ToString().Trim())
                        {
                            needHighLight = true;
                        }
                        break;
                    default:
                        break;
                }
            }
            if (needHighLight)
            {
                uploadCell.CellStyle = highLightStyle;
            }
        }

        /// <summary>
        /// Check Sheet 是否 Valid
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static bool CheckSheetIsValid(IWorkbook workbook, ISheet sheet)
        {
            if (workbook == null || sheet == null)
            {
                return false;
            }
            if (workbook.IsSheetVeryHidden(workbook.GetSheetIndex(sheet.SheetName)))
            {
                return false;
            }
            if (workbook.IsSheetHidden(workbook.GetSheetIndex(sheet.SheetName)))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 分页版添加页脚
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="footerYear"></param>
        public static void SetSheetPrintFooter(ISheet sheet, string footerYear)
        {
            sheet.Footer.Right = $"© {footerYear} 上海德勤资产评估有限公司 - 机密";
        }

        public static byte[] SetSheetPrintSetUp(byte[] bytes, bool isLandscape = true, bool needFit = true, string footerYear = "")
        {
            byte[] result = null;
            Stream stream = new MemoryStream(bytes);
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                if (!CheckSheetIsValid(workbook, sheet)) continue;
                //设置打印相关格式
                sheet.PrintSetup.NoColor = true;
                sheet.PrintSetup.Landscape = isLandscape;
                sheet.PrintSetup.PaperSize = (int)PaperSize.A4_Small;
                sheet.HorizontallyCenter = true;
                if (needFit)
                {
                    sheet.FitToPage = true;
                    sheet.PrintSetup.FitWidth = 1;
                    sheet.PrintSetup.FitHeight = 0;
                }
                // 设置打印页边距
                sheet.SetMargin(MarginType.LeftMargin, (double)2.225 / 3); // 1cm 约等于 0.8 左边距为1.8cm 1/0.8 * 1.8 = 2.25
                sheet.SetMargin(MarginType.RightMargin, (double)2.225 / 3);
                sheet.SetMargin(MarginType.TopMargin, (double)3.1 / 3);
                sheet.SetMargin(MarginType.BottomMargin, (double)2.5 / 3);

                // 设置页脚
                if (!string.IsNullOrEmpty(footerYear))
                {
                    SetSheetPrintFooter(sheet, footerYear);
                }
            }
            using (var outputStream = new ByteArrayOutputStream())
            {
                workbook.Write(outputStream);
                var byteArray = outputStream.ToByteArray();
                outputStream.Dispose();
                result = byteArray;
            }
            stream.Dispose();

            return result;
        }

        /// <summary>
        /// 在 【建筑材料工业生产者出厂价格指数】 sheet中定位年的数值
        /// </summary>
        /// <param name="sourceSheet"></param>
        /// <param name="yearStr"></param>
        /// <param name="findRowIndex"></param>
        /// <returns></returns>
        public static string FindCellAddressForBLD(ISheet sourceSheet, string yearStr, int findRowIndex)
        {
            string cellAddress = string.Empty;
            IRow sourceRow = sourceSheet.GetRow(findRowIndex);
            if (sourceRow != null)
            {
                for (int i = 1; i < sourceRow.LastCellNum; i++)
                {
                    ICell cell = sourceRow.GetCell(i);
                    if (cell != null && cell.ToString().Trim().StartsWith(yearStr))
                    {
                        cellAddress = $"{ToName(i)}{findRowIndex + 2}";
                    }
                }
            }
            return cellAddress;
        }

        /// <summary>
        /// 通过年，在【人工费-Wind】 Sheet中查找比例
        /// </summary>
        /// <param name="sourceSheet"></param>
        /// <param name="yearStr"></param>
        /// <param name="findRowIndex"></param>
        /// <returns></returns>
        public static string FindCellAddressForBLDWind(ISheet sourceSheet, string yearStr, int findRowIndex)
        {
            string cellAddress = string.Empty;
            for (int i = findRowIndex; i <= sourceSheet.LastRowNum; i++)
            {
                IRow row = sourceSheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(0);
                    if (cell != null && !string.IsNullOrEmpty(cell.ToString()) && yearStr.StartsWith(cell.ToString().Trim()))
                    {
                        cellAddress = $"C{i + 1}";
                    }
                }
            }
            return cellAddress;
        }

        /// <summary>
        /// 初始化sheet给定行号的位置，使之不为null，
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex">开始行，从0开始的行号下标</param>
        /// <param name="endRowIndex">结束行，从0开始的行号下标</param>
        public static void InitRowsCells(ISheet sheet, int startRowIndex, int endRowIndex = -1)
        {
            int realEndRowIndex = startRowIndex;
            if (endRowIndex > startRowIndex)
            {
                realEndRowIndex = endRowIndex;
            }

            for (int i = startRowIndex; i <= realEndRowIndex; i++)
            {
                //初始化行 和 单元格
                IRow thisRow = null;

                thisRow = sheet.GetRow(i);
                if (thisRow == null)
                {
                    thisRow = sheet.CreateRow(i);
                }
                for (int r = 0; r < 10; r++)
                {
                    ICell cell = thisRow.GetCell(r);
                    if (cell == null)
                    {
                        thisRow.CreateCell(r);
                    }
                }
            }
        }

        /// <summary>
        /// 获取工作簿中有效的sheet
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public static List<ISheet> GetVisibleSheet(XSSFWorkbook wb)
        {
            List<ISheet> sheets = new List<ISheet>();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                if (sheet != null && CheckSheetIsValid(wb, sheet))
                {
                    sheets.Add(sheet);
                }
            }

            return sheets;

        }
    }
}
