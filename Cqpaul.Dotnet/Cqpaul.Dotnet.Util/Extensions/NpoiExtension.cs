using Cqpaul.Dotnet.Util.Helpers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cqpaul.Dotnet.Util.Extensions
{
    public static class NpoiExtension
    {
        //扩展方法,根据sheet名称集合，得到sheet.
        //因为有时候模板sheet的名字改了，所以我们可以传输多个sheet名称，直到找到sheet为止。
        public static ISheet GetSheetByNameList(this IWorkbook workbook, params string[] sheetNameList)    //扩展方法
        {
            foreach (var sheetName in sheetNameList)
            {
                var sheet = workbook.GetSheet(sheetName);
                if (sheet != null)
                {
                    return sheet;
                }
            }

            return null;
        }

        public static void SetCellDatetimeValueIfNotNull(this ICell cell, string value)
        {
            if (!string.IsNullOrEmpty(value.ToString()))
            {
                cell.SetCellValue(DataValueFormatter.StringToExactDate(value).Date);
            }
        }

        public static void SetCellDoubleValueIfNotNull(this ICell cell, string value)
        {
            if (value == null) return;
            if (!string.IsNullOrEmpty(value.ToString()))
            {
                //if (cell.CellType == CellType.Formula && !string.IsNullOrEmpty(cell.CellFormula))
                //{
                //    cell.SetCellFormula(null);//既然要赋具体值，就去掉该cell的formula，以免冲突
                //}

                cell.SetCellValue(DataValueFormatter.CellDataValueToDouble(value));
            }
        }

        //扩展方法,根据sheet名称集合，得到sheet的位置index.
        //因为有时候模板sheet的名字改了，所以我们可以传输多个sheet名称，直到找到sheet为止。
        public static int GetSheetIndexByNameList(this IWorkbook workbook, params string[] sheetNameList)    //扩展方法
        {
            foreach (var sheetName in sheetNameList)
            {
                var sheetIndex = workbook.GetSheetIndex(sheetName);
                if (sheetIndex >= 0)
                {
                    return sheetIndex;
                }
            }

            return -1;
        }

        //扩展方法,得到cell的数字值
        public static double? CellNumberValue(this ICell cell, XSSFFormulaEvaluator evalor = null)    //扩展方法
        {
            return ExcelFunction.GetNumberValueFromCell(cell, evalor);
        }

        //扩展方法，给cell赋number类型的值
        public static void SetCellNumberValue(this ICell cell, double? cellNumberValue)    //扩展方法
        {
            if (cellNumberValue != null)
            {
                cell.SetCellValue(cellNumberValue.Value);
            }
        }

        //扩展方法，把某个cell的数字的值，赋给目标cell
        public static void SetCellNumberValueFromTargetCell(this ICell toCell, ICell fromCell)    //扩展方法
        {
            var cellNumberValue = fromCell.CellNumberValue();
            if (cellNumberValue != null)
            {
                toCell.SetCellValue(cellNumberValue.Value);
            }
        }

        public static void SetCellStringValueIfNotNull(this ICell cell, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                cell.SetCellValue(value.ToString());
            }
            if (cell == null)
            {

            }
        }


        /// <summary>
        /// 重载 CopyTo 方法，添加keepDataValidations选项
        /// </summary> 
        /// <param name="sheet"></param>
        /// <param name="dest"></param>
        /// <param name="name"></param>
        /// <param name="copyStyle"></param>
        /// <param name="keepFormulas"></param>
        /// <param name="keepDataValidations"></param>
        public static void CopyTo(this ISheet sheet, IWorkbook dest, string name, bool copyStyle = true, bool keepFormulas = true, bool keepDataValidations = false)
        {
            sheet.CopyTo(dest, name, copyStyle, keepFormulas);
            ISheet newSheet = dest.GetSheet(name);
            if (keepDataValidations)
            {
                // 需要从旧sheet中，取出之后，再添加。
                var oldValidations = sheet.GetDataValidations();
                oldValidations.ForEach(validation => newSheet.AddValidationData(validation));
            }

            //复制condition format
            for (int i = 0; i < sheet.SheetConditionalFormatting.NumConditionalFormattings; i++)
            {
                var conditionalFormatting = sheet.SheetConditionalFormatting.GetConditionalFormattingAt(i);
                newSheet.SheetConditionalFormatting.AddConditionalFormatting(conditionalFormatting);
            }

            var oldComments = sheet.GetCellComments();
            IDrawing patriarch = newSheet.CreateDrawingPatriarch();

            foreach (var comment in oldComments)
            {
                CopyComment(patriarch, newSheet, comment.Value);
            }

        }


        /// <summary>
        /// 复制批注
        /// </summary>
        /// <param name="patriarch"></param>
        /// <param name="sheet"></param>
        /// <param name="comment"></param>
        private static void CopyComment(IDrawing patriarch, ISheet sheet, IComment comment)
        {
            IComment newComment = patriarch.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, comment.Row, comment.Column, comment.Row + 2, comment.Column + 2));
            if (!string.IsNullOrEmpty(comment.Author))
            {
                newComment.Author = comment.Author;
            }
            newComment.Address = comment.Address;
            if (!string.IsNullOrEmpty(comment.String.ToString()))
            {
                newComment.String = comment.String;
                // newComment.String = new XSSFRichTextString($"{comment.String.ToString()}");
            }
            sheet.GetRow(comment.Row).GetCell(comment.Column).CellComment = newComment;
        }


        /// <summary>
        /// 测试能否删除多余的网格线
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dest"></param>
        /// <param name="name"></param>
        /// <param name="copyStyle"></param>
        /// <param name="keepFormulas"></param>
        /// <param name="keepDataValidations"></param>
        public static void CopyToUpdate(this ISheet sheet, IWorkbook dest, string name, bool copyStyle = true, bool keepFormulas = true, bool keepDataValidations = false)
        {
            sheet.CopyTo(dest, name, copyStyle, keepFormulas);
            ISheet newSheet = dest.GetSheet(name);
            if (keepDataValidations)
            {
                // 需要从旧sheet中，取出之后，再添加。
                var oldValidations = sheet.GetDataValidations();
                oldValidations.ForEach(validation => newSheet.AddValidationData(validation));
            }

            ICellStyle noBorderCellStyle = dest.CreateCellStyle();
            noBorderCellStyle.BorderTop = BorderStyle.None;
            noBorderCellStyle.BorderBottom = BorderStyle.None;
            noBorderCellStyle.BorderLeft = BorderStyle.None;
            noBorderCellStyle.BorderRight = BorderStyle.None;

            #region 测试设置无边框
            for (int i = 0; i <= newSheet.LastRowNum; i++)
            {
                IRow sheetRow = newSheet.GetRow(i);
                if (sheetRow != null)
                {
                    for (int j = 0; j < sheetRow.LastCellNum; j++)
                    {
                        ICell cell = sheetRow.GetCell(j);
                        if (cell != null)
                        {
                            cell.CellStyle = noBorderCellStyle;
                        }
                    }
                }
            }
            #endregion
        }

        /// <summary>
        /// Sheet中，按某列为模板，添加新列
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="templateColumn"></param>
        /// <param name="copyStartCellIndex"></param>
        /// <param name="copyEndCellIndex"></param>
        /// <param name="copyColumnCount"></param>
        public static void AddNewColumnByCopyTemplateCells(this ISheet sheet, int templateColumn, int copyStartCellIndex, int copyEndCellIndex, int copyColumnCount)
        {
            for (int i = copyStartCellIndex; i <= copyEndCellIndex; i++)
            {
                IRow dataRow = sheet.GetRow(i);
                for (int j = 1; j <= copyColumnCount; j++)
                {
                    var newCell = dataRow.CreateCell(templateColumn + j);
                    //var templateCell = dataRow.GetCell(templateColumn);
                    dataRow.CopyCell(templateColumn, templateColumn + j);
                }
            }
        }

        public static void SetCellAdaptiveTextSize(this ICell cell, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                cell.SetCellValue(value);
                var cellStyle = cell.CellStyle;
                cellStyle.ShrinkToFit = true;
            }
        }

        public static void SetCellAdaptiveTextSize(this ICell cell, double value)
        {
            cell.SetCellValue(value);
            var cellStyle = cell.CellStyle;
            cellStyle.ShrinkToFit = true;
        }

        public static void SetCellAdaptiveTextSize(this ICell cell, double? value)
        {
            if (value == null)
            {
                value = 0;
            }
            cell.SetCellValue(value.Value);
            var cellStyle = cell.CellStyle;
            cellStyle.ShrinkToFit = true;
        }

        public static void SetCellAdaptiveTextSize(this ICell cell, DateTime? value)
        {
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
            }
            else
            {
                cell.SetCellValue((DateTime)value);
            }
            var cellStyle = cell.CellStyle;
            cellStyle.ShrinkToFit = true;
        }

        /// <summary>
        /// 合并单元格，如果指定合并单元格已经合并，则不合并
        /// </summary>
        /// <param name="sheet">工作表单</param>
        /// <param name="cellRangeAddress">合并范围</param>
        public static void AddMergedRegionSmart(this ISheet sheet, CellRangeAddress cellRangeAddress)
        {
            if (cellRangeAddress.FirstRow == cellRangeAddress.LastRow
                && cellRangeAddress.FirstColumn == cellRangeAddress.LastColumn)
            {
                //单一单元格，不能合并
                return;
            }

            if (!sheet.IsMergedRegion(cellRangeAddress))
            {
                sheet.AddMergedRegion(cellRangeAddress);
            }

        }

        /// <summary>
        /// 取得Excel实际的行号，从1开始。
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static int DisplayRowNum(this IRow row)
        {
            //应为RowNum是0 based的。
            return row.RowNum + 1;
        }

        public static ICellStyle SetBackGroundColor(this ICellStyle cellStyle, IWorkbook workbook)
        {
            XSSFCellStyle xSSFCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            ReflectFunction.CopyPropertys(cellStyle, xSSFCellStyle);
            XSSFColor xssfColor = new XSSFColor();
            //根据自己需要设置RGB
            byte[] colorRgb = { 226, 239, 218 };
            xssfColor.SetRgb(colorRgb);
            xSSFCellStyle.FillForegroundColorColor = xssfColor;
            xSSFCellStyle.FillPattern = FillPattern.SolidForeground;

            return xSSFCellStyle;
        }

        public static List<string> GetVisibleSheetsName(this IWorkbook workbook)
        {
            List<string> sheetNames = new List<string>();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                if (!workbook.IsSheetHidden(i) && !workbook.IsSheetVeryHidden(i))
                {
                    var sheet = workbook.GetSheetAt(i);
                    sheetNames.Add(sheet.SheetName);
                }
            }
            return sheetNames;
        }

        /// <summary>
        /// 在sheet指定位置插入行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="rowToAdd"></param>
        public static void InsertRow(this ISheet sheet, int startRowIndex, int rowToAdd)
        {
            if (sheet != null && sheet.LastRowNum > startRowIndex)
            {
                //startRow,endRow,startCol,endCol
                List<(int, int, int, int)> mergedCells = new List<(int, int, int, int)>();
                foreach (var merged in sheet.MergedRegions)
                {
                    //记住插入后后面的行的merge情况
                    if (merged.FirstRow >= startRowIndex)
                    {
                        mergedCells.Add((merged.FirstRow, merged.LastRow, merged.FirstColumn, merged.LastColumn));
                    }
                }
                sheet.ShiftRows(startRowIndex, sheet.LastRowNum, rowToAdd, true, false);
                for (int i = startRowIndex; i < startRowIndex + rowToAdd; i++)
                {
                    sheet.CreateRow(i);
                }

                foreach (var merged in mergedCells)
                {
                    sheet.AddMergedRegionSmart(new CellRangeAddress(merged.Item1 + rowToAdd, merged.Item2 + rowToAdd, merged.Item3, merged.Item4));
                }
            }
            else
            {
                throw new Exception("插入位置超过了最大行数");
            }
        }

        public static void DeleteRow(this ISheet sheet, int rowIndex)
        {
            if (rowIndex <= sheet.LastRowNum)
            {
                //如果要删除的是最后一行，无法进行shift，只能用remove来置空
                if (rowIndex == sheet.LastRowNum)
                {
                    IRow toRemoveRow = sheet.GetRow(rowIndex);
                    if (toRemoveRow != null)
                    {
                        sheet.RemoveRow(toRemoveRow);
                    }
                }
                else
                {
                    //startRow,endRow,startCol,endCol
                    List<(int, int, int, int)> mergedCells = new List<(int, int, int, int)>();
                    foreach (var merged in sheet.MergedRegions)
                    {
                        //记住插入后后面的行的merge情况
                        if (merged.FirstRow >= rowIndex)
                        {
                            mergedCells.Add((merged.FirstRow, merged.LastRow, merged.FirstColumn, merged.LastColumn));
                        }
                    }
                    sheet.ShiftRows(rowIndex + 1, sheet.LastRowNum, -1, true, false);
                    foreach (var merged in mergedCells)
                    {
                        sheet.AddMergedRegionSmart(new CellRangeAddress(merged.Item1 - 1, merged.Item2 - 1, merged.Item3, merged.Item4));
                    }
                }
            }
        }

        public static void CopyFromRow(this IRow row, int sourceRowIndex)
        {
            IRow sourceRow = row.Sheet.GetRow(sourceRowIndex);
            row.CopyFromRow(sourceRow);
        }

        public static void CopyFromRow(this IRow row, IRow sourceRow)
        {
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                ICell sourceCell = sourceRow.GetCell(i);
                if (sourceCell != null)
                {
                    ICell destCell = row.GetCell(i);
                    if (destCell == null)
                    {
                        destCell = row.CreateCell(i);
                    }
                    if (sourceCell.IsMergedCell)
                    {
                        var mergedRegion = sourceCell.Sheet.MergedRegions.FirstOrDefault(x => x.FirstRow == sourceCell.RowIndex && x.FirstColumn == sourceCell.ColumnIndex);
                        if (mergedRegion != null)
                        {
                            int mergedRows = mergedRegion.LastRow - sourceCell.RowIndex;
                            int mergedCols = mergedRegion.LastColumn - sourceCell.ColumnIndex;
                            destCell.Sheet.AddMergedRegionSmart(new CellRangeAddress(destCell.RowIndex, destCell.RowIndex + mergedRows, destCell.ColumnIndex, destCell.ColumnIndex + mergedCols));
                        }
                    }
                    destCell.CellStyle = sourceCell.CellStyle;
                    destCell.SetCellValue(sourceCell.StringCellValue);
                }
            }
            row.Height = sourceRow.Height;

        }

        /// <summary>
        /// 将申报表分页后。申报表限定！！
        /// </summary>
        /// <param name="originalSheet"></param>
        /// <param name="workbook"></param>
        /// <param name="headerCount"></param>
        /// <param name="footerCount"></param>
        public static void SplitPagedSheet(this ISheet sheet)
        {
            if (sheet != null && sheet.SheetName.Contains("申报表"))
            {
                Regex regex = new Regex(@"\d{4}\s*年");

                //清理空行,从尾到头，避免删行出现问题
                int originalLastRow = sheet.LastRowNum;
                for (int i = originalLastRow; i > 6; i--)
                {
                    bool empltyRow = true;
                    IRow thisRow = sheet.GetRow(i);
                    if (thisRow != null)
                    {
                        for (int c = 0; c <= thisRow.LastCellNum; c++)
                        {
                            ICell thisCell = thisRow.GetCell(c);
                            if (thisCell != null && !string.IsNullOrWhiteSpace(thisCell.ToString()))
                            {
                                empltyRow = false; break;
                            }
                        }
                    }

                    if (empltyRow)
                    {
                        sheet.DeleteRow(i);
                    }
                }

                string tableFilledBy = string.Empty;
                string tableFilledAt = string.Empty;
                ICellStyle lastCellStyle = null;
                //得到真实的最后一行
                int actuallyLastRow = 0;

                //得到页脚信息
                for (int i = sheet.LastRowNum; i > 0; i--)
                {
                    IRow loopedRow = sheet.GetRow(i);
                    if (loopedRow == null) continue;

                    ICell cell = loopedRow.GetCell(0);
                    if (cell == null) continue;

                    if (cell.StringCellValue.Contains("填表日期"))
                    {
                        tableFilledBy = sheet.GetRow(i - 1).GetCell(0).StringCellValue;
                        tableFilledAt = cell.StringCellValue;
                        actuallyLastRow = i;
                        lastCellStyle = cell.CellStyle;
                        break;
                    }
                }

                //获得评估基准日年份
                string year = regex.Match(sheet.GetRow(1).GetCell(0).StringCellValue)?.Value.Replace("年", "").Trim();

                //预处理被评估单位,填表人,填表时间
                sheet.AddMergedRegionSmart(new CellRangeAddress(3, 3, 0, sheet.GetRow(3).LastCellNum / 2));
                sheet.AddMergedRegionSmart(new CellRangeAddress(actuallyLastRow - 1, actuallyLastRow - 1, 0, sheet.GetRow(3).LastCellNum / 2));
                sheet.AddMergedRegionSmart(new CellRangeAddress(actuallyLastRow, actuallyLastRow, 0, sheet.GetRow(3).LastCellNum / 2));

                //真实最后一行 与 文档最后一行的 差距
                int lastRowGap = sheet.LastRowNum - actuallyLastRow;

                //Add paging info row
                sheet.InsertRow(3, 1);
                //获得最后一列的位置
                int tableLastColumn = sheet.MergedRegions.Single(x => x.FirstRow == 0 && x.FirstColumn == 0).LastColumn;
                sheet.AddMergedRegionSmart(new CellRangeAddress(3, 3, 0, tableLastColumn));
                ICell row3Cell = sheet.GetRow(3).CreateCell(0);
                row3Cell.SetCellValue("共 ** 页第 * 页");
                row3Cell.CellStyle = sheet.GetRow(2).GetCell(0).CellStyle;

                //7是起始行，
                for (int i = 7; i < sheet.LastRowNum - lastRowGap; i++)
                {
                    //每页24行
                    //判断是需要分页，
                    //   需要分页判断依据，除此之外就不分页：
                    //      1 - 理论上该页的最后一行比最后一行要小
                    if (i + 23 < sheet.LastRowNum - lastRowGap)
                    {
                        //分页
                        //至少保证最后涉及合计（含页脚）的行在一起，约占6行
                        int addrow = 23;
                        int rowGap = sheet.LastRowNum - lastRowGap - (i + 23);
                        if (rowGap < 6)
                        {
                            addrow = addrow - (6 - rowGap);
                        }
                        i += addrow;

                        //加新行
                        sheet.InsertRow(i, 9);

                        //填充新行
                        //当前页页脚
                        sheet.GetRow(i).CopyFromRow(sheet.LastRowNum - lastRowGap - 1);
                        sheet.GetRow(i + 1).CopyFromRow(sheet.LastRowNum - lastRowGap);
                        i += 2;

                        //分页符
                        sheet.SetRowBreak(i - 1);

                        //下一页页头
                        for (int ii = 0; ii < 7; ii++)
                        {
                            sheet.GetRow(i + ii).CopyFromRow(0 + ii);
                        }

                        i += 7;
                    }
                }

                //设置页号
                List<int> pageInfoRows = new List<int>();
                for (int i = 0; i < sheet.LastRowNum - lastRowGap; i++)
                {
                    IRow loopedRow = sheet.GetRow(i);
                    if (loopedRow == null) continue;

                    ICell cell = loopedRow.GetCell(0);
                    if (cell.ToString().Contains("共 ** 页第 * 页"))
                    {
                        pageInfoRows.Add(i);
                    }
                }

                for (int p = 1; p <= pageInfoRows.Count; p++)
                {
                    int pageInfoRow = pageInfoRows[p - 1];
                    sheet.GetRow(pageInfoRow).GetCell(0).SetCellValue("共 ** 页第 * 页".Replace("**", pageInfoRows.Count.ToString()).Replace("*", p.ToString()));
                }

                //设置打印相关格式
                sheet.PrintSetup.NoColor = true;
                sheet.PrintSetup.Landscape = true;
                sheet.FitToPage = true;
                sheet.PrintSetup.FitWidth = 1;
                sheet.PrintSetup.FitHeight = 0;
                //ExcelFunction.SetSheetPrintFooter(sheet, year);
            }
        }
    }
}
