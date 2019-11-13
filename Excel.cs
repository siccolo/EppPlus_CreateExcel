using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//
using System.IO;
using OfficeOpenXml;
//using System.Xml;
//using System.Drawing;
using OfficeOpenXml.Style;
//
using NET_Excel.Helpers;
using NET_Excel.Models;

namespace NET_Excel
{
    public class Excel
    {
        private readonly Models.IExcelOptions _ExcelOptions;
        public Excel(Models.IExcelOptions options)
        {
            if (options == null)
            {
                throw new System.ArgumentNullException("options are mssing");
            }
            if (String.IsNullOrEmpty(options.FileName))
            {
                throw new System.ArgumentNullException("filename is mssing");
            }
            if (options.ColumnList == null || !options.ColumnList.Any())
            {
                throw new System.ArgumentNullException("columns are mssing");
            }
            //
            _ExcelOptions = options;//?? throw new System.ArgumentNullException("options are mssing");
        }

        public Models.IResult<bool> CreateExcelReport(System.Data.DataTable dtData)
        {
            if (dtData == null)
            {
                return new Result<bool>( new System.ArgumentNullException("dtData"));
            }

            using (var package = new ExcelPackage())
            {
                int rowStart = 2;
                int columnEnd = dtData.Columns.Count;
                //.Workbooks.Add()
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(_ExcelOptions.WorksheetName);

                try { 
                    rowStart = InitWorksheet(worksheet, rowStart, columnEnd);
                    AddPageSetup(worksheet, rowStart);

                    //load data:
                    //expression.CopyFromRecordset (Data, MaxRows, MaxColumns)
                    rowStart++;
                    AddData(worksheet, dtData, rowStart);

                    worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Font.SetFromFont(_ExcelOptions.DefaultWorksheetFont);
                    worksheet.Cells[2, 1].Style.Font.SetFromFont(_ExcelOptions.TitleFont);

                    //freeze
                    worksheet.View.FreezePanes(rowStart, columnEnd);

                    var file = UtilExtensions.GetFileInfo(_ExcelOptions.FileName);
                    package.SaveAs(file);
                }
                catch (Exception ex)
                {
                    return new Result<bool>(ex);
                }
            }

            return new Result<bool>(true);
        }

        /*public bool CreateExcelReport(ADODB.Recordset rsData)
        {
            using (var package = new ExcelPackage())
            {
                int rowStart = 2;
                int columnEnd = rsData.Fields.Count;
                //.Workbooks.Add()
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(_ExcelOptions.WorksheetName);
                InitWorksheet(worksheet, rowStart, columnEnd);

                //load data:
                //expression.CopyFromRecordset (Data, MaxRows, MaxColumns)
                rowStart = 4;
                AddData(worksheet, rsData, rowStart);
                var file = UtilExtensions.GetFileInfo("sample1.xlsx");
                package.SaveAs(file);
            }

            return true;
        }*/

        private int InitWorksheet(ExcelWorksheet worksheet, int startRow, int endColumn)
        {
            //var xlFile = Utils.GetFileInfo("sample1.xlsx");
            System.IO.File.Delete(_ExcelOptions.FileName );

            worksheet.View.RightToLeft = true;
            
            //worksheet.InsertRow(1, 1);
            //worksheet.InsertColumn(1, endColumn);

            //add worksheet title:
            //gf_MergeCells(eApp, 2, 1, 25) ' 19
            if (!String.IsNullOrEmpty(_ExcelOptions.WorksheetTitle))
            {
                //worksheet.InsertRow(2, 3);
                AddWorksheetTitle(worksheet, startRow, endColumn);
                startRow = 4;
            }

            //
            AddColumns(worksheet, startRow);
            
            //gf_MergeCells(ByRef objExcel As Object, ByVal lngLine As Integer, ByVal lngFromCol As Integer, ByVal lngToCol As Integer)
            //gf_MergeCells(eApp, 2, 1, 25) ' 19
            AddMerges(worksheet);

            return startRow;
        }

        private void AddColumns(ExcelWorksheet worksheet, int startRow)
        {
            //
            var columns = _ExcelOptions.ColumnList.ToList();
            for (int i = 1; i <= columns.Count(); i++)
            {
                worksheet.Cells[startRow, i].Value = columns[i - 1].ColumnName;
                worksheet.Column(i).Width = columns[i - 1].ColumnWidth;
            }
            //
            //border around headers
            //if (_ExcelOptions.BorderAroundColumnHeaders || _ExcelOptions.SetColumnHeadersBackgroundColor)
            Models.MergeInfo mergeBorder = new Models.MergeInfo(startRow, "A", startRow, columns.Count().ToColumnLetter());
            using (ExcelRange r = worksheet.Cells[mergeBorder.RangeName])
            {
                r.AutoFilter = true;
                if (_ExcelOptions.SetColumnHeadersBackgroundColor)
                {
                    r.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                    r.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    r.Style.Border.Left.Style = ExcelBorderStyle.Medium;
                }
                //.Selection.Interior.ColorIndex = 47
                //.Selection.Font.ColorIndex = 2
                if (_ExcelOptions.SetColumnHeadersBackgroundColor)
                {
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#666699"));
                    r.Style.Fill.BackgroundColor.SetColor(_ExcelOptions.ColumnHeadersBackgroundColor);

                    //r.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"));
                    r.Style.Font.Color.SetColor(_ExcelOptions.ColumnHeadersTextColor);
                }
            }
        }

        private void AddMerges(ExcelWorksheet worksheet)
        {
            if (_ExcelOptions.MergeInfoList == null || !_ExcelOptions.MergeInfoList.Any())
            {
                return;
            }
            foreach (var merge in _ExcelOptions.MergeInfoList)
            {
                AddMerge(worksheet, merge);
            }
            //
        }

        private void AddMerge(ExcelWorksheet worksheet, Models.IMergeInfo merge)
        {
            using (ExcelRange r = worksheet.Cells[merge.RangeName])
            {
                r.Merge = true;
                r.Style.WrapText = false;
                //OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                r.Style.HorizontalAlignment = merge.HorizontalAlignment;
                //OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                r.Style.VerticalAlignment = merge.VerticalAlignment;
                //
                r.Style.ReadingOrder = ExcelReadingOrder.ContextDependent;
            }
        }

        private void AddWorksheetTitle(ExcelWorksheet worksheet, int startRow, int endColumn)
        {
             Models.MergeInfo merge = new Models.MergeInfo(2, "A", 2, 25.ToColumnLetter() );
             AddMerge(worksheet, merge);
             worksheet.Cells[startRow, 1].Value = _ExcelOptions.WorksheetTitle;
             //worksheet.Cells[merge.RangeName].Style.Font.Bold = true;
             worksheet.Cells[startRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
             worksheet.Cells[startRow, 1].Style.Font.SetFromFont(_ExcelOptions.TitleFont);
             worksheet.Cells[startRow, 1].Style.ReadingOrder = ExcelReadingOrder.ContextDependent;
        }

        private void AddData(ExcelWorksheet worksheet, System.Data.DataTable data, int startRow)
        {
            //only load upto provided column:
            // .Range("A" & lngRow).CopyFromRecordset(ADORs, , lngLastCol)
            using (var dtCopy = data.Copy())
            {
                var columnsCount = _ExcelOptions.ColumnList.ToList().Count();
                while (dtCopy.Columns.Count > columnsCount)
                {
                    dtCopy.Columns.RemoveAt(columnsCount);
                }
                worksheet.Cells[startRow, 1].LoadFromDataTable(dtCopy, false);
                //check for datetime:
                var ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                for (int i = 0; i < dtCopy.Columns.Count; i++)
                {
                    if (dtCopy.Columns[i].DataType == typeof(DateTime))
                    {
                        worksheet.Column(i+1).Style.Numberformat.Format = ci.DateTimeFormat.ShortDatePattern;
                    }
                }
            }
        }

        private void AddPageSetup(ExcelWorksheet worksheet, int headerRow)
        {
            worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

            var repeatedRows = String.Format ("${0}:${0}", headerRow.ToString());
            worksheet.PrinterSettings.RepeatRows = new ExcelAddress(repeatedRows);
            
            var footer =  "עמ' &P  מתוך &N";
            worksheet.HeaderFooter.EvenFooter.CenteredText = footer;
            worksheet.HeaderFooter.OddFooter.CenteredText = footer;
            
            worksheet.PrinterSettings.HorizontalCentered = true;
            worksheet.PrinterSettings.FitToPage = true;
            worksheet.PrinterSettings.FitToWidth = 1;
            worksheet.PrinterSettings.FitToHeight = 0;
        }

        /*private void AddData(ExcelWorksheet worksheet, ADODB.Recordset rsData, int startRow)
        {
            worksheet.Cells[startRow, 1].LoadFromArrays (rsData.GetRows() );
        }*/
    }
}
