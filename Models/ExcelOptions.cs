using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NET_Excel.Models
{
    public interface IExcelOptions {
        string FileName { get;  }
        string WorksheetName { get;  }

        string WorksheetTitle { get;  }
        System.Drawing.Font TitleFont { get; set; }
        System.Drawing.Font DefaultWorksheetFont { get; set; }

        IEnumerable<IExcelColumn> ColumnList { get;   }

        IEnumerable<IMergeInfo> MergeInfoList { get; }
        OfficeOpenXml.Style.ExcelHorizontalAlignment DataHorizontalAlignment { get; }
        OfficeOpenXml.Style.ExcelVerticalAlignment DataVerticalAlignment { get; }

        bool BorderAroundColumnHeaders { get; set; }
        bool SetColumnHeadersBackgroundColor { get; set; }
        System.Drawing.Color ColumnHeadersBackgroundColor { get; set; }
        System.Drawing.Color ColumnHeadersTextColor { get; set; }
    }

    public sealed class ExcelOptions : IExcelOptions
    {

        public const int HorizontalAligmentCenter =2;// (int)OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        public const int VerticalAligmentCenter = 2;//(int)OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        public string FileName { get; private set; }
        
        public string WorksheetName { get; private set; }

        //.ActiveCell.FormulaR1C1 = strHeader & " מתאריך " & strFromDate & " עד תאריך " & DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(strToDate))
        public string WorksheetTitle{ get; private set; }
        public System.Drawing.Font TitleFont { get; set; }

        public System.Drawing.Font DefaultWorksheetFont { get; set; }

        public IEnumerable<IExcelColumn> ColumnList { get; private set; }
        public bool BorderAroundColumnHeaders { get; set; }//=false;
        public bool SetColumnHeadersBackgroundColor { get; set; }//=false;
        public System.Drawing.Color ColumnHeadersBackgroundColor { get; set; }
        public System.Drawing.Color ColumnHeadersTextColor { get; set; }

        public OfficeOpenXml.Style.ExcelHorizontalAlignment DataHorizontalAlignment { get; private set; }
        public OfficeOpenXml.Style.ExcelVerticalAlignment DataVerticalAlignment { get; private set; }

        //public ExcelOptions()
        //{
         //   FileName = "";
          //  WorksheetName = "";
       // }
        public ExcelOptions(string fileName
                                , string worksheetName
                                ,  string worksheetTitle
                                , IEnumerable<IExcelColumn> columns
                                //xlCenter
                                , OfficeOpenXml.Style.ExcelHorizontalAlignment dataHorizontalAligment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
                                //xlCenter
                                , OfficeOpenXml.Style.ExcelVerticalAlignment dataVerticalAligment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center
                                , IEnumerable<IMergeInfo> merges = null)
        {
            FileName = fileName;
            WorksheetName = worksheetName;
            WorksheetTitle = worksheetTitle;
            ColumnList = columns;
            DataHorizontalAlignment = dataHorizontalAligment;
            DataVerticalAlignment = dataVerticalAligment;
            MergeInfoList = merges;

            BorderAroundColumnHeaders = false;
            SetColumnHeadersBackgroundColor = false;
            //colors at http://dmcritchie.mvps.org/excel/colors.htm
            ColumnHeadersBackgroundColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
            ColumnHeadersTextColor = System.Drawing.ColorTranslator.FromHtml("#000000");

            //default title font
            TitleFont = new System.Drawing.Font("Tahoma", 10, System.Drawing.FontStyle.Bold);
            DefaultWorksheetFont = new System.Drawing.Font("Tahoma", 8, System.Drawing.FontStyle.Regular);
        }

        public IEnumerable<IMergeInfo> MergeInfoList { get; private set; }
        /*
        public ExcelOptions(string fileName, string worksheetName, string worksheetTitle, IEnumerable<IExcelColumn> columns, IEnumerable<IMergeInfo> merges)
            : this(fileName, worksheetName, worksheetTitle, columns)
        {
            MergeInfoList = merges;
        }
        */
    }
}
