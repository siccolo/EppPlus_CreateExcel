using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NET_Excel.Models
{
    public interface IExcelColumn
    {
        string ColumnName { get;  }
        int ColumnWidth { get;  }
    }

    public sealed class ExcelColumn : IExcelColumn
    {
        public string ColumnName { get;private set;}
        public int ColumnWidth { get; private set; }

        public ExcelColumn(string name, int width)
        {
            ColumnName = name;
            ColumnWidth = width;
        }
    }
}
