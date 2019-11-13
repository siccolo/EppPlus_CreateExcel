using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NET_Excel.Models
{
    public interface IMergeInfo
    {
        int FromRow { get;  }
        string FromColumn { get;   }

        int ToRow { get;   }
        string ToColumn { get;   }

        string RangeName { get; }

        OfficeOpenXml.Style.ExcelHorizontalAlignment HorizontalAlignment { get; }
        OfficeOpenXml.Style.ExcelVerticalAlignment VerticalAlignment { get; }
    }
    public sealed class MergeInfo : IMergeInfo
    {
        public int FromRow { get; private set; }
        public string FromColumn { get; private set; }

        public int ToRow { get; private set; }
        public string ToColumn { get; private set; }

        public OfficeOpenXml.Style.ExcelHorizontalAlignment HorizontalAlignment { get; private set; }
        public OfficeOpenXml.Style.ExcelVerticalAlignment VerticalAlignment { get; private set; }

        public string RangeName
        {
            get
            {
                return String.Format("{0}{1}:{2}{3}"
                        , FromColumn, FromRow.ToString()
                        , ToColumn, ToRow.ToString()
                        );
            }
        }

        public MergeInfo(int fromRow
                , string fromColumn
                , int toRow
                , string toColumn
                
                //xlHAlignCenter	-4108	Center.
                , OfficeOpenXml.Style.ExcelHorizontalAlignment hAligment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
                //xlVAlignBottom	-4107	Bottom.
                , OfficeOpenXml.Style.ExcelVerticalAlignment vAligment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom
            ) 
        {
            FromRow = fromRow;
            FromColumn = fromColumn;
            ToRow = toRow;
            ToColumn = toColumn;
            HorizontalAlignment = hAligment ;
            VerticalAlignment = vAligment;
        }
    }
}
