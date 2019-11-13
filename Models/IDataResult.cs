using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NET_Excel.Models
{
    public interface IResult<T>
    {
        bool Success { get; }
        Exception Error { get; }
        T Data { get; }
    }

    public class Result<T> : IResult<T>
    {
        public bool Success { get; private set; }
        public Exception Error { get; private set; }
        public T Data { get; private set; }

        public Result(Exception error)
        {
            Success = false;
            Error = error;
        }

        public Result(T data)
        {
            Success = true;
            Data = data;
        }
    }
}
