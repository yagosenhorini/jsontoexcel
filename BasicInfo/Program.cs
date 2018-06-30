using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicInfo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var js = new JsonToCsv();
            var rd = js.ReadDirectory();
            var tj = js.TransformToString(rd);
            var sj = js.GetMapExcel(tj);

        }
    }
}
