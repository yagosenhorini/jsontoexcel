using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BasicInfo
{
    class StateTax
    {
        public string[] ReadDirectory() => Directory.GetFiles("D:\\Users\\staff\\Desktop\\teste\\");

        public List<string> TransformToString(string[] dir)
        {
            var list = new List<string>();

            for (var i = 0; i < dir.Length; i++)
            {
                var dataJson = File.ReadAllText(dir[i]);
                Console.WriteLine("Line " + i);
                list.Add(dataJson);
            }

            return list;
        }
        public string StateTaxExcel(List<string> list)
        {

            var indexs = MapStateTaxToExcel(list);


            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open("D:\\Users\\staff\\Desktop\\Book1.xlsx");
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Sheet2"];

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                oSheet.Cells[i + 1, countParent++] = ((string)obj["federalTaxNumber"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["createdOnn"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["name"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["tradeName"] ?? "");

                if ((obj["taxPayer"]) != null)
                {
                    var couPhone = countParent;
                    var x = 0;
                    for (x = 0; x < obj["taxPayer"].Count(); x++)
                    {
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["taxPayer"][x]["state"]["abbreviation"] ?? "");
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["taxPayer"][x]["stateTaxNumber"] ?? "");
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["taxPayer"][x]["statusStateTax"] ?? "");
                    }
                }

                countParent = indexs.Item2;
            }

            oWB.Save();

            if (oWB != null)
                oWB.Close();

            return null;
        }
        public void Cells(int c, int d, string f)
        {

        }

        public Tuple<int, int> MapStateTaxToExcel(List<string> list)
        {
            var coEC = 0;
            var coPh = 0;

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                Cells(i + 1, countParent++, ((string)obj["federalTaxNumber"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["createdOn"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["name"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["tradeName"] ?? ""));

                if ((obj["taxPayer"]) != null)
                {
                    var couPhone = countParent;
                    var x = 0;
                    for (x = 0; x < obj["taxPayer"].Count(); x++)
                    {
                        Cells(i + 1, couPhone++, ((string)obj["taxPayer"][x]["state"]["abbreviation"] ?? ""));
                        Cells(i + 1, couPhone++, ((string)obj["taxPayer"][x]["stateTaxNumber"] ?? ""));
                        Cells(i + 1, couPhone++, ((string)obj["taxPayer"][x]["statusStateTax"] ?? ""));
                    }

                    coPh = (couPhone > coPh) ? couPhone : coPh;
                    countParent = couPhone;
                }
            }

            return new Tuple<int, int>(coEC, coPh);
        }
    }
}
