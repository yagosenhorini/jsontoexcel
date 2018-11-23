using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BasicInfo
{
    class TaxRegime
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
        public string TaxRegimeExcel(List<string> list)
        {

            var indexs = MapTaxRegimeToExcel(list);


            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open("D:\\Users\\staff\\Desktop\\Book1.xlsx");
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Sheet3"];

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                oSheet.Cells[i + 1, countParent++] = ((string)obj["federalTaxNumber"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["name"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["taxRegime"] ?? "");
                oSheet.Cells[i + 1, countParent++] = ((string)obj["optedInOn"] ?? "");



                if ((obj["previousDetails"]) != null)
                {
                    var couPhone = countParent;
                    var x = 0;
                    for (x = 0; x < obj["previousDetails"].Count(); x++)
                    {
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["previousDetails"][x]["endOn"] ?? "");
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["previousDetails"][x]["beginOn"] ?? "");
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["previousDetails"][x]["description"] ?? "");
                        oSheet.Cells[i + 1, couPhone++] = ((string)obj["previousDetails"][x]["taxRegime"] ?? "");

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

        public Tuple<int, int> MapTaxRegimeToExcel(List<string> list)
        {
            var coEC = 0;
            var coPh = 0;

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                Cells(i + 1, countParent++, ((string)obj["federalTaxNumber"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["name"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["taxRegime"] ?? ""));
                Cells(i + 1, countParent++, ((string)obj["optedInOn"] ?? ""));


                if ((obj["previousDetails"]) != null)
                {
                    var couPhone = countParent;
                    var x = 0;
                    for (x = 0; x < obj["previousDetails"].Count(); x++)
                    {
                        Cells(i + 1, couPhone++, ((string)obj["previousDetails"][x]["endOn"] ?? ""));
                        Cells(i + 1, couPhone++, ((string)obj["previousDetails"][x]["beginOn"] ?? ""));
                        Cells(i + 1, couPhone++, ((string)obj["previousDetails"][x]["description"] ?? ""));
                        Cells(i + 1, couPhone++, ((string)obj["previousDetails"][x]["taxRegime"] ?? ""));

                    }

                    coPh = (couPhone > coPh) ? couPhone : coPh;
                    countParent = couPhone;
                }
            }

            return new Tuple<int, int>(coEC, coPh);
        }
    }
}
