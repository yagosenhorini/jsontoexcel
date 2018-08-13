using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BasicInfo
{
    public class JsonToCsv
    {
        public string[] ReadDirectory() => Directory.GetFiles("D:\\Users\\staff\\Desktop\\cpf\\");

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

        public string GetMapExcel(List<string> list)
        {

            //var indexs = MapToExcel(list);


            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open("D:\\Users\\staff\\Desktop\\cpf_gabriel.xlsx");
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Sheet2"];

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                //oSheet.Cells[i + 1, countParent++] = ((string)obj["federalTaxNumber"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["openedOn"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["name"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["tradeName"] ?? "");

                if ((obj["relationships"]) != null)
                {
                    var couRelation = 6;
                    var k = 0;
                    for (k = 0; k < obj["relationships"].Count(); k++)
                    {
                        oSheet.Cells[i + 1, couRelation++] = ((string)obj["relationships"][k]["name"] ?? "");
                        oSheet.Cells[i + 1, couRelation++] = ((string)obj["relationships"][k]["relationType"] ?? "");
                        oSheet.Cells[i + 1, couRelation++] = ((string)obj["relationships"][k]["gender"] ?? "");
                        oSheet.Cells[i + 1, couRelation++] = ((string)obj["relationships"][k]["federalTaxNumber"] ?? "");
                        oSheet.Cells[i + 1, couRelation++] = ((string)obj["relationships"][k]["status"] ?? "");
                    }
                }

                //countParent = indexs.Item1;

                //if ((obj["address"]) != null)
                //{
                //    var couRelationship = 7;
                //    var x = 0;
                //    for (x = 0; x < obj["address"].Count(); x++)
                //    {
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["street"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["number"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["additionalInformation"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["postalCode"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["district"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["city"]["name"] ?? "");
                //        oSheet.Cells[i + 1, couRelationship++] = ((string)obj["address"][x]["state"] ?? "");
                //    }
                //}

                //oSheet.Cells[i + 1, countParent++] = ((string)obj["name"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["birthOn"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["federalTaxNumber"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["gender"] ?? "");
                //oSheet.Cells[i + 1, countParent++] = ((string)obj["status"] ?? "");

            }

            oWB.Save();

            if (oWB != null)
                oWB.Close();

            return null;
        }

        public void Cells(int c, int d, string f)
        {

        }

        public Tuple<int, int> MapToExcel(List<string> list)
        {
            var coEC = 0;
            var coPh = 0;

            for (var i = 1; i <= list.Count; i++)
            {
                var countParent = 1;
                JObject obj = JObject.Parse(list[i - 1]);

                //Cells(i + 1, countParent++, ((string)obj["federalTaxNumber"] ?? ""));
                //Cells(i + 1, countParent++, ((string)obj["openedOn"] ?? ""));
                //Cells(i + 1, countParent++, ((string)obj["name"] ?? ""));
                //Cells(i + 1, countParent++, ((string)obj["tradeName"] ?? ""));

            //    if ((obj["relationships"]) != null)
            //    {
            //        var couCode = 5;
            //        var k = 0;
            //        for (k = 0; k < obj["relationships"].Count(); k++)
            //        {
            //            Cells(i + 1, couCode++, ((string)obj["relationships"][k]["federalTaxNumber"] ?? ""));
            //            Cells(i + 1, couCode++, ((string)obj["relationships"][k]["relationType"] ?? ""));
            //            Cells(i + 1, couCode++, ((string)obj["relationships"][k]["gender"] ?? ""));
            //            Cells(i + 1, couCode++, ((string)obj["relationships"][k]["status"] ?? ""));
            //            Cells(i + 1, couCode++, ((string)obj["relationships"][k]["name"] ?? ""));
            //        }
            //    }

            //    coEC = (couCode > coEC) ? couCode : coEC;
            //    countParent = couCode;
            //}

            //Cells(i + 1, countParent++, ((string)obj["legalNature"]["code"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["legalNature"]["description"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["streetSuffix"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["street"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["number"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["additionalInformation"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["district"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["city"]["code"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["city"]["name"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["state"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["address"]["country"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["email"] ?? ""));

            //if ((obj["phones"]) != null)
            //{
            //    var couPhone = countParent;
            //    var x = 0;
            //    for (x = 0; x < obj["phones"].Count(); x++)
            //    {
            //        Cells(i + 1, couPhone++, ((string)obj["phones"][x]["source"] ?? ""));
            //        Cells(i + 1, couPhone++, ((string)obj["phones"][x]["ddd"] ?? ""));
            //        Cells(i + 1, couPhone++, ((string)obj["phones"][x]["number"] ?? ""));
            //    }

            //    coPh = (couPhone > coPh) ? couPhone : coPh;
            //    countParent = couPhone;
            //}

            //Cells(i + 1, countParent++, ((string)obj["status"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["statusOn"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["statusReason"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["shareCapital"] ?? ""));

            //if ((obj["partners"]) != null)
            //{
            //    var couPartner = countParent;
            //    var j = 0;
            //    for (j = 0; j < obj["partners"].Count(); j++)
            //    {
            //        Cells(i + 1, countParent++, ((string)obj["partners"][j]["qualification"]["code"] ?? ""));
            //        Cells(i + 1, countParent++, ((string)obj["partners"][j]["qualification"]["description"] ?? ""));
            //        Cells(i + 1, countParent++, ((string)obj["partners"][j]["name"] ?? ""));
            //    }

            //    countParent = couPartner;
            //}

            //Cells(i + 1, countParent++, ((string)obj["unit"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["issuedOn"] ?? ""));
            //Cells(i + 1, countParent++, ((string)obj["issuedOn"] ?? ""));
        }

            return new Tuple<int, int>(coEC, coPh);
        }
}
}
