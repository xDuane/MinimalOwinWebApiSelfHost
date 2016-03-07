using MinimalOwinWebApiSelfHost.Models;
using OfficeOpenXml;
//using RDS.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Caching;

namespace RDS.Classes
{
    public class FileProcess
    {
        public static IEnumerable<Company> loadFile() // string fileType, Stream r)
        {
            //TODO: Config this
            var r = new FileStream(@"../../Resources/file.xlsx", FileMode.Open);
            //switch (fileType)
            //{
            //    case "Excel":
            using (var package = new ExcelPackage(r))
            {
                var currentSheet = package.Workbook.Worksheets["Sheet 1"];
                var workSheet = currentSheet;
                var noOfCol = workSheet.Dimension.End.Column;
                var noOfRow = workSheet.Dimension.End.Row;

                ////Ordinal
                //var columnHeaders = new Dictionary<string, int>();
                //for (int colIterator = 1; colIterator <= noOfCol; colIterator++)
                //{
                //    var cCell = workSheet.Cells[2, colIterator];
                //    columnHeaders.Add((cCell.Value ?? "").ToString(), cCell.Start.Column);
                //}

                //var db = new Entities();// RDBIEbCntities();
                //TODO: Use a SEQUENCE identifier
                //int nextValue = 1;
                //var maxRow = db.RDBIs.OrderByDescending(n => n.FILEID).FirstOrDefault();
                //if (maxRow != null)
                //{
                //    nextValue = 1 + maxRow.FILEID;
                //}

                //var file = new Models.File() { FILEID = nextValue, FileDate = DateTime.Now, UserID = 1 };
                List<Company> companies = new List<Company>();
                for (int rowIterator = 3; rowIterator <= noOfRow; rowIterator++)
                {
                    try
                    {
                        var company = new Company();
                        company.Id = Convert.ToInt32(workSheet.Cells[String.Format("A{0}", rowIterator)].Value);
                        company.Name = (workSheet.Cells[rowIterator, 4].Value ?? "").ToString();
                        company.zip = (workSheet.Cells[String.Format("R{0}", rowIterator)].Value ?? "").ToString();
                        companies.Add(company);
                    }
                    catch (Exception ee)
                    {
                        string s = ee.Message;
                    }
                    //var rdbi = new RDBI();
                    //rdbi.FILEID = nextValue; //TODO: CONCURRENCY ALERT
                    //foreach (var propertyInfo in rdbi.GetType().GetProperties())
                    //{
                    //    if (propertyInfo.CanRead)
                    //    {
                    //        if (!columnHeaders.ContainsKey(propertyInfo.Name))
                    //        {
                    //            //Some column(s) added i.e. FILEID
                    //            continue;
                    //        }

                    //        var colIdx = columnHeaders[propertyInfo.Name];
                    //        var val = workSheet.Cells[rowIterator, colIdx].Value;
                    //        if (val != null)
                    //        {
                    //            var typ = propertyInfo.PropertyType;
                    //            if (typ == typeof(Int32))
                    //            {
                    //                var cval = Convert.ToInt32(val);
                    //                propertyInfo.SetValue(rdbi, cval, null);
                    //            }
                    //            else if (typ == typeof(Int32?))
                    //            {
                    //                Int32? nint;
                    //                nint = Convert.ToInt32(val);
                    //                propertyInfo.SetValue(rdbi, nint, null);
                    //            }
                    //            else if (typ == typeof(double))
                    //            {
                    //                propertyInfo.SetValue(rdbi, val, null);
                    //            }
                    //            else if (typ == typeof(double?))
                    //            {
                    //                double? nint;
                    //                nint = Convert.ToDouble(val);
                    //                propertyInfo.SetValue(rdbi, nint, null);
                    //            }
                    //            else {
                    //                propertyInfo.SetValue(rdbi, val.ToString(), null);
                    //            }
                    //        }
                    //    }
                    //}
                    //file.RDBIs.Add(rdbi);
                }
                //db.Files.Add(file);
                //db.SaveChanges();
                //    }
                //    break;
                //case "CSV":
                //    break;
                r.Close();
                r.Dispose();

                var cache = MemoryCache.Default;
                cache["_companies"] = companies;
                cache["_cacheTime"] = DateTime.Now;

                return companies;
            }
        }

        public static void GenerateSchema(string fileType, Stream r)
        {
            switch (fileType)
            {
                case "Excel":
                    using (var package = new ExcelPackage(r))
                    {
                        var currentSheet = package.Workbook.Worksheets["Sheet 1"];
                        var workSheet = currentSheet; //.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = 10; // workSheet.Dimension.End.Row;

                        Debug.WriteLine("CREATE TABLE RDBI(");
                        //var types = new List<string>();
                        for (int colIterator = 1; colIterator <= noOfCol; colIterator++)
                        {
                            //Get the max value length 
                            int len = 0;
                            Type prevType = null;
                            var isDouble = false;
                            for (int rowIterator = 3; rowIterator <= noOfRow; rowIterator++)
                            {
                                var cCell = workSheet.Cells[rowIterator, colIterator];
                                if (cCell.Address == "CE4")
                                {
                                    string f = "";
                                }

                                if (cCell.Value == null)
                                {
                                    continue;
                                }

                                var currType = cCell.Value.GetType();
                                if ((prevType ?? currType) != currType)
                                {
                                    if (prevType == typeof(string))
                                    {
                                        //Skip it. String wins
                                        continue;
                                    }
                                }

                                if (!isDouble && currType == typeof(Double))
                                {
                                    int outInt = -1;
                                    isDouble = !int.TryParse(cCell.Value.ToString(), out outInt);
                                }

                                len = Math.Max(len, cCell.Value.ToString().Length);
                                prevType = currType;
                            }

                            var cell = workSheet.Cells[2, colIterator];
                            if (prevType == typeof(Double) && !isDouble)
                            {
                                Debug.WriteLine(String.Format("{0} int,", cell.Value));
                            }
                            else
                            if (prevType == typeof(Double))
                            {
                                Debug.WriteLine(String.Format("{0} float,", cell.Value));
                            }
                            else
                            {
                                Debug.WriteLine(String.Format("{0} varchar({1}),", cell.Value, Math.Max(8, roundUp(len))));
                            }
                        }
                        Debug.WriteLine(");");
                    }
                    break;
                case "CSV":
                    break;
            }
        }

        private static int roundUp(int numToRound)
        {
            int multiple = 8;
            int remainder = numToRound % multiple;
            if (remainder == 0)
                return numToRound;
            return numToRound + multiple - remainder;
        }


    }
}