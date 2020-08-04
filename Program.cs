using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Spire;
using Spire.Xls;
using Spire.Xls.Collections;

namespace chart
{
    class Program
    {
      
        public static void updateExcelFile(string key, double value){
            // try
            // {   if(!data.ContainsKey(key)) return;
            //     var file = new FileInfo(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\template.xlsx");
            //     if(!file.Exists) return;

            //     using(ExcelPackage p = new ExcelPackage(file))
            //     {
            //         var ws = p.Workbook.Worksheets["DS"];
            //         data[key] = value;

            //         var pieChart = ws.Drawings["Chart 2"] as ExcelPieChart;
            //         var series = pieChart.Series[0];
                
            //         var startCell = ws.Cells[3,28];
            //         startCell.Offset(0, 0).Value = "Agent";
            //         startCell.Offset(0, 1).Value = "Client";


                   
            
            //         var pieSeries = (ExcelPieChartSerie)series;
            //         pieSeries.Explosion = 5;
            //         Console.Write(data.Count());
            //         int j = 0;
            //         foreach(KeyValuePair<string, double> entry in data)
            //         {
            //             startCell.Offset(j + 1, 0).Value = entry.Key;
            //             startCell.Offset(j + 1, 1).Value = entry.Value;
            //             ++j;
            //         }
            //         p.Save();
            //     }
            // }
            // catch (System.Exception)
            // {
            //     Console.Write("Error");
            //     throw;
            // }
            
        }
        public static void PieChartClientServedPerAgent(FileInfo file, Dictionary<string, double> newData)
        {
            // Dictionary<string, double> data = new Dictionary<string, double>();
            if (!file.Exists)
                return;
            try
            {
                var p = new ExcelPackage(file);
                var ws = p.Workbook.Worksheets["DS"];
              
                int jumpToData = 3;
                int backup = 30;
                // for(int i = ws.Dimension.Start.Row + jumpToData; i<= numberDataOld + jumpToData; i++)
                // {
                //     var row = ws.Cells["AB"+ i +":" + "AC" + i];
                //     int count = 0;
                 
                //     foreach(var cell in row)
                //     {   
                      
                //         if(count % 2 == 0)
                //         {
                //            name = cell.Text;
                //         }
                //         else{
                //            value = (double)cell.Value;
                //         }
                        
                //         count++;
                //     }
                //     data.Add(name, value);
                // }
                
                //new data
                // foreach( KeyValuePair<string, double> kvp in newData )
                // {
                //     data.Add(kvp.Key, kvp.Value);
                // }
                var sortedDict = from entry in newData orderby entry.Value ascending select entry;
                
                Dictionary<string, double> test = new Dictionary<string, double>();
                for(int i = 0, y = sortedDict.ToList().Count - 1; i <= sortedDict.ToList().Count / 2; i++, y--)
                {
                    if(i < y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                        test.Add(sortedDict.ToList()[y].Key, sortedDict.ToList()[y].Value);
                    }
                    else if (i == y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                    }
                    
                  
                }
               
                // data["Nam (1)"] = 0.01;
                using( var range = ws.Cells["AB" + jumpToData + ":" + "AC" + (newData.Count + jumpToData +backup) ])
                {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Font.Bold = true;
                        range.Style.Font.Name ="Cambria";
                        range.Style.Font.Size = 12;
                        range.Style.Font.Color.SetColor(Color.Black);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Value = "";
                }
                var pieChart = ws.Drawings["Chart 2"] as ExcelPieChart;
                var series = pieChart.Series[0];
               
                var startCell = ws.Cells[3,28];
                startCell.Offset(0, 0).Value = "Agent";
                startCell.Offset(0, 1).Value = "Client";
                ws.Cells["AD3"].Value = newData.Count;
           
                series.XSeries = ws.Cells[4, 28, newData.Count + jumpToData + backup, 28].FullAddress;
                series.Series = ws.Cells[4, 29, newData.Count + jumpToData + backup, 29].FullAddress;
             
        
                var pieSeries = (ExcelPieChartSerie)series;
                pieSeries.Explosion = 5;
                try
                {
                    for(int i = 0; i< test.ToList().Count; i++)
                    {
                        startCell.Offset(i + 1, 0).Value = test.ToList()[i].Key;
                        startCell.Offset(i + 1, 1).Value = test.ToList()[i].Value;
                    }
                    int startRow = test.ToList().Count + jumpToData + 1;
                    int endRow = startRow + backup;
                    Console.Write(endRow);
                    for (int s = startRow; s < endRow; s++)
                    {
                        ws.Cells[s, 29].Formula = "IF(AB" + (s) +"=\"\"" + "," + "NA()" + ")" ;
                        ws.Cells[s, 28, s, 29].Style.Font.Color.SetColor(Color.White);
                        ws.Cells[s, 28, s, 29].Style.Border.Top.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 28, s, 29].Style.Border.Bottom.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 28, s, 29].Style.Border.Right.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 28, s, 29].Style.Border.Left.Style =ExcelBorderStyle.None;
                    }
                
                
                }
                catch (System.Exception)
                {
                    throw;
                }
              
                 p.Save();
                // Byte[] bin = p.GetAsByteArray();
                // File.WriteAllBytes(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\temp_new.xlsx", bin);
                // p.Save();
            }
            catch (System.Exception)
            {
                
                throw;
            }
            
        
        }
        public static void PieChartInBoundCallPerAgent(FileInfo file, Dictionary<string, double> newData)
        {
            if (!file.Exists)
                return;
            try
            {
                var p = new ExcelPackage(file);
                var ws = p.Workbook.Worksheets["DS"];
              
                int jumpToData = 3;
                int backup = 30;
               
                var sortedDict = from entry in newData orderby entry.Value ascending select entry;
                
                Dictionary<string, double> test = new Dictionary<string, double>();
                for(int i = 0, y = sortedDict.ToList().Count - 1; i <= sortedDict.ToList().Count / 2; i++, y--)
                {
                    if(i < y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                        test.Add(sortedDict.ToList()[y].Key, sortedDict.ToList()[y].Value);
                    }
                    else if (i == y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                    }
                    
                  
                }
               
                // data["Nam (1)"] = 0.01;
                using( var range = ws.Cells["AF" + jumpToData + ":" + "AG" + (newData.Count + jumpToData +backup) ])
                {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Font.Bold = true;
                        range.Style.Font.Name ="Cambria";
                        range.Style.Font.Size = 12;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Value = "";
                        range.Style.Font.Color.SetColor(Color.Black);
                }
                var pieChart = ws.Drawings["Chart 3"] as ExcelPieChart;
                var series = pieChart.Series[0];
               
                var startCell = ws.Cells[3,32];
                startCell.Offset(0, 0).Value = "Agent";
                startCell.Offset(0, 1).Value = "Inbound";
                ws.Cells["AH3"].Value = newData.Count;
           
                series.XSeries = ws.Cells[4, 32, newData.Count + jumpToData + backup, 32].FullAddress;
                series.Series = ws.Cells[4, 33, newData.Count + jumpToData + backup, 33].FullAddress;
             
        
                var pieSeries = (ExcelPieChartSerie)series;
                pieSeries.Explosion = 5;
                try
                {
                    for(int i = 0; i< test.ToList().Count; i++)
                    {
                        startCell.Offset(i + 1, 0).Value = test.ToList()[i].Key;
                        startCell.Offset(i + 1, 1).Value = test.ToList()[i].Value;
                    }
                    int startRow = test.ToList().Count + jumpToData + 1;
                    int endRow = startRow + backup;
                   
                    for (int s = startRow; s < endRow; s++)
                    {
                        ws.Cells[s, 33].Formula = "IF(AF" + (s) +"=\"\"" + "," + "NA()" + ")" ;
                        ws.Cells[s, 32, s, 33].Style.Font.Color.SetColor(Color.White);
                        ws.Cells[s, 32, s, 33].Style.Border.Top.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 32, s, 33].Style.Border.Bottom.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 32, s, 33].Style.Border.Right.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 32, s, 33].Style.Border.Left.Style =ExcelBorderStyle.None;
                    }
                
                
                }
                catch (System.Exception)
                {
                    throw;
                }
              
                 p.Save();
                // Byte[] bin = p.GetAsByteArray();
                // File.WriteAllBytes(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\temp_new.xlsx", bin);
                // p.Save();
            }
            catch (System.Exception)
            {
                
                throw;
            }
            
        
        }
        public static void PieChartOutBoundCallPerAgent(FileInfo file, Dictionary<string, double> newData)
        {
            if (!file.Exists)
                return;
            try
            {
                var p = new ExcelPackage(file);
                var ws = p.Workbook.Worksheets["DS"];
              
                int jumpToData = 3;
                int backup = 30;
               
                var sortedDict = from entry in newData orderby entry.Value ascending select entry;
                
                Dictionary<string, double> test = new Dictionary<string, double>();
                for(int i = 0, y = sortedDict.ToList().Count - 1; i <= sortedDict.ToList().Count / 2; i++, y--)
                {
                    if(i < y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                        test.Add(sortedDict.ToList()[y].Key, sortedDict.ToList()[y].Value);
                    }
                    else if (i == y)
                    {
                        test.Add(sortedDict.ToList()[i].Key, sortedDict.ToList()[i].Value);
                    }
                    
                  
                }
               
                // data["Nam (1)"] = 0.01;
                using( var range = ws.Cells["AJ" + jumpToData + ":" + "AK" + (newData.Count + jumpToData +backup) ])
                {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Font.Bold = true;
                        range.Style.Font.Name ="Cambria";
                        range.Style.Font.Size = 12;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Value = "";
                        range.Style.Font.Color.SetColor(Color.Black);
                }
                var pieChart = ws.Drawings["Chart 5"] as ExcelPieChart;
                var series = pieChart.Series[0];
               
                var startCell = ws.Cells[3,36];
                startCell.Offset(0, 0).Value = "Agent";
                startCell.Offset(0, 1).Value = "Outbound";
                ws.Cells["AL3"].Value = newData.Count;
           
                series.XSeries = ws.Cells[4, 36, newData.Count + jumpToData + backup, 36].FullAddress;
                series.Series = ws.Cells[4, 37, newData.Count + jumpToData + backup, 37].FullAddress;
             
        
                var pieSeries = (ExcelPieChartSerie)series;
                pieSeries.Explosion = 5;
                try
                {
                    for(int i = 0; i< test.ToList().Count; i++)
                    {
                        startCell.Offset(i + 1, 0).Value = test.ToList()[i].Key;
                        startCell.Offset(i + 1, 1).Value = test.ToList()[i].Value;
                    }
                    int startRow = test.ToList().Count + jumpToData + 1;
                    int endRow = startRow + backup;
                   
                    for (int s = startRow; s < endRow; s++)
                    {
                        ws.Cells[s, 37].Formula = "IF(AJ" + (s) +"=\"\"" + "," + "NA()" + ")" ;
                        ws.Cells[s, 36, s, 37].Style.Font.Color.SetColor(Color.White);
                        ws.Cells[s, 36, s, 37].Style.Border.Top.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 36, s, 37].Style.Border.Bottom.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 36, s, 37].Style.Border.Right.Style =ExcelBorderStyle.None;
                        ws.Cells[s, 36, s, 37].Style.Border.Left.Style =ExcelBorderStyle.None;
                    }
                
                
                }
                catch (System.Exception)
                {
                    throw;
                }
              
                 p.Save();
                // Byte[] bin = p.GetAsByteArray();
                // File.WriteAllBytes(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\temp_new.xlsx", bin);
                // p.Save();
            }
            catch (System.Exception)
            {
                
                throw;
            }
            
        
        }
        public static void PieChartCreate(ref ExcelPackage p, Dictionary<string, double> newData){
        
            var worksheet = p.Workbook.Worksheets["Agent"];
            
            //Fill the table
            var startCell = worksheet.Cells["T2"];
            for(int i = 0; i< newData.ToList().Count; i++)
            {
                startCell.Offset(i , 0).Formula ="=\""+newData.ToList()[i].Key +  " (\"&V"+(i + 2)+"&\" mins)\"" ;
                startCell.Offset(i , 2).Value = newData.ToList()[i].Value;
                startCell.Offset(i , 1).Formula ="=V" + (i + 2) + "/" + 480;
            }
            //Delete old Chart
            var excelDrawing = worksheet.Drawings["Chart 1"];
            worksheet.Drawings.Remove(excelDrawing);
            // Add the chart to the sheet
            var pieChart = worksheet.Drawings.AddChart("Chart 1", eChartType.Pie);
            pieChart.SetPosition(newData.Count + 8, 4, newData.Count + 1, 0);
            pieChart.Title.Text = "Test Chart";
            pieChart.Title.Font.Bold = true;
            pieChart.Title.Font.Size = 12;
            pieChart.SetSize(800, 300);
             //Set the data range
            var series = pieChart.Series.Add("U2:U4", "T2:T4");
            var pieSeries = (ExcelPieChartSerie)series;
            pieSeries.Explosion = 5;

            //Format the labels
            pieSeries.DataLabel.Font.Bold = true;
            pieSeries.DataLabel.ShowPercent = true;

            //Format the legend
            pieChart.Legend.Add();
            pieChart.Legend.Border.Width = 0;
            pieChart.Legend.Font.Size = 12;
            pieChart.Legend.Font.Bold = true;
            pieChart.Legend.Position = eLegendPosition.Right;

            // Byte[] bin = p.GetAsByteArray();
            // File.WriteAllBytes(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\temp_new.xlsx", bin);
            p.Save();

        }
        public static void RenderAgent(FileInfo file, string name, Dictionary<string, double> newData )
        {
             var p = new ExcelPackage(file);
            var worksheet = p.Workbook.Worksheets["Agent"];
            worksheet.Cells["B2"].Value = name;
            RenderTableGlip(ref p, 30 , 80 );
            RenderTableEpic(ref p, 10, 20);
            RenderTableEmail(ref p, 10, 20, 30);
            RenderTablePhone(ref p, 10,20,30,40);
            PieChartCreate(ref p, newData);
        }
        public static void RenderDS(FileInfo file, Dictionary<string, double> newData)
        {
            PieChartClientServedPerAgent(file, newData);
            PieChartInBoundCallPerAgent(file, newData);
            PieChartOutBoundCallPerAgent(file, newData);
        }
        public static void RenderTableGlip(ref ExcelPackage p ,double served, double totalDuration)
        {
            var ws = p.Workbook.Worksheets["Agent"];
            ws.Cells["C5"].Value = served;
            ws.Cells["C6"].Value = totalDuration;

        }
        public static void RenderTableEpic(ref ExcelPackage p, double activityLogged, double attachment)
        {
            var ws = p.Workbook.Worksheets["Agent"];
            ws.Cells["J5"].Value = activityLogged;
            ws.Cells["J6"].Value = attachment;
        }
        public static void RenderTableEmail(ref ExcelPackage p, double received, double sent, double conversations)
        {
            var ws = p.Workbook.Worksheets["Agent"];
            ws.Cells["M5"].Value = received;
            ws.Cells["M6"].Value = sent;
            ws.Cells["M7"].Value = conversations;
        }
        public static void RenderTablePhone (ref ExcelPackage p, int numberCallInbound, int durationInbound, int numberCallOutbound, int durationOutbound)
        {
            var ws = p.Workbook.Worksheets["Agent"];
            ws.Cells["F6"].Value = numberCallInbound;
            ws.Cells["F7"].Value = durationInbound + " min";
            ws.Cells["G6"].Value = numberCallOutbound ;
            ws.Cells["G7"].Value = durationOutbound + " min";
        }
        public static void AutoFilter_Test()
        {
            
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\template.xlsx");
            Worksheet sheet = workbook.Worksheets["FD"];
            sheet.Activate();
            AutoFiltersCollection filters = sheet.AutoFilters;
            filters.Range = sheet.Range[3, 8, sheet.LastRow, 8];   
           
            filters.AddFilter(0, "Walk-in");
            filters.Filter(); 
            workbook.SaveToFile("output.xlsx", ExcelVersion.Version2010); 
        }
        public static void ClearEvaluationWarning()
        {
            var file = new FileInfo("output.xlsx");
            using (var excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Evaluation Warning");
                excelPackage.Workbook.Worksheets.Delete(worksheet);
                excelPackage.Save();
            }
          
            
        }
        static void Main(string[] args)
        {
            // updateExcelFile(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\Commercial_Template.xlsx");
            //  PieChartCreate();
             var newData = new Dictionary<string, double>();
            
                newData.Add("Node (1)", 0.05);
                newData.Add("Node (2)", 0.1);
                newData.Add("Node (3)", 0.2);
                newData.Add("Node (4)", 0.25);
                newData.Add("Node (5)", 0.15);
                newData.Add("Node (6)", 0.03);
                newData.Add("Node (7)", 0.2);
                newData.Add("Node (8)", 0.22);
                newData.Add("Node (9)", 0.12);
                newData.Add("Node (10)", 0.02);
                newData.Add("Node (11)", 0.02);
                newData.Add("Node (12)", 0.1);
                newData.Add("Node (13)", 0.02);
                newData.Add("Node (14)", 0.1);
                newData.Add("Node (15)", 0.02);
                newData.Add("Node (16)", 0.1);
                newData.Add("Node (17)", 0.02);
                // newData.Add("Node (18)", 0.1);
                // newData.Add("Node (19)", 0.02);
                // newData.Add("Node (20)", 0.1);
                // newData.Add("Node (21)", 0.02);
                // newData.Add("Node (22)", 0.1);
                // newData.Add("Node (23)", 0.1);
                // newData.Add("Node (24)", 0.02);
                // newData.Add("Node (25)", 0.1);
                // newData.Add("Node (26)", 0.02);
                // newData.Add("Node (27)", 0.1);
                // newData.Add("Node (28)", 0.1);
                // newData.Add("Node (29)", 0.02);
                // newData.Add("Node (30)", 0.1);
                // newData.Add("Node (31)", 0.1);
                // newData.Add("Node (32)", 0.02);
                // newData.Add("Node (33)", 0.1);
                
            
            var testData = new Dictionary<string, double>();
            testData.Add("Glip - FD & TF", 335);
            testData.Add("Glip - FD & TF 2", 100);
            testData.Add("Glip - FD & TF 3", 40);
     
            //  RenderDS(new FileInfo(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\template.xlsx"), newData);
            //    RenderAgent(new FileInfo(@"D:\dotnet\chart\bin\Debug\netcoreapp3.1\template.xlsx"),"Erik",  testData);
             AutoFilter_Test();
            //  ClearEvaluationWarning();
        }
    }
}
