using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Collections.Generic;
using Excel=Microsoft.Office.Interop.Excel;

namespace excelProcess
{
    public class patient
    {
        public string? BarCode { get; set; }
        public string? ProjectCode { get; set; }
        public string? ProjectResult { get; set; }
    }
    public class Program
    {
        public static void Main(string[] args)
        {
            const int insertColumnStart = 10;
            const int barCodeIndex = 7;
            const int projectCodeIndex = 10;
            const int projectResultIndex = 11;
            const int deleteColumns = 2;

            Console.Write($"XLS/XLSX file: ");
            string? inputName=Console.ReadLine();
            if(inputName==null)
            { return; }
            else { inputName = AdjustXLSAndXLSX(inputName); }

            string xlsPath = GetAbsolutePath(inputName);
            string resultFileName = GetResultFileName(xlsPath);

            var excelApp = new Excel.Application();
            //Show the Excel Window
            excelApp.Visible = false;

            Excel.Workbook xlsWorkbook;
            try
            {
                xlsWorkbook = excelApp.Workbooks.Open(xlsPath);
            }
            catch(Exception ex){
                Console.WriteLine(ex.Message);
                return;
            }

            Excel.Sheets allWorksheets = xlsWorkbook.Worksheets;
            Excel.Worksheet firstSheet = allWorksheets[1];
            firstSheet.Activate();

            int totalRows, totalColumns;
            (totalRows, totalColumns) = CurrentUsedRange(firstSheet);

            var ResultInfo=ReadAll(firstSheet,totalRows,barCodeIndex,projectCodeIndex,projectResultIndex);
            List<string?> uniqueProjectName=new List<string?>();

            uniqueProjectName=ResultInfo.Select(x => x.ProjectCode).Distinct().ToList();

            for (int i = 0; i < deleteColumns; i++)
            {
                (firstSheet.Columns[insertColumnStart]).Delete();
            }

            for(int i=0;i<uniqueProjectName.Count;i++)
            {
                int insertColumn=insertColumnStart+i;
                (firstSheet.Columns[insertColumn]).Insert();
                (firstSheet.Cells[1,insertColumn]).Value=uniqueProjectName[i];
            }

            DeleteDuplicate(firstSheet,totalRows,barCodeIndex);

            int totalUniqueRows;
            (totalUniqueRows,_) = CurrentUsedRange(firstSheet);
            var barCodeUnique=new List<string>();
            for(int i=2;i<=totalUniqueRows;i++)
            {
                string barCode = (string)(firstSheet.Cells[i, barCodeIndex]).Value;
                barCodeUnique.Add(barCode);
            }

            InsertAllData(barCodeUnique, ResultInfo, uniqueProjectName, firstSheet, insertColumnStart);
            xlsWorkbook.SaveAs2(resultFileName);
            xlsWorkbook.Close();
            excelApp.Quit();
        }
        public static void DeleteDuplicate(Excel.Worksheet sheet,int edge,int codeIndex)
        {
            var deleteIndex=new List<int>();
            int i = 2;
            int inI = 1;
            //Don't include last row in i
            while (i < edge)
            {
                string barCodeStandard=(string)(sheet.Cells[i,codeIndex]).Value;
                string barCodeNext=(string)(sheet.Cells[(i+inI),codeIndex]).Value;
                if(barCodeStandard==barCodeNext)
                {
                    deleteIndex.Add((i+inI));
                    inI++;
                }
                else
                {
                    i += inI;
                    inI = 1;
                }
            }
            deleteIndex.Reverse();
            foreach(int delNum in deleteIndex)
            {
                (sheet.Rows[delNum]).Delete();
            }
        }
        public static List<patient> ReadAll(Excel.Worksheet sheet,int edgeCount,int codeIndex,int projectCodeIndex,int projectResultIndex)
        {
            var allPatients=new List<patient>();
            for (int i = 2; i <=edgeCount; i++)
            {
                string barCode=(string)(sheet.Cells[i,codeIndex]).Value;
                string projectCode=(string)(sheet.Cells[i,projectCodeIndex]).Value;
                string projectResult=(string)(sheet.Cells[i,projectResultIndex]).Value;

                allPatients.Add(new patient() { BarCode=barCode,ProjectCode=projectCode,ProjectResult=projectResult,});
            }
            return allPatients;
        }
        public static string AdjustXLSAndXLSX(string inputString)
        {
            if (inputString.Contains('.'))
            {
                inputString = inputString.Trim();
            }
            else
            {
                string realName;
                realName = (inputString.Trim()) + @".xls";
                bool exist = File.Exists(realName);
                if (exist)
                {
                    inputString = realName;
                }
                else
                {
                    realName = (inputString.Trim()) + @".xlsx";
                    inputString = realName;
                }
            }
            return inputString;
        }

        public static string GetResultFileName(string inputAbsoluteName)
        {
            DateTime dateTime = DateTime.Now;
            string suffixFile = dateTime.ToString("_yyyyMMdd_HHmmss_FFF");
            string xlsPath = inputAbsoluteName;
            string ResultFileName = xlsPath.Insert(xlsPath.Length - (((xlsPath.Split('.')).Last()).Length) - 1, suffixFile);
            return ResultFileName;
        }

        public static string GetAbsolutePath(string reletiveName)=>Directory.GetCurrentDirectory()+Path.DirectorySeparatorChar
            + reletiveName;
        public static ValueTuple<int,int> CurrentUsedRange(Excel.Worksheet worksheet)
        {
            return (worksheet.UsedRange.Rows.Count,worksheet.UsedRange.Columns.Count);
        }

        public static void InsertAllData(List<string> barCodeUnique,List<patient> resultInfo,List<string?> uniqueProjectName
            ,Excel.Worksheet firstSheet,int insertColumnStart)
        {
            int allI = 0;
            int insertFlag = 0;
            while (insertFlag < barCodeUnique.Count)
            {
                if (allI >= resultInfo.Count)
                {
                    break;
                }
                else
                {
                    if (barCodeUnique[insertFlag] == resultInfo[allI].BarCode)
                    {
                        int indexOffset = (uniqueProjectName).FindIndex(a => a == resultInfo[allI].ProjectCode);
                        (firstSheet.Cells[(insertFlag + 2), (insertColumnStart + indexOffset)]).Value = resultInfo[allI].ProjectResult;
                        allI++;
                    }
                    else
                    {
                        insertFlag++;
                    }
                }
            }
        }
    }
}
