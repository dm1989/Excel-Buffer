﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Buffer
{
    public static class Output
    {
        public static void OutputListToExcel<T>(List<T> dataToOutput, string targetPath, string targetSheetName, int startingRow)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Range xlRange;
            Excel.Workbook xlWorkbook;
            xlWorkbook = xlApp.Workbooks.Open(targetPath);
            Type model = typeof(T);
            PropertyInfo[] properties = model.GetProperties();
            ConstructorInfo constructor = model.GetConstructor(Type.EmptyTypes);
            List<MethodInfo> members = model.GetMethods().ToList();
            foreach (Excel._Worksheet xlWorksheet in xlWorkbook.Worksheets)
            {
                if (xlWorksheet.Name.Equals(targetSheetName))
                {
                    xlRange = xlWorksheet.UsedRange;
                    xlRange.Clear();
                    for (int i = 1; i <= properties.Count(); i++)
                    {
                        xlWorksheet.Cells[startingRow, i] = properties[i - 1].Name;
                    }
                    int rowIndex = startingRow;
                    foreach (var entry in dataToOutput)
                    {
                        rowIndex++;
                        Console.WriteLine("Populating " + targetPath + ": " + targetSheetName + " Row " + rowIndex);
                        int methodIndex = -2;
                        for (int columnIndex = 1; columnIndex <= properties.Count(); columnIndex++)
                        {
                            Type columnType = properties[columnIndex - 1].PropertyType;
                            methodIndex += 2;
                            if (columnType.Equals(typeof(string)))
                            {
                                xlWorksheet.Cells[rowIndex, columnIndex].NumberFormat = "@";
                                xlWorksheet.Cells[rowIndex, columnIndex] = members[methodIndex].Invoke(entry, new object[] { });
                            }
                            else if (columnType.Equals(typeof(DateTime)))
                            {
                                xlWorksheet.Cells[rowIndex, columnIndex].NumberFormat = "Mmm-DD-YYYY";
                                xlWorksheet.Cells[rowIndex, columnIndex] = members[methodIndex].Invoke(entry, new object[] { });
                            }
                            else if (columnType.Equals(typeof(int)))
                            {
                                xlWorksheet.Cells[rowIndex, columnIndex].NumberFormat = "#";
                                xlWorksheet.Cells[rowIndex, columnIndex] = members[methodIndex].Invoke(entry, new object[] { });
                            }
                            else if (columnType.Equals(typeof(decimal)))
                            {
                                xlWorksheet.Cells[rowIndex, columnIndex].NumberFormat = "#.##";
                                xlWorksheet.Cells[rowIndex, columnIndex] = members[methodIndex].Invoke(entry, new object[] { });
                            }
                        }
                    }
                }
            }
            xlWorkbook.Close(true);
        }
    }
}
}