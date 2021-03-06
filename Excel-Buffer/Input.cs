﻿using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Excel_Buffer
{
    public static class Input
    {
        public static List<T> PullExcelAsList<T>(string excelPath, int startingRow, string sheetName = "", DataSet blankPathSource = null)
        {
            
            DataSet ExcelSheet;
            if (blankPathSource == null)
            {
                ExcelSheet = GetExcelAsDataSet(excelPath);
            }
            else
            {
                ExcelSheet = blankPathSource;
            }
            Type model = typeof(T);
            PropertyInfo[] properties = model.GetProperties();
            ConstructorInfo constructor = model.GetConstructor(Type.EmptyTypes);
            List<MethodInfo> members = model.GetMethods().ToList();
            var returnList = new List<T>();
            if (sheetName.Equals(""))
            {
                sheetName = ExcelSheet.Tables[0].TableName;
            }
            foreach (DataTable table in ExcelSheet.Tables)
            {
                if (table.TableName.Equals(sheetName))
                {
                    int rowIndex = 0;
                    foreach (DataRow row in table.Rows)
                    {
                        rowIndex++;
                        if (rowIndex < startingRow) { continue; }
                        object entry = constructor.Invoke(new object[] { });
                        int columnIndex = 0;
                        int methodIndex = -1;
                        foreach (DataColumn column in table.Columns)
                        {
                            columnIndex++;
                            methodIndex += 2;
                            if (columnIndex > properties.Count()) { break; }
                            Type columnType = properties[columnIndex - 1].PropertyType;
                            if (columnType.Equals(typeof(string)))
                            {
                                try
                                {
                                    members[methodIndex].Invoke(entry, new object[] { row[column].ToString() });
                                }
                                catch
                                {
                                    members[methodIndex].Invoke(entry, new object[] { "" });
                                }

                            }
                            else if (columnType.Equals(typeof(DateTime)))
                            {
                                try
                                {
                                    members[methodIndex].Invoke(entry, new object[] { DateTime.Parse(row[column].ToString()) });
                                }
                                catch
                                {
                                    members[methodIndex].Invoke(entry, new object[] { new DateTime() });
                                }

                            }
                            else if (columnType.Equals(typeof(int)))
                            {
                                try
                                {
                                    members[methodIndex].Invoke(entry, new object[] { Int32.Parse(row[column].ToString()) });
                                }
                                catch
                                {
                                    members[methodIndex].Invoke(entry, new object[] { 0 });
                                }

                            }
                            else if (columnType.Equals(typeof(decimal)))
                            {
                                try
                                {
                                    members[methodIndex].Invoke(entry, new object[] { Decimal.Parse(row[column].ToString()) });
                                }
                                catch
                                {
                                    members[methodIndex].Invoke(entry, new object[] { (decimal)0 });
                                }
                            }
                            else if (columnType.Equals(typeof(double)))
                            {
                                try
                                {
                                    members[methodIndex].Invoke(entry, new object[] { Double.Parse(row[column].ToString()) });
                                }
                                catch
                                {
                                    members[methodIndex].Invoke(entry, new object[] { (double)0 });
                                }
                            }
                        }
                        returnList.Add((T)entry);
                    }
                }
            }
            return returnList;
        }
        private static DataSet GetExcelAsDataSet(string filePath)
        {
            DataSet result = null;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))//excel 2016 is xml
                {
                    result = reader.AsDataSet();
                }
            }
            return result;
        }
    }
}
