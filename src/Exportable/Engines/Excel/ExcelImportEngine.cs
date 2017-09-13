using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Exportable.Resources;
using Exportable.Tools;
using Infodinamica.Framework.Core.Containers;
using Infodinamica.Framework.Core.Extensions.Common;
using Infodinamica.Framework.Core.Extensions.Reflection;
using NPOI.SS.UserModel;

namespace Exportable.Engines.Excel
{
    /// <summary>
    /// Clase para importar Excel
    /// </summary>
    public class ExcelImportEngine: ExcelEngine, IExcelImportEngine
    {
        private static readonly object locker = new object();
        private readonly IList<Tuple<string, Type, string, int>> _containers;
        private MemoryStream _file;
        private bool _wasReaded;

        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelImportEngine()
        {
            _containers = new List<Tuple<string, Type, string, int>>();
            _wasReaded = false;
        }

        public void AddContainer<T>(string key) where T : class
        {
            _containers.Add(Tuple.Create(key, typeof (T), string.Empty, 0));
        }

        public void AddContainer<T>(string key, string sheetName) where T : class
        {
            _containers.Add(Tuple.Create(key, typeof(T), sheetName, 0));
        }

        public void AddContainer<T>(string key, string sheetName, int firsRowWithData) where T : class
        {
            _containers.Add(Tuple.Create(key, typeof(T), sheetName, firsRowWithData));
        }

        public void SetDocument(MemoryStream file)
        {
            _file = file;
        }

        public void SetDocument(string path)
        {
            _file = new MemoryStream();

            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                byte[] buffer = new byte[32 * 1024]; // 32K buffer for example
                int bytesRead;
                while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    _file.Write(buffer, 0, bytesRead);
                }
            }
            _file.Position = 0;
        }
        
        public IList<T> GetList<T>(string key) where T : class
        {
            if (!_wasReaded)
            {
                lock (locker)
                {
                    if (!_wasReaded)
                    {
                        ReadFromMemoryStream(_file);
                        _wasReaded = true;
                    }
                }
            }
            
            IList<T> data = new List<T>();
            var tuple = _containers.FirstOrDefault(c => c.Item1 == key);

            //Check tuple exists
            if (tuple == null || StringMethods.IsNullOrWhiteSpace(tuple.Item1))
                throw new Exception(string.Format(ErrorMessage.KeyNotFound, key));

            //Get sheet name from tuple or attributes
            var sheetName = tuple.Item3;
            if (StringMethods.IsNullOrWhiteSpace(sheetName))
                sheetName = MetadataHelper.GetSheetNameFromAttribute(tuple.Item2);

            if(StringMethods.IsNullOrWhiteSpace(sheetName))
                throw new Exception(ErrorMessage.SheetNameNotProvided);

            //Get sheet
            ISheet sheet = GetSheet(sheetName);

            //Check sheet exists
            if(sheet == null)
                throw new Exception(string.Format(ErrorMessage.SheetNotFound, sheetName));

            //Get properties from T
            var properties = MetadataHelper.GetImportableMetadatas(typeof(T));

            //Get firs row with data
            var firstRowInAttribute = MetadataHelper.GetFirstRowWithDataFromAttribute(typeof (T));
            var firstRowInTuple = tuple.Item4;
            var firstRow = sheet.FirstRowNum <= firstRowInAttribute ? firstRowInAttribute : sheet.FirstRowNum;
            firstRow = firstRow >= firstRowInTuple ? firstRow : firstRowInTuple;

            for (int rowIndex = firstRow; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                //Fix error when excel has blank/empty lines
                if (row == null)
                    continue;
                
                int colIndex = 0;
                T t = (T)Activator.CreateInstance(typeof(T));
                foreach (var cell in row.Cells)
                {
                    //Get metadata of column
                    var instanceMetadata = properties.FirstOrDefault(p => p.Position == colIndex);
                    
                    //there is not field mapeable to T
                    if (instanceMetadata == null)
                    {
                        colIndex++; 
                        continue;
                    }

                    //Get propertyInfo
                    var instanceProperty = t.GetType().GetProperty(instanceMetadata.Name);
                    
                    //Set numeric value
                    if (instanceProperty != null && instanceProperty.IsNumeric())
                    {
                        double numberValue;
                        if (!TryGetNumber(cell, out numberValue))
                            if (!double.TryParse(instanceMetadata.DefaultForNullOrInvalidValues, out numberValue))
                                throw new Exception(string.Format(ErrorMessage.CannotParseNumber, GetText(cell)));
                        instanceProperty.SetValue(t, Convert.ChangeType(numberValue, instanceProperty.PropertyType), null);
                    }

                    //Set date value
                    else if (instanceProperty != null && instanceProperty.IsDateOrTime())
                    {
                        DateTime dateValue;
                        if (!TryGetDate(cell, out dateValue))
                            if (!DateTime.TryParse(instanceMetadata.DefaultForNullOrInvalidValues, out dateValue))
                                throw new Exception(string.Format(ErrorMessage.CannotParseDatetime, GetText(cell)));
                        instanceProperty.SetValue(t, Convert.ChangeType(dateValue, instanceProperty.PropertyType), null);
                    }

                    //Set boolean value
                    else if (instanceProperty != null && instanceProperty.IsBoolean())
                    {
                        bool boolValue;
                        if(!TryGetBool(cell, out boolValue))
                            if (!bool.TryParse(instanceMetadata.DefaultForNullOrInvalidValues, out boolValue))
                                throw new Exception(string.Format(ErrorMessage.CannotParseBoolean, GetText(cell)));
                        instanceProperty.SetValue(t, Convert.ChangeType(boolValue, instanceProperty.PropertyType), null);
                    }

                    //Else is string
                    else
                    {
                        string value = string.Empty;
                        if (TryGetText(cell, out value))
                            instanceProperty.SetValue(t, Convert.ChangeType(value, instanceProperty.PropertyType), null);
                        else if (instanceMetadata.DefaultForNullOrInvalidValues != null)
                            instanceProperty.SetValue(t, Convert.ChangeType(instanceMetadata.DefaultForNullOrInvalidValues, instanceProperty.PropertyType), null);
                        else
                            throw new Exception(ErrorMessage.CannotParseString);
                    }

                    colIndex++; 
                }

                //Add data to collection
                data.Add(t);
            }

            return data;
        }

        public IDictionary<string, string> RunBusinessRules()
        {
            throw new NotImplementedException();
        }

        public IExcelImportEngine AsExcel()
        {
            return this;
        }

        private bool TryGetNumber(ICell cell, out double returnValue)
        {
            returnValue = -1;

            try
            {
                returnValue = cell.NumericCellValue;
                return true;
            } catch { }

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Boolean:
                        returnValue = cell.BooleanCellValue ? 1 : 0;
                        return true;
                    case CellType.Error:
                        return false;
                    case CellType.Formula:
                        return false;
                    case CellType.Numeric:
                        returnValue = cell.NumericCellValue;
                        return true;
                    case CellType.String:
                        return double.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    case CellType.Unknown:
                        return double.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetText(ICell cell, out string returnValue)
        {
            returnValue = string.Empty;

            try
            {
                returnValue = cell.StringCellValue;
                return true;
            } catch { }

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        returnValue = string.Empty;
                        return true;
                    case CellType.Boolean:
                        returnValue = cell.BooleanCellValue.ToString();
                        return true;
                    case CellType.Error:
                        return false;
                    case CellType.Formula:
                        return false;
                    case CellType.Numeric:
                        returnValue = cell.NumericCellValue.ToString();
                        return true;
                    case CellType.String:
                        returnValue = cell.StringCellValue;
                        return true;
                    case CellType.Unknown:
                        returnValue = cell.StringCellValue;
                        return true;
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetBool(ICell cell, out bool returnValue)
        {
            returnValue = false;

            try
            {
                returnValue = cell.BooleanCellValue;
                return true;
            } catch { }


            try
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        returnValue = false;
                        return true;
                    case CellType.Boolean:
                        returnValue = cell.BooleanCellValue;
                        return true;
                    case CellType.Error:
                        return false;
                    case CellType.Formula:
                        return false;
                    case CellType.Numeric:
                        if (cell.NumericCellValue == 0)
                        {
                            returnValue = false;
                            return true;
                        } else if (cell.NumericCellValue == 1)
                        {
                            returnValue = true;
                            return true;
                        }
                        else
                            return false;
                    case CellType.String:
                        return bool.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    case CellType.Unknown:
                        return bool.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetDate(ICell cell, out DateTime returnValue)
        {
            returnValue = new DateTime();

            //First try to get the value from date cell
            try
            {
                returnValue = cell.DateCellValue;
                return true;
            }
            catch { }


            try
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Boolean:
                        return false;
                    case CellType.Error:
                        return false;
                    case CellType.Formula:
                        return false;
                    case CellType.Numeric:
                        return DateTime.TryParse(cell.NumericCellValue.ToString(), out returnValue);
                    case CellType.String:
                        return DateTime.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    case CellType.Unknown:
                        return DateTime.TryParse(cell.StringCellValue.Trim(), out returnValue);
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private string GetText(ICell cell)
        {
            var value = string.Empty;
            if (TryGetText(cell, out value))
                return value;
            else
                return String.Empty;
        }
    }
}
