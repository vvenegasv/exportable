using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Infodinamica.Framework.Core.Extensions.Common;
using Infodinamica.Framework.Core.Extensions.Reflection;
using Infodinamica.Framework.Core.Types;
using Infodinamica.Framework.Exportable.Resources;
using Infodinamica.Framework.Exportable.Tools;
using NPOI.SS.UserModel;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public class ExcelImportEngine: ExcelEngine, IExcelImportEngine
    {
        private readonly IList<Tuple<string, Type, string, int>> _containers;
        private MemoryStream _file;
        private bool _wasReaded;

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
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[file.Length];
                file.Read(bytes, 0, (int) file.Length);
                _file.Write(bytes, 0, (int) file.Length);
            }
            _file.Position = 0;
        }



        public IList<T> GetList<T>(string key) where T : class
        {
            if (!_wasReaded)
            {
                ReadFromMemoryStream(_file);
                _wasReaded = true;
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
                        instanceProperty.SetValue(t, Convert.ChangeType(cell.NumericCellValue, instanceProperty.PropertyType), null);
                    }

                    //Set date value
                    else if (instanceProperty != null && instanceProperty.IsDateOrTime())
                    {
                        instanceProperty.SetValue(t, Convert.ChangeType(cell.DateCellValue, instanceProperty.PropertyType), null);
                    }

                    //Set boolean value
                    else if (instanceProperty != null && instanceProperty.IsBoolean())
                    {
                        instanceProperty.SetValue(t, Convert.ChangeType(cell.BooleanCellValue, instanceProperty.PropertyType), null);
                    }

                    //Else is string
                    else
                    {
                        instanceProperty.SetValue(t, Convert.ChangeType(cell.StringCellValue, instanceProperty.PropertyType), null);
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

        private void LoadStreamFromPath(string path)
        {
            using (_file = new MemoryStream())
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[file.Length];
                file.Read(bytes, 0, (int)file.Length);
                _file.Write(bytes, 0, (int)file.Length);
            }
        }

        public IExcelImportEngine AsExcel()
        {
            return this;
        }
    }
}
