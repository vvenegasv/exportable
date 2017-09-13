using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Resources;
using Exportable.Attribute;
using Exportable.Helpers;
using Exportable.InternalModels;
using Exportable.Models;
using Exportable.Resources;
using Microsoft.VisualBasic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Exportable.Engines.Excel
{
    /// <summary>
    /// Clase para exportar a Excel
    /// </summary>
    public class ExcelExportEngine : ExcelEngine, IExcelExportEngine
    {
        private readonly IDictionary<string, object> _sheets;
        protected IDictionary<string, IDictionary<string, string>> _newColumnNames;
        protected IDictionary<string, HashSet<string>> _columnsToIgnore;


        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelExportEngine()
        {
            _sheets = new Dictionary<string, object>();
            _newColumnNames = new Dictionary<string, IDictionary<string, string>>();
            _columnsToIgnore = new Dictionary<string, HashSet<string>>();
        }

        public string AddData<TModel>(IEnumerable<TModel> data) where TModel : class
        {
            var sheetName = "Sheet " + _sheets.Count + 1;
            _sheets.Add(new KeyValuePair<string, object>(sheetName, data));
            return sheetName;
        }
        
        public string AddData<TModel>(IEnumerable<TModel> data, string sheetName) where TModel : class
        {
            _sheets.Add(new KeyValuePair<string, object>(sheetName, data));
            return sheetName;
        }

        public void AddIgnoreColumns<TModel>(string datasetKey, Expression<Func<TModel, object>> propertyExpression)
            where TModel : class
        {
            var propInfo = GetPropertyInfo(propertyExpression);
            if (_columnsToIgnore.ContainsKey(datasetKey))
            {
                var ignoredColumns = _columnsToIgnore[datasetKey];
                if (!ignoredColumns.Contains(propInfo.Name))
                    ignoredColumns.Add(propInfo.Name);
            }
            else
            {
                var ignoredColumns = new HashSet<string>();
                ignoredColumns.Add(propInfo.Name);
                _columnsToIgnore.Add(datasetKey, ignoredColumns);
            }
        }

        public void AddColumnsNames<TModel>(string datasetKey, Expression<Func<TModel, object>> propertyExpression, string name)
            where TModel : class
        {

            var propInfo = GetPropertyInfo(propertyExpression);

            if (_newColumnNames.ContainsKey(datasetKey))
            {
                var newNamesDictionary = _newColumnNames[datasetKey];
                newNamesDictionary[propInfo.Name] = name;
            }
            else
            {
                var newNamesDictionary = new Dictionary<string, string>()
                {
                    {propInfo.Name, name}
                };
                _newColumnNames.Add(datasetKey, newNamesDictionary);
            }
        }

        public MemoryStream Export()
        {
            if(ExcelVersion == ExcelVersion.XLSX)
                ExcelXlsx = new XSSFWorkbook();

            if (ExcelVersion == ExcelVersion.XLS)
                ExcelXls = new HSSFWorkbook();
            

            foreach (var sheet in _sheets)
            {
                CreateSheet(sheet);
                AddDataToSheet(sheet);
            }
            
            var memoryStream = new MemoryStream();
            Write(ref memoryStream);

            //Fix memorystream closed by NPOI for XLSX format
            if (!memoryStream.CanRead)
            {
                MemoryStream newMemoryStream = new MemoryStream(memoryStream.ToArray());
                newMemoryStream.Position = 0;
                return newMemoryStream;
            }
            else
            {
                memoryStream.Position = 0;
                return memoryStream;
                
            }

            //ExcelXls.Write(memoryStream);
            //memoryStream.Position = 0;
        }

        public void Export(string path)
        {
            var fileExtension = Path.GetExtension(path);
            if(string.IsNullOrWhiteSpace(fileExtension))
                throw new Exception(ErrorMessage.FileExtensionNotProvided);

            if(!fileExtension.Equals(".xls") && !fileExtension.Equals(".xlsx"))
                throw new Exception(string.Format(ErrorMessage.InvalidFileExtension, fileExtension));

            var excelInMemory = Export();
            using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bytes = new byte[excelInMemory.Length];
                excelInMemory.Read(bytes, 0, (int)excelInMemory.Length);
                file.Write(bytes, 0, bytes.Length);
                excelInMemory.Close();
            }
        }

        public IDictionary<string, string> RunBusinessRules()
        {
            IDictionary<string, string> errors = new Dictionary<string, string>();

            //Check position with same id
            foreach (var sheet in _sheets)
            {
                var genericType = MetadataHelper.GetGenericType(sheet.Value);
                var customProperties = GetExportableMetadatas(sheet.Key, genericType);

                var positions = customProperties
                    .GroupBy(x => x.Position)
                    .Select(y => new
                    {
                        Position = y.Key,
                        Quantity = y.Count(x => true)
                    });

                if (positions.Any(x => x.Quantity > 1))
                    errors.Add(new KeyValuePair<string, string>(sheet.Key, string.Format(ErrorMessage.RepeatedPosition, sheet.Key)));
            }

            //Check excel version
            if(ExcelVersion!= ExcelVersion.XLS && ExcelVersion!= ExcelVersion.XLSX)
                errors.Add("Excel Version", ErrorMessage.Excel_BadVersion);

            return errors;
        }

        public IExcelExportEngine AsExcel()
        {
            return this;
        }
        


        private void CreateSheet(KeyValuePair<string, object> excelSheet)
        {
            ISheet hoja = CreateSheet(excelSheet.Key);
            IRow header = hoja.CreateRow(0);

            var genericType = MetadataHelper.GetGenericType(excelSheet.Value);
            var headerFormat = MetadataHelper.GetHeadersFormat(genericType);

            ICellStyle cell = CreateCellStyle(headerFormat);

            int cellCount = 0;
            foreach (var column in GetHeadersName(excelSheet.Key, genericType))
            {
                header.CreateCell(cellCount);
                header.GetCell(cellCount).CellStyle = cell;
                header.GetCell(cellCount).SetCellValue(column);
                cellCount += 1;
            }

        }

        private void AddDataToSheet(KeyValuePair<string, object> excelSheet)
        {
            ISheet sheet = GetSheet(excelSheet.Key);
            ICellStyle cellStyleFila = CreateCellStyle();
            IDictionary<int, ICellStyle> styles = new Dictionary<int, ICellStyle>();
            var genericType = MetadataHelper.GetGenericType(excelSheet.Value);
            var customProperties = GetExportableMetadatas(excelSheet.Key, genericType);

            //Set all styles and save in a list for future usage
            customProperties
                .Select(x => new { x.Position, x.Format, x.FieldValueType })
                .Distinct()
                .ToList()
                .ForEach(cp =>
                {
                    if (cp.FieldValueType == FieldValueType.Date || cp.FieldValueType == FieldValueType.Numeric)
                        styles.Add(cp.Position, GetStyleWithFormat(cellStyleFila, cp.Format, true));
                    else
                        styles.Add(cp.Position, GetStyleWithFormat(cellStyleFila, cp.Format));
                });

            //Adds default format
            styles.Add(-1, GetStyleWithFormat(cellStyleFila, DefaultValues.DefaultNumericFormat, true));
            styles.Add(-2, GetStyleWithFormat(cellStyleFila, DefaultValues.DefaultDateFormat, true));


            //Al rows
            int rowCount = 1;
            foreach (var row in MetadataHelper.GetArrayData(excelSheet.Value))
            {
                //All cells
                int cellCount = 0;
                IRow fila = sheet.CreateRow(rowCount);

                //Order properties by position attribute
                foreach (var property in customProperties.OrderBy(cp => cp.Position))
                {
                    //Create row
                    fila.CreateCell(cellCount);
                    fila.GetCell(cellCount).CellStyle = cellStyleFila;

                    //Get value
                    var propValue = row.GetType().GetProperty(property.Name).GetValue(row, null);

                    //If value is null, set blank text an iterate again
                    if (propValue == null)
                    {
                        fila.GetCell(cellCount).SetCellValue("");
                        cellCount++;
                        continue;
                    }

                    switch (property.FieldValueType)
                    {
                        case FieldValueType.Date:
                            var dateCellValue = Convert.ToDateTime(propValue);
                            fila.GetCell(cellCount).SetCellValue(dateCellValue);
                            fila.GetCell(cellCount).CellStyle = styles.FirstOrDefault(x => x.Key == property.Position).Value;
                            break;

                        case FieldValueType.Numeric:
                            var numericCellValue = Convert.ToDouble(propValue);
                            fila.GetCell(cellCount).SetCellValue(numericCellValue);
                            fila.GetCell(cellCount).CellStyle = styles.FirstOrDefault(x => x.Key == property.Position).Value;
                            fila.GetCell(cellCount).SetCellType(CellType.Numeric);
                            break;

                        case FieldValueType.Text:
                            var stringCellValue = propValue.ToString();
                            fila.GetCell(cellCount).SetCellValue(stringCellValue);
                            fila.GetCell(cellCount).SetCellType(CellType.String);
                            break;

                        case FieldValueType.Bool:
                            var boolCellValue = Convert.ToBoolean(propValue);
                            fila.GetCell(cellCount).SetCellValue(boolCellValue);
                            fila.GetCell(cellCount).SetCellType(CellType.Boolean);
                            break;


                        case FieldValueType.Any:
                            if (propValue is bool)
                            {
                                var cellValue = Convert.ToBoolean(propValue);
                                fila.GetCell(cellCount).SetCellValue(cellValue);
                            }
                            else if (Information.IsDate(propValue))
                            {
                                var cellValue = Convert.ToDateTime(propValue);
                                fila.GetCell(cellCount).SetCellValue(cellValue);
                                fila.GetCell(cellCount).CellStyle = styles.FirstOrDefault(x => x.Key == -2).Value;
                            }
                            else if (Information.IsNumeric(propValue))
                            {
                                var cellValue = Convert.ToDouble(propValue);
                                fila.GetCell(cellCount).SetCellValue(cellValue);
                                fila.GetCell(cellCount).CellStyle = styles.FirstOrDefault(x => x.Key == -1).Value;
                                fila.GetCell(cellCount).SetCellType(CellType.Numeric);
                            }
                            else
                            {
                                //https://github.com/vvenegasv/exportable/issues/2
                                //Thanks to nesreeen
                                var cellValue = Convert.ToString(propValue);
                                fila.GetCell(cellCount).SetCellValue(cellValue);
                            }
                            break;
                    }

                    cellCount += 1;
                }
                rowCount += 1;
            }

            //Set autowidth
            int col = 0;
            foreach (var cell in sheet.GetRow(0).Cells)
            {
                sheet.AutoSizeColumn(col);
                col++;
            }
        }

        private IEnumerable<string> GetHeadersName(string sheetName, Type type)
        {
            var headerNames = new Dictionary<string, int>();
            var headersNameWithoutAttribute = new List<string>();
            IDictionary<string, string> newColumnsNames = null;
            HashSet<string> ignoredColumns = null;

            if (_newColumnNames.ContainsKey(sheetName))
                newColumnsNames = _newColumnNames[sheetName];
            if (_columnsToIgnore.ContainsKey(sheetName))
                ignoredColumns = _columnsToIgnore[sheetName];

            foreach (var property in type.GetProperties())
            {
                //Skip if column is ignored
                if (ignoredColumns != null && ignoredColumns.Contains(property.Name))
                    continue;

                //Get the exportable attribute
                var exportableAttribute = property.GetCustomAttribute<ExportableAttribute>();
                if (exportableAttribute != null)
                {
                    //Skip to next item if column need be ignored
                    if (exportableAttribute.IsIgnored)
                        continue;

                    //Set column name in runtime and skip to next item
                    if (newColumnsNames!=null && newColumnsNames.ContainsKey(property.Name))
                        exportableAttribute.HeaderName = newColumnsNames[property.Name];

                    //Try to get header's name from resource file
                    else if (exportableAttribute.ResourceType != null
                        && !string.IsNullOrWhiteSpace(exportableAttribute.HeaderName))
                    {
                        // Create a resource manager to retrieve resources.
                        var rm = new ResourceManager(exportableAttribute.ResourceType);

                        // Retrieve the value of the string resource named "welcome".
                        // The resource manager will retrieve the value of the  
                        // localized resource using the caller's current culture setting.
                        exportableAttribute.HeaderName = rm.GetString(exportableAttribute.HeaderName);
                    }

                    headerNames.Add(exportableAttribute.HeaderName, exportableAttribute.Position);
                }
                else
                {
                    //Because header names are ordered by position,
                    //We need to add header names without attribute in another
                    //list to add in the final
                    headersNameWithoutAttribute.Add(property.Name);
                }
            }


            //Add columns without attribute
            if (headersNameWithoutAttribute.Any())
            {
                var index = headerNames.Any() ? headerNames.Max(x => x.Value) + 1 : 1;
                foreach (var item in headersNameWithoutAttribute)
                {
                    headerNames.Add(item, index);
                    index++;
                }
            }


            //Returns the headers name ordered by position
            return headerNames.OrderBy(x => x.Value).Select(x => x.Key);
        }

        private IEnumerable<Metadata> GetExportableMetadatas(string sheetName, Type type)
        {
            var exportableMetadatas = new List<Metadata>();
            var exportableWithoutMetadata = new List<Metadata>();
            HashSet<string> ignoredColumns = null;
            
            if (_columnsToIgnore.ContainsKey(sheetName))
                ignoredColumns = _columnsToIgnore[sheetName];

            foreach (var property in type.GetProperties())
            {
                //Skip if column is ignored
                if (ignoredColumns != null && ignoredColumns.Contains(property.Name))
                    continue;

                var exportableAttribute = property.GetCustomAttribute<ExportableAttribute>();

                //Check if has ExportableAttribute. If it hasn't, get data
                //directly from property values
                if (exportableAttribute != null)
                {
                    //Skip to next item if column need be ignored
                    if (!exportableAttribute.IsIgnored)
                        exportableMetadatas.Add(new Metadata
                        {
                            Name = property.Name,
                            Position = exportableAttribute.Position,
                            Format = exportableAttribute.Format,
                            FieldValueType = exportableAttribute.TypeValue,
                            DefaultForNullOrInvalidValues = string.Empty
                        });
                }
                else
                {
                    //If it havent ExportableAttribute, it will be added to another list because they will be in the last records
                    exportableWithoutMetadata.Add(new Metadata
                    {
                        Name = property.Name,
                        Position = 0,
                        Format = null,
                        FieldValueType = FieldValueType.Any,
                        DefaultForNullOrInvalidValues = string.Empty
                    });
                }
            }

            //get biggest position
            int index = 0;
            if (exportableMetadatas.Any())
            {
                index = exportableMetadatas.Select(exp => exp.Position).Max();
                index++;
            }

            //Add elements without ExportableAttribute to returning list
            exportableWithoutMetadata.ForEach(exp =>
            {
                exp.Position = index;
                exportableMetadatas.Add(exp);
                index++;
            });

            return exportableMetadatas;
        }
    }

}
