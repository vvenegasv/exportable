using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Infodinamica.Framework.Core.Extensions.Common;
using Infodinamica.Framework.Exportable.Resources;
using Infodinamica.Framework.Exportable.Tools;
using Microsoft.VisualBasic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public class ExcelExportEngine : ExcelEngine, IExcelExportEngine
    {
        private readonly IDictionary<string, object> _sheets;

        public ExcelExportEngine()
        {
            _sheets = new Dictionary<string, object>();
        }

        public void AddData<T>(IList<T> data) where T : class
        {
            var sheetName = "Sheet " + _sheets.Count + 1;
            _sheets.Add(new KeyValuePair<string, object>(sheetName, data));
        }
        
        public void AddData<T>(IList<T> data, string sheetName) where T : class
        {
            _sheets.Add(new KeyValuePair<string, object>(sheetName, data));
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
            if(StringMethods.IsNullOrWhiteSpace(fileExtension))
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
                var customProperties = MetadataHelper.GetExportableMetadatas(genericType);

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
        
        


        private void CreateSheet(KeyValuePair<string, object> excelSheet)
        {
            ISheet hoja = CreateSheet(excelSheet.Key);
            IRow header = hoja.CreateRow(0);

            var genericType = MetadataHelper.GetGenericType(excelSheet.Value);
            var headerFormat = MetadataHelper.GetHeadersFormat(genericType);

            ICellStyle cell = CreateCellStyle(headerFormat);

            int cellCount = 0;
            foreach (var column in MetadataHelper.GetHeadersName(genericType))
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
            var customProperties = MetadataHelper.GetExportableMetadatas(genericType);

            //Set all styles and save in a list for future usage
            customProperties
                .Select(x => new { x.Position, x.Format })
                .Distinct()
                .ToList()
                .ForEach(cp =>
                {
                    styles.Add(cp.Position, GetStyleWithFormat(cellStyleFila, cp.Format));
                });

            //Adds default format
            styles.Add(-1, GetStyleWithFormat(cellStyleFila, DefaultValues.DefaultNumericFormat));
            styles.Add(-2, GetStyleWithFormat(cellStyleFila, DefaultValues.DefaultDateFormat));


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
                                var cellValue = (string)propValue;
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

        public IExcelExportEngine AsExcel()
        {
            return this;
        }

    }

}
