using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;
using Infodinamica.Framework.Core.Extensions.Common;
using Infodinamica.Framework.Exportable.Resources;
using Infodinamica.Framework.Exportable.Tools;
using Microsoft.VisualBasic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.Record;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    /// <summary>
    /// Clase base para Excel
    /// </summary>
    public abstract class ExcelEngine
    {
        internal HSSFWorkbook ExcelXls { get; set; }
        internal XSSFWorkbook ExcelXlsx { get; set; }
        internal ExcelVersion ExcelVersion;
        private short? _palleteColorSize = null;


        public ExcelEngine()
        {
            ExcelVersion = ExcelVersion.XLSX;
        }

        public void SetFormat(ExcelVersion version)
        {
            ExcelVersion = version;
        }
        
        internal void Write(ref MemoryStream stream)
        {
            if (stream == null)
                stream = new MemoryStream();

            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    ExcelXls.Write(stream);
                    break;
                case ExcelVersion.XLSX:
                    ExcelXlsx.Write(stream);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
        
        internal ICellStyle GetStyleWithFormat(ICellStyle baseStyle, string dataFormat)
        {
            ICellStyle newStyle = CreateCellStyle();
            newStyle.CloneStyleFrom(baseStyle);

            if (StringMethods.IsNullOrWhiteSpace(dataFormat))
                return newStyle;

            // check if this is a built-in format
            var builtinFormatId = GetBuiltIndDataFormat(dataFormat);

            if (builtinFormatId != -1)
            {
                newStyle.DataFormat = builtinFormatId;
            }
            else
            {
                // not a built-in format, so create a new one
                var newDataFormat = CreateDataFormat();
                newStyle.DataFormat = newDataFormat.GetFormat(dataFormat);
            }

            return newStyle;
        }
        
        internal ICellStyle CreateCellStyle()
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return ExcelXls.CreateCellStyle();
                case ExcelVersion.XLSX:
                    return ExcelXlsx.CreateCellStyle();
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
        
        internal ICellStyle CreateCellStyle(RowStyle rowStyle)
        {
            IColor borderColor = CreateColor(rowStyle.BorderColor);
            IColor backColor = CreateColor(rowStyle.BackColor);
            IFont font = CreateFont(rowStyle);
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    var CellStyleXLS = new XLSStrategyStyle(ExcelXls);
                    return CellStyleXLS.GetCellStyle(backColor, borderColor, font);
                case ExcelVersion.XLSX:
                    var CellStyleXLSX = new XLSXStrategyStyle(ExcelXlsx);
                    return CellStyleXLSX.GetCellStyle(backColor, borderColor, font);
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
        
        internal IFont CreateFont()
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return ExcelXls.CreateFont();
                case ExcelVersion.XLSX:
                    return ExcelXlsx.CreateFont();
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
        
        internal IFont CreateFont(RowStyle rowStyle)
        {
            IColor color = CreateColor(rowStyle.FontColor);
            switch (ExcelVersion){
                case ExcelVersion.XLS:
                    var fontExcelXls = new XLSStrategyStyle(ExcelXls);
                    return fontExcelXls.GetFont(rowStyle.FontSize, rowStyle.FontName, color);
                case ExcelVersion.XLSX:
                    var fontExcelXlsx = new XLSXStrategyStyle(ExcelXlsx);
                    return fontExcelXlsx.GetFont(rowStyle.FontSize, rowStyle.FontName, color);
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
        
        internal ISheet CreateSheet(string name)
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return ExcelXls.CreateSheet(name);
                case ExcelVersion.XLSX:
                    return ExcelXlsx.CreateSheet(name);
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }

        internal ISheet GetSheet(string name)
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return ExcelXls.GetSheet(name);
                case ExcelVersion.XLSX:
                    return ExcelXlsx.GetSheet(name);
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }

        internal IDataFormat CreateDataFormat()
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return ExcelXls.CreateDataFormat();
                case ExcelVersion.XLSX:
                    return ExcelXlsx.CreateDataFormat();
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }

        internal IColor CreateColor(string htmlColor)
        {
            Color color = ColorTranslator.FromHtml(htmlColor);
            byte[] rgbColor = new byte[3] { color.R, color.G, color.B };

            switch (ExcelVersion)
            {
                case Excel.ExcelVersion.XLS:

                    var palette = ExcelXls.GetCustomPalette();
                    HSSFColor hfcolor = null;

                    //Limit palette
                    if (_palleteColorSize >= 63)
                    {
                        hfcolor = palette.FindColor(color.R, color.G, color.B);
                        if (hfcolor == null)
                        {
                            hfcolor = palette.FindSimilarColor(color.R, color.G, color.B);
                        }
                        _palleteColorSize++;
                        return hfcolor;
                    }
                    else
                    {
                        if (!_palleteColorSize.HasValue)
                            _palleteColorSize = PaletteRecord.FIRST_COLOR_INDEX;
                        else
                            _palleteColorSize++;
                        palette.SetColorAtIndex(_palleteColorSize.Value, color.R, color.G, color.B);
                        hfcolor = palette.GetColor((short)_palleteColorSize);
                        return hfcolor;
                    }

                case Excel.ExcelVersion.XLSX:
                    return new XSSFColor(color);
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }

        internal short GetBuiltIndDataFormat(string dataFormat)
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    return HSSFDataFormat.GetBuiltinFormat(dataFormat);
                    break;
                case ExcelVersion.XLSX:
                    return new XSSFDataFormat(new StylesTable()).GetFormat(dataFormat);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }

        }

        internal void ReadFromMemoryStream(MemoryStream stream)
        {
            switch (ExcelVersion)
            {
                case ExcelVersion.XLS:
                    ExcelXls = new HSSFWorkbook(stream);
                    break;
                case ExcelVersion.XLSX:
                    ExcelXlsx = new XSSFWorkbook(stream);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
    }
}
