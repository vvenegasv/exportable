using System;
using System.Drawing;
using System.IO;
using Exportable.Resources;
using Exportable.Tools;
using Infodinamica.Framework.Core.Extensions.Common;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace Exportable.Engines.Excel
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
            return GetStyleWithFormat(baseStyle, dataFormat, false);
        }


        internal ICellStyle GetStyleWithFormat(ICellStyle baseStyle, string dataFormat, bool isNumberOrDate)
        {
            ICellStyle cellStyle = this.CreateCellStyle();
            cellStyle.CloneStyleFrom(baseStyle);

            //Se valida que tenga formato
            if (StringMethods.IsNullOrWhiteSpace(dataFormat))
                return cellStyle;

            //Las fechas las forzamos a utilizar un nuevo formato y no uno predefinido
            if (isNumberOrDate)
            {
                IDataFormat dataFormat2 = this.CreateDataFormat();
                cellStyle.DataFormat = dataFormat2.GetFormat(dataFormat);
                return cellStyle;

            } else
            {
                short builtIndDataFormat = this.GetBuiltIndDataFormat(dataFormat);
                if (builtIndDataFormat != -1)
                {
                    cellStyle.DataFormat = builtIndDataFormat;
                }
                else
                {
                    IDataFormat dataFormat2 = this.CreateDataFormat();
                    cellStyle.DataFormat = dataFormat2.GetFormat(dataFormat);
                }
                return cellStyle;
            }
        }

        internal ICellStyle CreateCellStyle()
        {
            ICellStyle result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    result = this.ExcelXls.CreateCellStyle();
                    break;
                case ExcelVersion.XLSX:
                    result = this.ExcelXlsx.CreateCellStyle();
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal ICellStyle CreateCellStyle(RowStyle rowStyle)
        {
            IColor borderColor = this.CreateColor(rowStyle.BorderColor);
            IColor backColor = this.CreateColor(rowStyle.BackColor);
            IFont font = this.CreateFont(rowStyle);
            ICellStyle cellStyle;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                {
                    XLSStrategyStyle xLSStrategyStyle = new XLSStrategyStyle(this.ExcelXls);
                    cellStyle = xLSStrategyStyle.GetCellStyle(backColor, borderColor, font);
                    break;
                }
                case ExcelVersion.XLSX:
                {
                    XLSXStrategyStyle xLSXStrategyStyle = new XLSXStrategyStyle(this.ExcelXlsx);
                    cellStyle = xLSXStrategyStyle.GetCellStyle(backColor, borderColor, font);
                    break;
                }
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return cellStyle;
        }

        internal IFont CreateFont()
        {
            IFont result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    result = this.ExcelXls.CreateFont();
                    break;
                case ExcelVersion.XLSX:
                    result = this.ExcelXlsx.CreateFont();
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal IFont CreateFont(RowStyle rowStyle)
        {
            IColor fontColor = this.CreateColor(rowStyle.FontColor);
            IFont font;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                {
                    XLSStrategyStyle xLSStrategyStyle = new XLSStrategyStyle(this.ExcelXls);
                    font = xLSStrategyStyle.GetFont(rowStyle.FontSize, rowStyle.FontName, fontColor);
                    break;
                }
                case ExcelVersion.XLSX:
                {
                    XLSXStrategyStyle xLSXStrategyStyle = new XLSXStrategyStyle(this.ExcelXlsx);
                    font = xLSXStrategyStyle.GetFont(rowStyle.FontSize, rowStyle.FontName, fontColor);
                    break;
                }
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return font;
        }

        internal ISheet CreateSheet(string name)
        {
            ISheet result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    result = this.ExcelXls.CreateSheet(name);
                    break;
                case ExcelVersion.XLSX:
                    result = this.ExcelXlsx.CreateSheet(name);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal ISheet GetSheet(string name)
        {
            ISheet sheet;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    sheet = this.ExcelXls.GetSheet(name);
                    break;
                case ExcelVersion.XLSX:
                    sheet = this.ExcelXlsx.GetSheet(name);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return sheet;
        }

        internal IDataFormat CreateDataFormat()
        {
            IDataFormat result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    result = this.ExcelXls.CreateDataFormat();
                    break;
                case ExcelVersion.XLSX:
                    result = this.ExcelXlsx.CreateDataFormat();
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal IColor CreateColor(string htmlColor)
        {
            Color color = ColorTranslator.FromHtml(htmlColor);
            byte[] array = new byte[]
            {
                color.R,
                color.G,
                color.B
            };
            IColor result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                {
                    HSSFPalette customPalette = this.ExcelXls.GetCustomPalette();
                    if (this._palleteColorSize >= 63)
                    {
                        HSSFColor hSSFColor = customPalette.FindColor(color.R, color.G, color.B);
                        if (hSSFColor == null)
                        {
                            hSSFColor = customPalette.FindSimilarColor(color.R, color.G, color.B);
                        }
                        short? palleteColorSize = this._palleteColorSize;
                        this._palleteColorSize = (palleteColorSize.HasValue
                            ? new short?((short)(palleteColorSize.GetValueOrDefault() + 1))
                            : null);
                        result = hSSFColor;
                    }
                    else
                    {
                        if (!this._palleteColorSize.HasValue)
                        {
                            this._palleteColorSize = new short?(8);
                        }
                        else
                        {
                            short? palleteColorSize = this._palleteColorSize;
                            this._palleteColorSize = (palleteColorSize.HasValue
                                ? new short?((short)(palleteColorSize.GetValueOrDefault() + 1))
                                : null);
                        }
                        customPalette.SetColorAtIndex(this._palleteColorSize.Value, color.R, color.G, color.B);
                        HSSFColor hSSFColor = customPalette.GetColor(this._palleteColorSize.Value);
                        result = hSSFColor;
                    }
                    break;
                }
                case ExcelVersion.XLSX:
                    result = new XSSFColor(color);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal short GetBuiltIndDataFormat(string dataFormat)
        {
            short result;
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    result = HSSFDataFormat.GetBuiltinFormat(dataFormat);
                    break;
                case ExcelVersion.XLSX:
                    result = new XSSFDataFormat(new StylesTable()).GetFormat(dataFormat);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
            return result;
        }

        internal void ReadFromMemoryStream(MemoryStream stream)
        {
            switch (this.ExcelVersion)
            {
                case ExcelVersion.XLS:
                    this.ExcelXls = new HSSFWorkbook(stream);
                    break;
                case ExcelVersion.XLSX:
                    this.ExcelXlsx = new XSSFWorkbook(stream);
                    break;
                default:
                    throw new Exception(ErrorMessage.Excel_BadVersion);
            }
        }
    }
}
