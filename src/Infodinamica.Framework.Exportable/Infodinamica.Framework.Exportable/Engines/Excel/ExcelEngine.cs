using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
    public abstract class ExcelEngine
    {
        protected HSSFWorkbook ExcelXls { get; set; }
        protected XSSFWorkbook ExcelXlsx { get; set; }
        protected ExcelVersion ExcelVersion;
        

        public ExcelEngine()
        {
            ExcelVersion = ExcelVersion.XLSX;
        }

        public void SetFormat(ExcelVersion version)
        {
            ExcelVersion = version;
        }

        protected void Write(ref MemoryStream stream)
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
        
        protected ICellStyle GetStyleWithFormat(ICellStyle baseStyle, string dataFormat)
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

        protected ICellStyle CreateCellStyle()
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

        protected IFont CreateFont()
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

        protected ISheet CreateSheet(string name)
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

        protected ISheet GetSheet(string name)
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

        protected IDataFormat CreateDataFormat()
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

        protected short GetBuiltIndDataFormat(string dataFormat)
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

        protected void ReadFromMemoryStream(MemoryStream stream)
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
