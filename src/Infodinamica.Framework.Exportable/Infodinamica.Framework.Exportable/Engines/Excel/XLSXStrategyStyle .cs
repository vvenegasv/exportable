using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public class XLSXStrategyStyle : IExcelStyle
    {
        private XSSFWorkbook Workbook;

        public XLSXStrategyStyle(XSSFWorkbook workbook)
        {
            Workbook = workbook;
        }

        public ICellStyle GetCellStyle(IColor backColor, IColor borderColor, IFont font)
        {
            var cell = Workbook.CreateCellStyle();
            ((XSSFCellStyle)cell).SetFillForegroundColor((XSSFColor)backColor);
            ((XSSFCellStyle)cell).SetLeftBorderColor((XSSFColor)borderColor);
            ((XSSFCellStyle)cell).SetRightBorderColor((XSSFColor)borderColor);
            ((XSSFCellStyle)cell).SetTopBorderColor((XSSFColor)borderColor);
            ((XSSFCellStyle)cell).SetBottomBorderColor((XSSFColor)borderColor);
            cell.Alignment = HorizontalAlignment.Center;
            cell.BorderBottom = BorderStyle.Thin;
            cell.BorderLeft = BorderStyle.Thin;
            cell.BorderRight = BorderStyle.Thin;
            cell.BorderTop = BorderStyle.Thin;
            cell.FillPattern = FillPattern.SolidForeground;
            cell.VerticalAlignment = VerticalAlignment.Center;
            cell.SetFont(font);
            return cell;
        }

        public IFont GetFont(short fontSize, string fontName, IColor fontColor)
        {
            var font = Workbook.CreateFont();
            font.Boldweight = 100;
            ((XSSFFont)font).SetColor((XSSFColor)fontColor);
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            return font;
        }


    }
}
