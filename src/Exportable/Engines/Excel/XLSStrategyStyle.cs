using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Exportable.Engines.Excel
{
    internal class XLSStrategyStyle : IExcelStyle
    {

        private HSSFWorkbook Workbook;


        public XLSStrategyStyle(HSSFWorkbook workbook)
        {

            Workbook = workbook;
        }


        public ICellStyle GetCellStyle(IColor backColor, IColor borderColor, IFont font)
        {
            var cell = Workbook.CreateCellStyle();
            cell.FillForegroundColor = backColor.Indexed;
            cell.BottomBorderColor = borderColor.Indexed;
            cell.TopBorderColor = borderColor.Indexed;
            cell.LeftBorderColor = borderColor.Indexed;
            cell.RightBorderColor = borderColor.Indexed;
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
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            font.Color = fontColor.Indexed;
            return font;
        }


    }
}
