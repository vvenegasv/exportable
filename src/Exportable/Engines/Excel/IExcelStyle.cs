using NPOI.SS.UserModel;

namespace Exportable.Engines.Excel
{
    internal interface IExcelStyle
    {
        IFont GetFont(short fontSize, string fontName, IColor fontColor);
        ICellStyle GetCellStyle(IColor backColor, IColor borderColor, IFont font);
    }
}
