using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public interface IExcelStyle
    {
        IFont GetFont(short fontSize, string fontName, IColor fontColor);
        ICellStyle GetCellStyle(IColor backColor, IColor borderColor, IFont font);
    }
}
