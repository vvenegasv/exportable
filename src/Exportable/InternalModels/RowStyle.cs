namespace Exportable.InternalModels
{
    internal class RowStyle
    {
        public string FontColor { get; set; }
        public string FontName { get; set; }
        public short FontSize { get; set; }
        public string BorderColor { get; set; }
        public string BackColor { get; set; }



        public RowStyle()
        {

        }

        public RowStyle(string fontColor, string fontName, short fontSize, string borderColor, string backColor)
        {

            FontColor = fontColor;
            FontName = fontName;
            FontSize = fontSize;
            BorderColor = borderColor;
            BackColor = backColor;

        }
    }
}