using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using TextFrame2 = Microsoft.Office.Interop.PowerPoint.TextFrame2;

namespace PowerPointSlideHtmlLayoutDemo;

class HtmlPageTextGenerator
{
    private readonly string _template;
    private readonly int _widthInPixels;
    private readonly int _heightInPixels;
    private readonly float _slideWidth;
    private readonly float _slideHeight;
    private readonly StringBuilder _shapeHtml = new();

    public HtmlPageTextGenerator(string template, int widthInPixels, int heightInPixels, float slideWidth, float slideHeight)
    {
        _template = template
            .Replace("$$width$$", widthInPixels.ToString())
            .Replace("$$height$$", heightInPixels.ToString());

        _widthInPixels = widthInPixels;
        _heightInPixels = heightInPixels;
        _slideWidth = slideWidth;
        _slideHeight = slideHeight;
    }

    public void AddInsertionPoint(Shape shape)
    {
        var styleBuilder = new StringBuilder();

        var textFrame = shape.TextFrame2;
        var marginLeft = GetHorizontalPixels(textFrame.MarginLeft);
        var marginRight = GetHorizontalPixels(textFrame.MarginRight);
        var marginTop = GetVerticalPixels(textFrame.MarginTop);
        var marginBottom = GetVerticalPixels(textFrame.MarginBottom);
        styleBuilder.Append($"padding: {marginTop}px {marginRight}px {marginBottom}px {marginLeft}px;");

        var verticalAlignment = GetVerticalAlignment(textFrame);
        styleBuilder.Append($"align-items: {verticalAlignment};");

        var textRange = textFrame.TextRange;
        var (horizontalAlignment, justifyContent) = GetHorizontalAlignment(textRange);
        styleBuilder.Append($"justify-content: {justifyContent};");

        switch (horizontalAlignment)
        {
            case HorizontalAlignment.Left:
                styleBuilder.Append($"left: {GetHorizontalPixels(shape.Left)}px;");
                styleBuilder.Append($"text-align: left;");
                break;
            case HorizontalAlignment.Center:
                styleBuilder.Append($"left: {GetHorizontalPixels(shape.Left + shape.Width / 2)}px;");
                styleBuilder.Append($"text-align: center;");
                styleBuilder.Append("transform: translate(-50%, 0);");
                break;
            case HorizontalAlignment.Right:
                styleBuilder.Append($"right: {GetHorizontalPixels(_slideWidth - (shape.Left + shape.Width))}px;");
                styleBuilder.Append($"text-align: right;");
                break;
        }
        styleBuilder.Append($"width: {GetHorizontalPixels(shape.Width)}px;");
        styleBuilder.Append($"top: {GetVerticalPixels(shape.Top)}px;");
        switch (textFrame.AutoSize)
        {
            case MsoAutoSize.msoAutoSizeNone:
            case MsoAutoSize.msoAutoSizeTextToFitShape: // not supported, so we'll keep the container size
                styleBuilder.Append($"width: {GetHorizontalPixels(shape.Width)}px;");
                styleBuilder.Append($"height: {GetVerticalPixels(shape.Height)}px;");
                break;
            case MsoAutoSize.msoAutoSizeShapeToFitText:
                styleBuilder.Append("width: auto;");
                styleBuilder.Append("height: auto;");
                break;
        }

        switch (textFrame.WordWrap)
        {
            case MsoTriState.msoTrue:
                styleBuilder.Append("white-space: pre-wrap;overflow-wrap: break-word;");
                break;
            case MsoTriState.msoFalse:
                styleBuilder.Append("white-space: pre;");
                break;
        }

        var font = textRange.Font;
        styleBuilder.Append($"font-family: '{font.Name}';");
        styleBuilder.Append($"font-size: {((font.Size * _widthInPixels) / _slideWidth).ToString(CultureInfo.InvariantCulture)}px;");
        if (font.Italic == MsoTriState.msoTrue)
        {
            styleBuilder.Append("font-style: italic;");
        }
        if (font.Bold == MsoTriState.msoTrue)
        {
            styleBuilder.Append("font-weight: bold;");
        }
        if (font.Allcaps == MsoTriState.msoTrue)
        {
            styleBuilder.Append("text-transform: uppercase");
        }

        var paragraphFormat = textRange.ParagraphFormat;
        if (paragraphFormat.LineRuleWithin == MsoTriState.msoTrue)
        {
            if (paragraphFormat.SpaceWithin!=1.0)
                styleBuilder.Append($"line-height: {paragraphFormat.SpaceWithin.ToString(CultureInfo.InvariantCulture)};");
        }
        else
        {
            var points = paragraphFormat.SpaceWithin;
            var pixels = (points * 96) / 72; // For HTML, 1pt = 1/72th of 1in, 1px = 1/96th of 1in (according to W3C, see https://www.w3.org/TR/css3-values/#absolute-lengths)
            styleBuilder.Append($"line-height: {pixels.ToString(CultureInfo.InvariantCulture)}px;");
        }
        Marshal.FinalReleaseComObject(paragraphFormat);

        string textDecoration = font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline
            ? "underline"
            : String.Empty;
        if (font.StrikeThrough == MsoTriState.msoTrue)
        {
            if (textDecoration.Length > 0)
                textDecoration += " ";
            textDecoration += "line-through";
        }
        if (textDecoration.Length > 0)
        {
            styleBuilder.Append($"text-decoration: {textDecoration};");
        }

        styleBuilder.Append($"color: {GetFontColor(font)};");

        var insertionPointText = GetInsertionPointText(shape);
        Marshal.FinalReleaseComObject(font);


        _shapeHtml.AppendLine($"<div class=\"shape\" style=\"{styleBuilder}\"><div>{insertionPointText}</div></div>");
    }

    private static string GetFontColor(Font2 font)
    {
        var fillFormat = font.Fill;
        var foreColor = fillFormat.ForeColor;
        var color = System.Drawing.ColorTranslator.ToHtml(System.Drawing.ColorTranslator.FromOle(foreColor.RGB));
        Marshal.FinalReleaseComObject(foreColor);
        Marshal.FinalReleaseComObject(fillFormat);
        return color;
    }

    private string GetInsertionPointText(Shape shape)
    {
        var textFrame = shape.TextFrame2;
        var textRange = textFrame.TextRange;
        var text = textRange.Text;
        Marshal.FinalReleaseComObject(textRange);
        Marshal.FinalReleaseComObject(textFrame);
        return text;
    }

    private string GetVerticalAlignment(TextFrame2 textFrame)
    {
        switch (textFrame.VerticalAnchor)
        {
            case MsoVerticalAnchor.msoAnchorTop:
                return "flex-start";
            case MsoVerticalAnchor.msoAnchorMiddle:
                return "center";
            case MsoVerticalAnchor.msoAnchorBottom:
                return "flex-end";
        }
        // others are not supported in this proof-of-concept
        return "flex-start";
    }

    enum HorizontalAlignment
    {
        Left,
        Center,
        Right
    }
    private (HorizontalAlignment, string) GetHorizontalAlignment(TextRange2 textRange)
    {
        // Note that this proof-of-concept code does not take right-to-left languages into account.
        var paragraphFormat = textRange.ParagraphFormat;
        switch (paragraphFormat.Alignment)
        {
            case MsoParagraphAlignment.msoAlignLeft:
                return (HorizontalAlignment.Left, "flex-start");
            case MsoParagraphAlignment.msoAlignCenter:
                return (HorizontalAlignment.Center, "center");
            case MsoParagraphAlignment.msoAlignRight:
                return (HorizontalAlignment.Right, "flex-end");
        }
        // Others are not supported
        return (HorizontalAlignment.Left, "flex-start");
    }


    public string GetHtmlText()
    {
        return _template.Replace("$$shapes$$", _shapeHtml.ToString());
    }

    public void AddInsertionPoints(IEnumerable<Shape> shapes)
    {
        foreach (var shape in shapes)
        {
            AddInsertionPoint(shape);
        }
    }

    private int GetVerticalPixels(float y)
    {
        return (int)Math.Round((_heightInPixels * y) / _slideHeight);
    }
    private int GetHorizontalPixels(float x)
    {
        return (int)Math.Round((_widthInPixels * x) / _slideWidth);
    }
}
