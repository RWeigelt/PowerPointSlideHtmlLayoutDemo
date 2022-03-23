using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointSlideHtmlLayoutDemo;

public static class SlideInsertionPointHelper
{
    public static Shape[] CollectInsertionPoints(Slide slide)
    {
        var shapes = slide.Shapes;
        var collectedItems = new List<Shape>();
        foreach (Shape shape in shapes)
        {
            if (shape.HasTable == MsoTriState.msoTrue)
            {
                CollectShapesInTable(shape.Table, collectedItems);
                continue;
            }
            if (shape.Type == MsoShapeType.msoGroup)
            {
                CollectGroupItems(shape.GroupItems, collectedItems);
                continue;
            }

            if (ContainsInsertionPoint(shape))
            {
                collectedItems.Add(shape);
            }
        }
        return collectedItems.ToArray();
    }

    public static void HideInsertionPoints(Shape[] insertionPointShapes)
    {
        foreach (var shape in insertionPointShapes)
        {
            var textFrame = shape.TextFrame2;
            var textRange = textFrame.TextRange;
            var font = textRange.Font;
            var fill = font.Fill;
            fill.Transparency = 1;
            shape.TextFrame.DeleteText();
        }
    }

    private static void CollectGroupItems(Microsoft.Office.Interop.PowerPoint.GroupShapes groupShapes, List<Shape> collectedShapes)
    {
        //
        // Implementation note: Even though grouped shapes can be nested inside the PowerPoint UI,
        // we get all shapes at once. So we don't need to implement recursive traversal.
        //
        foreach (Shape shape in groupShapes)
        {
            if (ContainsInsertionPoint(shape))
            {
                collectedShapes.Add(shape);
            }
        }
    }

    private static void CollectShapesInTable(Table table, List<Shape> collectedItems)
    {
        for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            for (int cellIndex = 1; cellIndex <= row.Cells.Count; cellIndex++)
            {
                var cell = row.Cells[cellIndex];
                var shape = cell.Shape;
                if (ContainsInsertionPoint(shape))
                {
                    collectedItems.Add(shape);
                }
            }
        }
    }

    private static bool ContainsInsertionPoint(Shape shape)
    {
        if (shape.HasTextFrame != MsoTriState.msoTrue)
            return false;

        var textFrame = shape.TextFrame;
        if (textFrame.HasText != MsoTriState.msoTrue)
            return false;

        var text = textFrame.TextRange.Text;
        var startIndex = text.IndexOf("{{", StringComparison.Ordinal);
        if (startIndex == -1)
            return false;
        var endIndex = text.IndexOf("}}", StringComparison.Ordinal);
        if (endIndex == -1)
            return false;
        return startIndex < endIndex;
    }

}