namespace Markdig.Renderer.Docx.Blocks;

public class RowBlockRenderer : DocxObjectRenderer<RowBlock>
{
    public void Write(IDocxRenderer owner, IDocument document, List<RowBlock> rows)
    {
        int totalColumns = rows.Max(r => r.Count);

        var documentTable = document.Add(new Table(rows.Count, totalColumns));
        documentTable.Design = TableDesign.None;

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            for (int colIndex = 0; colIndex < row.Count; colIndex++)
            {
                var documentCell = documentTable.Rows.ElementAt(rowIndex).Cells[colIndex];
                var cellParagraph = documentCell.Paragraphs.First();

                if (row[colIndex] is NestedColumnBlock cell)
                {
                    foreach (var child in cell)
                    {
                        Write(child, owner, document, cellParagraph);
                        cellParagraph = documentCell.AddParagraph();
                    }
                }
                else
                {
                    var child = row[colIndex];
                    Write(child, owner, document, cellParagraph);
                }
            }
        }

        documentTable.AutoFit();
    }

    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, RowBlock obj)
    {
        // Not used.
        throw new NotImplementedException();
    }
}