namespace Markdig.Renderer.Docx.Blocks;

/// <summary>
/// Renders a :::row::: extension.
/// </summary>
public class RowBlockRenderer : DocxObjectRenderer<RowBlock>
{
    public void Write(IDocxRenderer owner, IDocument document, List<RowBlock> rows)
    {
        int totalColumns = rows.Max(r => r.Count);
        var documentTable = document.Add(new Table(rows.Count, totalColumns));
        documentTable.Properties.Design = TableDesign.Normal;

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            for (int colIndex = 0; colIndex < row.Count; colIndex++)
            {
                var documentCell = documentTable.Rows.ElementAt(rowIndex).Cells[colIndex];
                var cellParagraph = documentCell.Paragraphs.First();

                if (row[colIndex] is NestedColumnBlock cell)
                {
                    int index = 0;
                    foreach (var child in cell)
                    {
                        if (index++ > 0)
                            cellParagraph = documentCell.AddParagraph();
                        Write(child, owner, document, cellParagraph);
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

    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, RowBlock block)
    {
        Table documentTable;
        TableRow row;

        if (currentParagraph?.Table == null)
        {
            int totalColumns = -1;

            // Backup to parent, grab all RowBlock elements adjacent to this one and get the max columns.
            var container = block.Parent;
            if (container != null)
            {
                var children = container.ToList();
                int pos = children.IndexOf(block);
                if (pos >= 0)
                {
                    for (; pos < children.Count; pos++)
                    {
                        if (children[pos] is RowBlock rb)
                            totalColumns = Math.Max(totalColumns, rb.Count);
                        else break;
                    }
                }
            }
            
            if (totalColumns == -1)
                totalColumns = Math.Max(1, block.Count);

            documentTable = new Table(1, totalColumns)
            {
                Properties =
                {
                    Design = TableDesign.Normal
                }
            };

            if (currentParagraph != null)
            {
                currentParagraph.InsertAfter(documentTable);
            }
            else
            {
                document.Add(documentTable);
            }

            row = documentTable.Rows[0];
        }
        else
        {
            documentTable = currentParagraph.Table;
            row = documentTable.AddRow();
        }

        for (int colIndex = 0; colIndex < block.Count; colIndex++)
        {
            var documentCell = row.Cells[colIndex];
            var cellParagraph = documentCell.Paragraphs.First();

            if (block[colIndex] is NestedColumnBlock cell)
            {
                int index = 0;
                foreach (var child in cell)
                {
                    if (index++ > 0)
                        cellParagraph = documentCell.AddParagraph();
                    Write(child, owner, document, cellParagraph);
                }
            }
            else
            {
                var child = block[colIndex];
                Write(child, owner, document, cellParagraph);
            }
        }
        
        documentTable.AutoFit();
    }
}