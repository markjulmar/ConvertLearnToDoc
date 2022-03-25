using Markdig.Extensions.Tables;
using Table = Markdig.Extensions.Tables.Table;
using TRow = Markdig.Extensions.Tables.TableRow;
using TCell = Markdig.Extensions.Tables.TableCell;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TableRenderer : DocxObjectRenderer<Table>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, Table table)
        {
            bool hasColumnWidth = table.ColumnDefinitions.Any(tableColumnDefinition 
                => tableColumnDefinition.Width != 0.0f && tableColumnDefinition.Width != 1.0f);

            var columnWidths = new List<double>();
            if (hasColumnWidth)
            {
                // Force column widths to be evaluated.
                _ = table.ColumnDefinitions
                    .Select(tableColumnDefinition => Math.Round(tableColumnDefinition.Width * 100) / 100)
                    .ToList();
            }

            int totalColumns = table.Max(tr => ((TRow) tr).Count);
            var documentTable = new DXPlus.Table(table.Count, totalColumns);
            if (currentParagraph != null)
            {
                currentParagraph.InsertAfter(documentTable);
            }
            else
            {
                document.Add(documentTable);
            }

            bool firstRow = true;

            for (var rowIndex = 0; rowIndex < table.Count; rowIndex++)
            {
                var row = (TRow) table[rowIndex];
                if (firstRow) {
                    documentTable.Design = row.IsHeader
                        ? TableDesign.Grid : TableDesign.None;
                }

                firstRow = false;

                for (int colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cell = (TCell) row[colIndex];
                    var documentCell = documentTable.Rows.ElementAt(rowIndex).Cells[colIndex];

                    if (columnWidths.Count > 0)
                    {
                        double width = columnWidths[colIndex];
                        documentCell.CellWidth = width == 0 
                            ? new TableWidth(TableWidthUnit.Auto, 0) 
                            : TableWidth.FromPercent(width);
                        documentCell.SetMargins(0);
                    }

                    var cellParagraph = documentCell.Paragraphs.First();
                    WriteChildren(cell, owner, document, cellParagraph);
                    
                    if (table.ColumnDefinitions.Count > 0)
                    {
                        var columnIndex = cell.ColumnIndex < 0 || cell.ColumnIndex >= table.ColumnDefinitions.Count
                            ? colIndex
                            : cell.ColumnIndex;
                        columnIndex = columnIndex >= table.ColumnDefinitions.Count
                            ? table.ColumnDefinitions.Count - 1
                            : columnIndex;
                        var alignment = table.ColumnDefinitions[columnIndex].Alignment;
                        if (alignment.HasValue)
                        {
                            switch (alignment)
                            {
                                case TableColumnAlign.Left:
                                    cellParagraph.Properties.Alignment = Alignment.Left;
                                    break;
                                case TableColumnAlign.Center:
                                    cellParagraph.Properties.Alignment = Alignment.Center;
                                    break;
                                case TableColumnAlign.Right:
                                    cellParagraph.Properties.Alignment = Alignment.Right;
                                    break;
                            }
                        }
                    }
                    
                    if (cell.ColumnSpan != 1)
                    {
                    }

                    if (cell.RowSpan != 1)
                    {
                    }
                }
            }

            if (columnWidths.Count == 0)
            {
                documentTable.AutoFit();
            }
        }
    }
}