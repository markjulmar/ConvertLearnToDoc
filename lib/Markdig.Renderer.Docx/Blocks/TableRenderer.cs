using System.Drawing;
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
                => tableColumnDefinition.Width != 0.0f && (int)tableColumnDefinition.Width != 1);

            List<double> columnWidths = null;
            if (hasColumnWidth)
            {
                // Force column widths to be evaluated.
                columnWidths = table.ColumnDefinitions
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
                if (firstRow)
                {
                    if (row.IsHeader)
                    {
                        documentTable.Properties.Design = AddCustomTableStyle(document);
                        documentTable.Properties.ConditionalFormatting = TableConditionalFormatting.FirstRow |
                                                                         TableConditionalFormatting.NoColumnBand;
                    }
                    else 
                        documentTable.Properties.Design = TableDesign.Normal;

                }

                firstRow = false;

                for (int colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cell = (TCell) row[colIndex];
                    var documentCell = documentTable.Rows.ElementAt(rowIndex).Cells[colIndex];

                    if (columnWidths?.Count > 0)
                    {
                        double width = columnWidths[colIndex];
                        documentCell.Properties.CellWidth = width == 0 
                            ? new TableElementWidth(TableWidthUnit.Auto, 0) 
                            : TableElementWidth.FromPercent(width);
                        documentCell.Properties.SetMargins(0);
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
                            cellParagraph.Properties.Alignment = alignment switch
                            {
                                TableColumnAlign.Left => Alignment.Left,
                                TableColumnAlign.Center => Alignment.Center,
                                TableColumnAlign.Right => Alignment.Right,
                                _ => cellParagraph.Properties.Alignment
                            };
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

            if (columnWidths == null || columnWidths.Count == 0)
            {
                documentTable.AutoFit();
            }
        }

        private static string AddCustomTableStyle(IDocument document)
        {
            const string docsTableStyle = "DocsTable";

            if (!document.Styles.Exists(docsTableStyle, StyleType.Table))
            {
                var style = document.Styles.Add(docsTableStyle, "Docs Table Style", StyleType.Table);

                var border = new Border(BorderStyle.Single, Uom.FromPoints(.5))
                {
                    Color = new ColorValue(Color.FromArgb(0x9C, 0xC2, 0xE5), ThemeColor.Accent5, 153)
                };

                style.ParagraphFormatting = new() { LineSpacingAfter = 0 };

                style.TableFormatting = new() { RowBands = 1 };
                style.TableFormatting.SetOutsideBorders(border);
                style.TableFormatting.SetInsideBorders(border);

                style.TableCellFormatting = new() { VerticalAlignment = VerticalAlignment.Center };

                var mainColor = new ColorValue(Color.FromArgb(0x5B, 0x9B, 0xD5), ThemeColor.Accent5);
                border = new Border(BorderStyle.Single, Uom.FromPoints(.5)) { Color = mainColor };

                style.TableStyles.Add(new TableStyle(TableStyleType.FirstRow)
                {
                    Formatting = new() { Bold = true, Color = new ColorValue(Color.White, ThemeColor.Background1) },
                    TableCellFormatting = new()
                    {
                        TopBorder = border,
                        BottomBorder = border,
                        LeftBorder = border,
                        RightBorder = border,
                        Shading = new()
                        {
                            Color = ColorValue.Auto,
                            Pattern = ShadePattern.Clear,
                            Fill = mainColor
                        }
                    }
                });

                style.TableStyles.Add(new TableStyle(TableStyleType.LastRow)
                {
                    Formatting = new() { Bold = true },
                    TableCellFormatting = new()
                    {
                        TopBorder = new Border(BorderStyle.Double, Uom.FromPoints(.5)) { Color = mainColor }
                    }
                });

                style.TableStyles.Add(new TableStyle(TableStyleType.FirstColumn)
                {
                    Formatting = new() { Bold = true },
                });

                style.TableStyles.Add(new TableStyle(TableStyleType.LastColumn)
                {
                    Formatting = new() { Bold = true },
                });

                var cellBand = new TableCellProperties
                {
                    Shading = new()
                    {
                        Color = ColorValue.Auto,
                        Pattern = ShadePattern.Clear,
                        Fill = new ColorValue(Color.FromArgb(0xDE, 0xEA, 0xF6), ThemeColor.Accent5, 51)
                    }
                };
                style.TableStyles.Add(new TableStyle(TableStyleType.BandedEvenRows) { TableCellFormatting = cellBand });
            }

            return docsTableStyle;
        }
    }
}