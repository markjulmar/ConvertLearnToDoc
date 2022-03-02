using MDTable = Markdig.Extensions.Tables.Table;

namespace Markdig.Renderer.Docx.Blocks;

public class ListRenderer : DocxObjectRenderer<ListBlock>
{
    private int currentLevel = -1;

    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, ListBlock block)
    {
        currentLevel++;
        try
        {
            NumberingDefinition nd = null;
            if (block.IsOrdered)
            {
                if (!int.TryParse(block.OrderedStart, out int startNumber))
                    startNumber = 1;

                if (startNumber > 1)
                {
                    // Backup and find the previous active list and see if this should fit in.
                    var foundList = document.Paragraphs.Reverse().FirstOrDefault(p => p.IsListItem() && p.GetNumberingFormat() == NumberingFormat.Numbered);
                    if (foundList != null)
                    {
                        nd = foundList.GetListNumberingDefinition();
                    }
                }

                nd ??= document.NumberingStyles.Create(NumberingFormat.Numbered, startNumber);
            }
            else
            {
                nd = document.NumberingStyles.Create(NumberingFormat.Bullet);
            }

            int count = 0;

            // ListBlock has a collection of ListItemBlock objects
            // ... which in turn contain paragraphs, tables, code blocks, etc.
            foreach (var listItem in block.Cast<ListItemBlock>())
            {
                if (count > 0 && currentParagraph != null)
                    currentParagraph = currentParagraph.AddParagraph();
                else if (currentParagraph == null)
                    currentParagraph = document.AddParagraph();
                currentParagraph.ListStyle(nd, currentLevel);

                count++;

                for (var index = 0; index < listItem.Count; index++)
                {
                    var childBlock = listItem[index];
                    if (index > 0)
                    {
                        if (childBlock is not MDTable)
                        {
                            // Create a new paragraph to hold this block.
                            // Unless it's a table - that gets appended to the prior paragraph.
                            currentParagraph = currentParagraph.AddParagraph().ListStyle();
                        }
                    }

                    Write(childBlock, owner, document, currentParagraph);

                    // Catch all the paragraphs added.
                    var paragraphs = document.Paragraphs.ToList();
                    int cp = paragraphs.IndexOf(currentParagraph);
                    if (cp >= 0)
                    {
                        for (int indentCheck = cp; indentCheck < paragraphs.Count; indentCheck++)
                        {
                            currentParagraph = paragraphs[indentCheck];
                            if (childBlock is MDTable && currentParagraph.Table != null)
                            {
                                currentParagraph.Table.Indent = Uom.FromInches(.5);
                            }
                            else if (currentParagraph.Properties.StyleName != "ListParagraph")
                            {
                                currentParagraph.Properties.LeftIndent = Uom.FromInches(.5);
                            }
                        }
                    }
                }
            }
        }
        finally
        {
            currentLevel--;
        }
    }
}