using Markdig.Extensions.TaskLists;

namespace Markdig.Renderer.Docx.Inlines;

public class TaskListRenderer : DocxObjectRenderer<TaskList>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TaskList taskListEntry)
    {
        currentParagraph.AddText(taskListEntry.Checked ? "❎" : "⬜ ");
    }
}