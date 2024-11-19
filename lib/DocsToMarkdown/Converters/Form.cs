using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal sealed class Form() : BaseConverter("form")
{
    public override string? Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        var quizQuestions = htmlInput.SelectNodes(".//div[@class='quiz-question']");
        if (quizQuestions == null)
            return string.Empty;

        var sb = new StringBuilder();
        foreach (var item in quizQuestions)
        {
            ProcessQuestion(item.Element("div"), sb);
        }

        return sb.ToString();
    }

    private static void ProcessQuestion(HtmlNode question, StringBuilder sb)
    {
        if (question.GetAttributeValue("role", "") != "radiogroup")
            return;

        var questionText = question.Element("div").Element("p").InnerText;
        sb.AppendLine($"## {questionText}")
          .AppendLine();
        
        var answers = question.SelectNodes(".//label//p");
        foreach (var answer in answers)
        {
            sb.AppendLine("- [ ] " + answer.InnerText);
        }

        sb.AppendLine();
    }
}