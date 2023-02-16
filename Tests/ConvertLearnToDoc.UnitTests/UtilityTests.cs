using LearnDocUtils;

namespace ConvertLearnToDoc.UnitTests
{
    public class UtilityTests
    {
        [Fact]
        public void CheckConvertedCharactersAreStripped()
        {
            string text = "\r\nThis is a test\xA0of the\u200bemergency broadcast\u202fsystem.\r\n";
            string expected = "This is a test of theemergency broadcast system.";

            var result = DocToMarkdownRenderer.PostProcessMarkdown(text);
            Assert.Equal(expected, result);
        }
    }
}