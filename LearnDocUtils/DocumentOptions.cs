namespace LearnDocUtils
{
    public class DocumentOptions
    {
        /// <summary>
        /// Keep intermediate files during conversion
        /// </summary>
        public bool Debug { get; set; }

        /// <summary>
        /// If a notebook is in a unit, render it inline in the document
        /// </summary>
        public bool EmbedNotebookContent { get; set; }
    }
}
