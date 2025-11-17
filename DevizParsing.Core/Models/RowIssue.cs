namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Descrie o problemă de validare asociată unui rând specific.
    /// </summary>
    public class RowIssue
    {
        public string Order { get; set; } = string.Empty;
        public string Symbol { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
