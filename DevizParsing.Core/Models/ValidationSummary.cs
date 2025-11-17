using System.Collections.Generic;

namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Centralizează totalurile de validare și problemele descoperite în timpul parsării.
    /// </summary>
    public class ValidationSummary
    {
        public decimal? GrandTotalFromSheet { get; set; }
        public bool GrandTotalMatchesSheet { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public Dictionary<string, SummaryCheck> CategorySummaries { get; set; } = new Dictionary<string, SummaryCheck>();
        public SummaryCheck? GrandTotalSummary { get; set; }
        public SummaryCheck? TotalQuantitySummary { get; set; }
        public List<RowIssue> RowIssues { get; set; } = new List<RowIssue>();
    }
}
