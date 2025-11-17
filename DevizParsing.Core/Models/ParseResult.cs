using System;
using System.Collections.Generic;

namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Reprezintă rezultatul complet al parsării, incluzând rândurile, totalurile și informațiile de validare.
    /// </summary>
    public class ParseResult
    {
        public string SourceFile { get; set; } = string.Empty;
        public string Sheet { get; set; } = string.Empty;
        public DevizMetadata Metadata { get; set; } = new DevizMetadata();
        public List<RowItem> Rows { get; set; } = new List<RowItem>();
        public Totals ComputedTotals { get; set; } = new Totals();
        public ValidationSummary Validation { get; set; } = new ValidationSummary();
    }
}
