using System.Collections.Generic;

namespace DevizParsing.Core.Models.Internal
{
    /// <summary>
    /// Structură internă ce reprezintă liniile de sumar identificate în foaie.
    /// </summary>
    public class SheetSummary
    {
        public Dictionary<string, SummaryLine> Categories { get; } = new Dictionary<string, SummaryLine>();
        public Dictionary<string, SummaryLine> ExtraTotals { get; } = new Dictionary<string, SummaryLine>();
        public SummaryLine? GrandTotal { get; set; }
        public SummaryLine? TotalQuantity { get; set; }
    }

    /// <summary>
    /// Descrie o linie de sumar găsită, de regulă, în partea inferioară a foii.
    /// </summary>
    public class SummaryLine
    {
        public int RowIndex { get; set; }
        public string Label { get; set; } = string.Empty;
        public decimal? Quantity { get; set; }
        public decimal? Total { get; set; }
    }
}
