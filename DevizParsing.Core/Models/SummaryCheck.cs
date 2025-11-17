namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Reprezintă un total de control (cantitate/valoare) extras din foaia sursă.
    /// </summary>
    public class SummaryCheck
    {
        public string Key { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
        public int RowIndex { get; set; }
        public decimal ComputedTotal { get; set; }
        public decimal? SheetTotal { get; set; }
        public bool TotalMatches { get; set; } = true;
        public decimal ComputedQuantity { get; set; }
        public decimal? SheetQuantity { get; set; }
        public bool QuantityMatches { get; set; } = true;
    }
}
