using System.Collections.Generic;

namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Reprezintă un rând logic extras dintr-o foaie de deviz.
    /// </summary>
    public class RowItem
    {
        public string Order { get; set; } = string.Empty;
        public string Symbol { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string UnitOfMeasure { get; set; } = string.Empty;
        public Categories Categories { get; set; } = new Categories();
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal LineTotal { get; set; }
        public decimal ComputedLineTotal { get; set; }
        public decimal? SheetLineTotal { get; set; }
        public RowValidationInfo Validation { get; set; } = new RowValidationInfo();
        public string Notes { get; set; } = string.Empty;
    }
}
