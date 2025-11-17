using System.Collections.Generic;

namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Stochează metricele de validare pentru un rând parse-at.
    /// </summary>
    public class RowValidationInfo
    {
        public decimal CategoriesTotal { get; set; }
        public decimal ComputedLineTotal { get; set; }
        public decimal? SheetLineTotal { get; set; }
        public decimal DifferenceToLineTotal { get; set; }
        public decimal DifferenceToComputedLineTotal { get; set; }
        public decimal? DifferenceToSheetLineTotal { get; set; }
        public bool CategoriesMatchLineTotal { get; set; } = true;
        public bool ComputedMatchesSheet { get; set; } = true;
        public List<string> Issues { get; set; } = new List<string>();
    }
}
