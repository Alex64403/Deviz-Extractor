namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Totaluri agregate pe toate rândurile parcurse din deviz.
    /// </summary>
    public class Totals
    {
        public decimal Materials { get; set; }
        public decimal Labor { get; set; }
        public decimal Equipment { get; set; }
        public decimal Transport { get; set; }
        public decimal GrandTotal { get; set; }
    public decimal DirectGrandTotal { get; set; }
    public decimal LeafGrandTotal { get; set; }
    public decimal AllRowsGrandTotal { get; set; }
    public decimal OtherDirectCosts { get; set; }
    public decimal TotalCheltuieliDirecte { get; set; }
    public decimal CheltuieliIndirecte { get; set; }
    public decimal Profit { get; set; }
    public decimal TotalDevizFaraTvaInitial { get; set; }
    public decimal TotalDevizFaraTvaFinal { get; set; }
    public decimal TotalGeneralFaraTva { get; set; }
    public decimal Vat { get; set; }
        public decimal MaterialsQuantity { get; set; }
        public decimal LaborQuantity { get; set; }
        public decimal EquipmentQuantity { get; set; }
        public decimal TransportQuantity { get; set; }
        public decimal OverallQuantity { get; set; }
        public decimal GrandTotalFromSheetLines { get; set; }
        public decimal GrandTotalFromComputedLines { get; set; }
    }
}
