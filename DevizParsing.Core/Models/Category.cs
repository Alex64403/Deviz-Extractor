namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Încapsulează cantitatea, prețul unitar și totalul pentru o categorie de cost.
    /// </summary>
    public class Category
    {
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Total { get; set; }
    }
}
