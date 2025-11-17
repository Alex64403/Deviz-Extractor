namespace DevizParsing.Core.Excel
{
    /// <summary>
    /// Definește rolurile logice asociate coloanelor dintr-o foaie de deviz.
    /// </summary>
    public enum DevizColumnRole
    {
        Order,
        Symbol,
        Name,
        UnitOfMeasure,
        Quantity,
        UnitPrice,
        LineTotal,
        MaterialsQuantity,
        MaterialsUnitPrice,
        LaborQuantity,
        LaborUnitPrice,
        EquipmentQuantity,
        EquipmentUnitPrice,
        TransportQuantity,
        TransportUnitPrice
    }
}
