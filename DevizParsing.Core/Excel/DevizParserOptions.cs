using System.Collections.Generic;

namespace DevizParsing.Core.Excel
{
    /// <summary>
    /// Opțiuni care controlează modul de lucru al parserului de foi de deviz.
    /// </summary>
    public class DevizParserOptions
    {
    /// <summary>
    /// Selectează unul dintre profilele poziționale predefinite folosite când inferența antetului eșuează.
    /// </summary>
    public DevizParserProfile Profile { get; set; } = DevizParserProfile.Intersoft;

    /// <summary>
    /// Indici de coloană expliciți (opțional) folosiți în locul profilului implicit.
    /// </summary>
        public Dictionary<DevizColumnRole, int>? CustomFallbackColumns { get; set; }
            = null;

    /// <summary>
    /// Numărul de rânduri analizate atunci când se caută rândul de antet.
    /// </summary>
    public int HeaderScanLimit { get; set; } = 20;

    /// <summary>
    /// Forțează parsarea pozițională chiar dacă anteturile par lizibile.
    /// </summary>
        public bool ForcePositionalFallback { get; set; }
            = false;

    /// <summary>
    /// Diferența absolută permisă când se compară totaluri de cantități sau valori.
    /// </summary>
    public decimal ValidationTolerance { get; set; } = 0.05m;
    }
}
