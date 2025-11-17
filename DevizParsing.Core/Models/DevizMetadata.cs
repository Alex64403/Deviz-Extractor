using System;
using System.Collections.Generic;

namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Reprezintă metadatele de antet extrase dintr-un deviz (beneficiar, obiectiv, etc.).
    /// </summary>
    public class DevizMetadata
    {
        public string? Beneficiar { get; set; }
        public string? Executant { get; set; }
        public string? Proiectant { get; set; }
        public string? Obiectiv { get; set; }
        public string? Obiect { get; set; }
        public string? Deviz { get; set; }
        public string? StadiuFizic { get; set; }
        public string? SectiuneTehnica { get; set; }
        public string? SectiuneFinanciara { get; set; }
        public string? DataDocument { get; set; }
        public Dictionary<string, string> Extra { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
