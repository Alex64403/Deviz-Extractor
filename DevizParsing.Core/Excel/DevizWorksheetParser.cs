using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using DevizParsing.Core.Models;
using DevizParsing.Core.Models.Internal;

namespace DevizParsing.Core.Excel
{
    /// <summary>
    /// Parsează foi Excel de tip deviz într-o structură <see cref="ParseResult"/> cu metadate de validare.
    /// </summary>
    public class DevizWorksheetParser
    {
        private static readonly string[] CandidatiAntetOrdine = { "număr ordine", "nr crt", "nr.", "nr", "ordine", "poz" };
        private static readonly string[] CandidatiAntetSimbol = { "simbol", "cod", "cod lucr.", "cod lucrare", "cod activitate", "cod articol" };
        private static readonly string[] CandidatiAntetDenumire = { "nume", "denumire", "descriere", "lucrare", "denumire activitate" };
        private static readonly string[] CandidatiAntetUnitate = { "um", "u.m.", "u/m", "unitate", "unitate de masura" };
        private static readonly string[] CandidatiCantMateriale = { "mat cant", "material cant", "cant material", "materiale cant", "mat. cant", "cant.mat" };
        private static readonly string[] CandidatiPretMateriale = { "mat pret", "material pret", "pret material", "materiale pret", "mat. pret" };
        private static readonly string[] CandidatiCantManopera = { "manopera cant", "manoperă cant", "manopera cantitate", "manopera" };
        private static readonly string[] CandidatiPretManopera = { "manopera pret", "manoperă pret", "manopera pret/unit", "manopera pret" };
        private static readonly string[] CandidatiCantUtilaje = { "utilaj cant", "utilaje cant", "utilaj" };
        private static readonly string[] CandidatiPretUtilaje = { "utilaj pret", "utilaje pret" };
        private static readonly string[] CandidatiCantTransport = { "transport cant", "transport" };
        private static readonly string[] CandidatiPretTransport = { "transport pret" };
        private static readonly string[] CandidatiCantGenerala = { "cantitatea", "cantitate", "cant.", "cant", "quantity", "qty" };
        private static readonly string[] CandidatiPretUnitarGeneral = { "pretul unitar", "prețul unitar", "pret unitar", "pret unit", "pret unit.", "preț unitar" };
        private static readonly string[] CandidatiTotalLinie = { "total", "valoare", "valoare totală", "valoare total", "total line", "suma" };

        private static readonly Dictionary<string, string[]> TokeniRezumatCategorii = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "materials", new[] { "material" } },
            { "labor", new[] { "manopera", "manoperă", "manopera" } },
            { "equipment", new[] { "utilaj", "utilaje" } },
            { "transport", new[] { "transport" } }
        };

        private static readonly Dictionary<string, string> NumeAfisareCategorii = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "materials", "Materiale" },
            { "labor", "Manoperă" },
            { "equipment", "Utilaje" },
            { "transport", "Transport" }
        };

        private const int MetadataScanLimit = 50;

        private static readonly char[] MetadataSeparators = { ':', '–', '-', '=' };

        private static readonly (string Key, string[] Aliases)[] MetadataAliasMap =
        {
            ("Beneficiar", new[] { "beneficiar" }),
            ("Executant", new[] { "executant" }),
            ("Proiectant", new[] { "proiectant" }),
            ("Obiectiv", new[] { "obiectiv" }),
            ("Obiect", new[] { "obiectul", "obiect" }),
            ("Deviz", new[] { "deviz", "denumiredevizului", "denumiredeviz", "devizul" }),
            ("StadiuFizic", new[] { "stadiulfizic", "stadiufizic" }),
            ("SectiuneTehnica", new[] { "sectiuneatehnica", "sectiunatehnica" }),
            ("SectiuneFinanciara", new[] { "sectiunafinanciara", "sectiuneafinanciara" }),
            ("DataDocument", new[] { "dataintocmirii", "datain", "data" })
        };

        private readonly DevizParserOptions _optiuni;

        /// <summary>
        /// Creează un parser configurat cu opțiunile furnizate sau cu valori implicite.
        /// </summary>
        /// <param name="options">Opțiunile de parsare; dacă sunt omise se folosesc valori implicite.</param>
        public DevizWorksheetParser(DevizParserOptions? options = null)
        {
            _optiuni = options ?? new DevizParserOptions();
        }

        /// <summary>
        /// Parsează prima foaie din fișierul Excel indicat.
        /// </summary>
        /// <param name="filePath">Calea către fișierul Excel.</param>
        /// <returns>Rezultatul parsării, incluzând rânduri și validări.</returns>
        public ParseResult Parse(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
            }

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();
            return ParseWorksheet(worksheet, Path.GetFileName(filePath));
        }

    /// <summary>
    /// Parsează o foaie Excel deja încărcată într-un <see cref="ParseResult"/> structururat.
    /// </summary>
    /// <param name="worksheet">Foaia Excel ce trebuie analizată.</param>
    /// <param name="sourceFile">Numele fișierului sursă pentru referință în rezultate.</param>
    /// <returns>Rezultatul parsării foii.</returns>
        public ParseResult ParseWorksheet(IXLWorksheet foaie, string fisierSursa)
        {
            var rezultat = new ParseResult
            {
                SourceFile = fisierSursa,
                Sheet = foaie.Name
            };

            var ultimulRand = foaie.LastRowUsed()?.RowNumber() ?? 0;
            var hartaAntet = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var randAntet = DetectHeaderRow(foaie, hartaAntet, ultimulRand);

            if (randAntet == -1)
            {
                randAntet = 1;
                AddValidationError(rezultat.Validation, "Randul de antet nu a fost detectat; se foloseste plasarea pozitionala");
            }

            rezultat.Metadata = ExtractMetadata(foaie, randAntet);

            int? colOrder = FindColumn(foaie, randAntet, hartaAntet, CandidatiAntetOrdine);
            int? colSymbol = FindColumn(foaie, randAntet, hartaAntet, CandidatiAntetSimbol);
            int? colName = FindColumn(foaie, randAntet, hartaAntet, CandidatiAntetDenumire);
            int? colUnit = FindColumn(foaie, randAntet, hartaAntet, CandidatiAntetUnitate);

            int? colQuantity = FindColumn(foaie, randAntet, hartaAntet, CandidatiCantGenerala);
            int? colUnitPrice = FindColumn(foaie, randAntet, hartaAntet, CandidatiPretUnitarGeneral);
            int? colLineTotal = FindColumn(foaie, randAntet, hartaAntet, CandidatiTotalLinie);

            int? colMatQty = FindColumn(foaie, randAntet, hartaAntet, CandidatiCantMateriale);
            int? colMatPrice = FindColumn(foaie, randAntet, hartaAntet, CandidatiPretMateriale);
            int? colLabQty = FindColumn(foaie, randAntet, hartaAntet, CandidatiCantManopera);
            int? colLabPrice = FindColumn(foaie, randAntet, hartaAntet, CandidatiPretManopera);
            int? colEqQty = FindColumn(foaie, randAntet, hartaAntet, CandidatiCantUtilaje);
            int? colEqPrice = FindColumn(foaie, randAntet, hartaAntet, CandidatiPretUtilaje);
            int? colTrQty = FindColumn(foaie, randAntet, hartaAntet, CandidatiCantTransport);
            int? colTrPrice = FindColumn(foaie, randAntet, hartaAntet, CandidatiPretTransport);

            var coloaneFallback = IaColoaneFallback();
            var trebuieFallbackPozitional = _optiuni.ForcePositionalFallback || TrebuieFallbackPozitional(foaie, randAntet, ultimulRand, colName, colSymbol);

            if (trebuieFallbackPozitional)
            {
                colOrder = IaFallback(coloaneFallback, DevizColumnRole.Order);
                colSymbol = IaFallback(coloaneFallback, DevizColumnRole.Symbol);
                colName = IaFallback(coloaneFallback, DevizColumnRole.Name);
                colUnit = IaFallback(coloaneFallback, DevizColumnRole.UnitOfMeasure);
                colQuantity = IaFallback(coloaneFallback, DevizColumnRole.Quantity);
                colUnitPrice = IaFallback(coloaneFallback, DevizColumnRole.UnitPrice);
                colLineTotal = IaFallback(coloaneFallback, DevizColumnRole.LineTotal);
                colMatQty = IaFallback(coloaneFallback, DevizColumnRole.MaterialsQuantity);
                colMatPrice = IaFallback(coloaneFallback, DevizColumnRole.MaterialsUnitPrice);
                colLabQty = IaFallback(coloaneFallback, DevizColumnRole.LaborQuantity);
                colLabPrice = IaFallback(coloaneFallback, DevizColumnRole.LaborUnitPrice);
                colEqQty = IaFallback(coloaneFallback, DevizColumnRole.EquipmentQuantity);
                colEqPrice = IaFallback(coloaneFallback, DevizColumnRole.EquipmentUnitPrice);
                colTrQty = IaFallback(coloaneFallback, DevizColumnRole.TransportQuantity);
                colTrPrice = IaFallback(coloaneFallback, DevizColumnRole.TransportUnitPrice);
            }

            var randCurent = randAntet + 1;
            var ultimulRandCuDate = randAntet;
            RowItem? randPrincipalAnterior = null;

            while (randCurent <= ultimulRand)
            {
                var primaCelula = foaie.Cell(randCurent, colOrder ?? 1);
                if ((primaCelula == null || primaCelula.IsEmpty()) && (colName == null || foaie.Cell(randCurent, colName.Value).IsEmpty()))
                {
                    var allEmpty = true;
                    for (var probe = 0; probe < 6; probe++)
                    {
                        if (!foaie.Row(randCurent + probe).IsEmpty())
                        {
                            allEmpty = false;
                            break;
                        }
                    }

                    if (allEmpty)
                    {
                        break;
                    }
                }

                var textOrdine = GetCellString(foaie, randCurent, colOrder);
                var celuleFolosite = foaie.Row(randCurent).CellsUsed().ToList();
                var textCombinat = string.Join(" ", celuleFolosite.Select(c => c.GetString().Trim().ToLowerInvariant()));
                var eRandCategorie = false;
                string? cheieCategorie = null;
                if (string.IsNullOrWhiteSpace(textOrdine))
                {
                    if (textCombinat.Contains("material")) { eRandCategorie = true; cheieCategorie = "materials"; }
                    else if (textCombinat.Contains("manopera") || textCombinat.Contains("manoper")) { eRandCategorie = true; cheieCategorie = "labor"; }
                    else if (textCombinat.Contains("utilaj")) { eRandCategorie = true; cheieCategorie = "equipment"; }
                    else if (textCombinat.Contains("transport")) { eRandCategorie = true; cheieCategorie = "transport"; }
                }

                if (eRandCategorie && randPrincipalAnterior != null)
                {
                    AplicaDefalcareCategorie(randPrincipalAnterior, cheieCategorie!, celuleFolosite);
                    ultimulRandCuDate = Math.Max(ultimulRandCuDate, randCurent);
                    randCurent++;
                    continue;
                }

                var randNou = new RowItem
                {
                    Order = textOrdine,
                    Symbol = GetCellString(foaie, randCurent, colSymbol),
                    Name = GetCellString(foaie, randCurent, colName),
                    UnitOfMeasure = GetCellString(foaie, randCurent, colUnit)
                };

                randNou.Quantity = GetCellDecimal(foaie, randCurent, colQuantity) ?? 0m;
                randNou.UnitPrice = GetCellDecimal(foaie, randCurent, colUnitPrice) ?? 0m;

                randNou.Categories.Materials.Quantity = GetCellDecimal(foaie, randCurent, colMatQty) ?? 0m;
                randNou.Categories.Materials.UnitPrice = GetCellDecimal(foaie, randCurent, colMatPrice) ?? 0m;
                randNou.Categories.Materials.Total = NormalizeTotal(randNou.Categories.Materials.Quantity, randNou.Categories.Materials.UnitPrice, null);

                randNou.Categories.Labor.Quantity = GetCellDecimal(foaie, randCurent, colLabQty) ?? 0m;
                randNou.Categories.Labor.UnitPrice = GetCellDecimal(foaie, randCurent, colLabPrice) ?? 0m;
                randNou.Categories.Labor.Total = NormalizeTotal(randNou.Categories.Labor.Quantity, randNou.Categories.Labor.UnitPrice, null);

                randNou.Categories.Equipment.Quantity = GetCellDecimal(foaie, randCurent, colEqQty) ?? 0m;
                randNou.Categories.Equipment.UnitPrice = GetCellDecimal(foaie, randCurent, colEqPrice) ?? 0m;
                randNou.Categories.Equipment.Total = NormalizeTotal(randNou.Categories.Equipment.Quantity, randNou.Categories.Equipment.UnitPrice, null);

                randNou.Categories.Transport.Quantity = GetCellDecimal(foaie, randCurent, colTrQty) ?? 0m;
                randNou.Categories.Transport.UnitPrice = GetCellDecimal(foaie, randCurent, colTrPrice) ?? 0m;
                randNou.Categories.Transport.Total = NormalizeTotal(randNou.Categories.Transport.Quantity, randNou.Categories.Transport.UnitPrice, null);

                var totalLinieDirect = GetCellDecimal(foaie, randCurent, colLineTotal);
                randNou.SheetLineTotal = totalLinieDirect;
                randNou.ComputedLineTotal = NormalizeTotal(randNou.Quantity, randNou.UnitPrice, null);
                randNou.LineTotal = totalLinieDirect ?? randNou.ComputedLineTotal;

                if (string.IsNullOrWhiteSpace(randNou.Name))
                {
                    randCurent++;
                    continue;
                }

                rezultat.Rows.Add(randNou);
                randPrincipalAnterior = randNou;
                ultimulRandCuDate = randCurent;

                randCurent++;
            }

            foreach (var rand in rezultat.Rows)
            {
                var totalCategorii = rand.Categories.Materials.Total + rand.Categories.Labor.Total + rand.Categories.Equipment.Total + rand.Categories.Transport.Total;
                if (totalCategorii == 0m && rand.LineTotal != 0m)
                {
                    rand.Categories.Materials.Total = rand.LineTotal;
                    if (rand.Categories.Materials.Quantity == 0m) rand.Categories.Materials.Quantity = rand.Quantity;
                    if (rand.Categories.Materials.UnitPrice == 0m && rand.Quantity != 0m)
                    {
                        rand.Categories.Materials.UnitPrice = Math.Round(rand.LineTotal / rand.Quantity, 4);
                    }
                }
            }

            VerificaRanduri(rezultat);

            rezultat.ComputedTotals.Materials = rezultat.Rows.Sum(x => x.Categories.Materials.Total);
            rezultat.ComputedTotals.Labor = rezultat.Rows.Sum(x => x.Categories.Labor.Total);
            rezultat.ComputedTotals.Equipment = rezultat.Rows.Sum(x => x.Categories.Equipment.Total);
            rezultat.ComputedTotals.Transport = rezultat.Rows.Sum(x => x.Categories.Transport.Total);

            var totalGeneralRanduri = rezultat.Rows.Sum(x => x.LineTotal);
            var totalDirect = rezultat.Rows.Where(IsTopLevelRow).Sum(x => x.LineTotal);
            var totalFrunze = totalGeneralRanduri - totalDirect;

            rezultat.ComputedTotals.AllRowsGrandTotal = totalGeneralRanduri;
            rezultat.ComputedTotals.DirectGrandTotal = totalDirect;
            rezultat.ComputedTotals.LeafGrandTotal = totalFrunze;
            rezultat.ComputedTotals.GrandTotal = totalDirect;

            rezultat.ComputedTotals.MaterialsQuantity = rezultat.Rows.Sum(x => x.Categories.Materials.Quantity);
            rezultat.ComputedTotals.LaborQuantity = rezultat.Rows.Sum(x => x.Categories.Labor.Quantity);
            rezultat.ComputedTotals.EquipmentQuantity = rezultat.Rows.Sum(x => x.Categories.Equipment.Quantity);
            rezultat.ComputedTotals.TransportQuantity = rezultat.Rows.Sum(x => x.Categories.Transport.Quantity);
            rezultat.ComputedTotals.OverallQuantity = rezultat.Rows.Sum(x => x.Quantity);
            rezultat.ComputedTotals.GrandTotalFromSheetLines = rezultat.Rows.Sum(x => x.SheetLineTotal ?? 0m);
            rezultat.ComputedTotals.GrandTotalFromComputedLines = rezultat.Rows.Sum(x => x.ComputedLineTotal);

            var randStartRezumat = Math.Min(ultimulRand, Math.Max(ultimulRandCuDate + 1, randAntet + 1));
            var sumarFoaie = ScanSummaryArea(foaie, randStartRezumat, ultimulRand);

            ApplyAdditionalSummaryTotals(rezultat, sumarFoaie);

            var totalGeneralFoaie = sumarFoaie.GrandTotal?.Total;
            if (!totalGeneralFoaie.HasValue)
            {
                totalGeneralFoaie = LocateFallbackGrandTotal(foaie, ultimulRand);
            }

            if (totalGeneralFoaie.HasValue)
            {
                rezultat.Validation.GrandTotalFromSheet = totalGeneralFoaie.Value;
                rezultat.Validation.GrandTotalMatchesSheet = AreClose(rezultat.ComputedTotals.GrandTotal, totalGeneralFoaie.Value);
                if (!rezultat.Validation.GrandTotalMatchesSheet)
                {
                    AddValidationError(rezultat.Validation, $"Totalul general calculat {rezultat.ComputedTotals.GrandTotal} difera de totalul din foaie {totalGeneralFoaie.Value}");
                }
            }

            ApplySummaryToValidation(rezultat, sumarFoaie);
            return rezultat;
        }

        private DevizMetadata ExtractMetadata(IXLWorksheet worksheet, int headerRow)
        {
            var metadata = new DevizMetadata();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            var scanLimit = headerRow > 1 ? Math.Min(headerRow - 1, MetadataScanLimit) : Math.Min(lastRow, MetadataScanLimit);
            if (scanLimit <= 0)
            {
                scanLimit = Math.Min(lastRow, MetadataScanLimit);
            }

            for (var rowNumber = 1; rowNumber <= scanLimit; rowNumber++)
            {
                var row = worksheet.Row(rowNumber);
                var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
                if (lastColumn == 0)
                {
                    continue;
                }

                var maxColumn = Math.Min(lastColumn, 20);
                for (var column = 1; column <= maxColumn; column++)
                {
                    var cell = row.Cell(column);
                    if (cell.IsEmpty())
                    {
                        continue;
                    }

                    var text = cell.GetString();
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (TryExtractMetadataKeyValue(row, column, text, out var key, out var value))
                    {
                        ApplyMetadataValue(metadata, key, value);
                    }
                }
            }

            return metadata;
        }

        private bool TryExtractMetadataKeyValue(IXLRow row, int column, string cellText, out string key, out string value)
        {
            key = string.Empty;
            value = string.Empty;

            var sanitized = cellText.Trim();
            if (string.IsNullOrEmpty(sanitized))
            {
                return false;
            }

            string? resolvedKey = null;
            string? resolvedValue = null;

            var split = sanitized.Split(MetadataSeparators, 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (split.Length == 2)
            {
                resolvedKey = ResolveMetadataKey(split[0]);
                if (resolvedKey != null)
                {
                    resolvedValue = CleanMetadataValue(split[1]);
                    if (string.IsNullOrEmpty(resolvedValue))
                    {
                        resolvedValue = FindNeighborMetadataValue(row, column);
                    }
                }
            }
            else
            {
                var trimmedKeyCandidate = sanitized.TrimEnd(MetadataSeparators);
                resolvedKey = ResolveMetadataKey(trimmedKeyCandidate);
                if (resolvedKey != null)
                {
                    resolvedValue = FindNeighborMetadataValue(row, column);
                }
            }

            if (resolvedKey != null && !string.IsNullOrEmpty(resolvedValue))
            {
                key = resolvedKey;
                value = resolvedValue;
                return true;
            }

            return false;
        }

        private static string? FindNeighborMetadataValue(IXLRow row, int column)
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? column;
            var maxColumn = Math.Min(lastColumn, column + 6);
            for (var col = column + 1; col <= maxColumn; col++)
            {
                var neighbor = row.Cell(col);
                if (neighbor.IsEmpty())
                {
                    continue;
                }

                var text = neighbor.GetString();
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (ResolveMetadataKey(text) != null)
                {
                    break;
                }

                var cleaned = CleanMetadataValue(text);
                if (!string.IsNullOrEmpty(cleaned))
                {
                    return cleaned;
                }
            }

            return null;
        }

        private static string CleanMetadataValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim().Trim('_');
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            trimmed = trimmed.Replace("_", string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            var builder = new StringBuilder(trimmed.Length);
            var previousSpace = false;
            foreach (var ch in trimmed)
            {
                if (char.IsWhiteSpace(ch))
                {
                    if (!previousSpace)
                    {
                        builder.Append(' ');
                        previousSpace = true;
                    }
                }
                else
                {
                    builder.Append(ch);
                    previousSpace = false;
                }
            }

            var cleaned = builder.ToString().Trim();
            return cleaned;
        }

        private static string? ResolveMetadataKey(string? rawKey)
        {
            if (string.IsNullOrWhiteSpace(rawKey))
            {
                return null;
            }

            var normalized = NormalizeMetadataKey(rawKey);
            if (string.IsNullOrEmpty(normalized))
            {
                return null;
            }

            foreach (var (key, aliases) in MetadataAliasMap)
            {
                foreach (var alias in aliases)
                {
                    if (string.Equals(normalized, alias, StringComparison.OrdinalIgnoreCase))
                    {
                        return key;
                    }
                }
            }

            return null;
        }

        private static string NormalizeMetadataKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(normalized.Length);
            foreach (var ch in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category == UnicodeCategory.NonSpacingMark)
                {
                    continue;
                }

                if (char.IsLetterOrDigit(ch))
                {
                    builder.Append(char.ToLowerInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private static void ApplyMetadataValue(DevizMetadata metadata, string key, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            value = value.Trim();
            if (value.Length == 0)
            {
                return;
            }

            switch (key)
            {
                case "Beneficiar":
                    if (string.IsNullOrEmpty(metadata.Beneficiar)) metadata.Beneficiar = value;
                    break;
                case "Executant":
                    if (string.IsNullOrEmpty(metadata.Executant)) metadata.Executant = value;
                    break;
                case "Proiectant":
                    if (string.IsNullOrEmpty(metadata.Proiectant)) metadata.Proiectant = value;
                    break;
                case "Obiectiv":
                    if (string.IsNullOrEmpty(metadata.Obiectiv)) metadata.Obiectiv = value;
                    break;
                case "Obiect":
                    if (string.IsNullOrEmpty(metadata.Obiect)) metadata.Obiect = value;
                    break;
                case "Deviz":
                    if (string.IsNullOrEmpty(metadata.Deviz)) metadata.Deviz = value;
                    break;
                case "StadiuFizic":
                    if (string.IsNullOrEmpty(metadata.StadiuFizic)) metadata.StadiuFizic = value;
                    break;
                case "SectiuneTehnica":
                    if (string.IsNullOrEmpty(metadata.SectiuneTehnica)) metadata.SectiuneTehnica = value;
                    break;
                case "SectiuneFinanciara":
                    if (string.IsNullOrEmpty(metadata.SectiuneFinanciara)) metadata.SectiuneFinanciara = value;
                    break;
                case "DataDocument":
                    if (string.IsNullOrEmpty(metadata.DataDocument)) metadata.DataDocument = value;
                    break;
                default:
                    metadata.Extra[key] = value;
                    break;
            }
        }

    /// <summary>
    /// Detectează rândul de antet și construiește maparea inițială a coloanelor.
    /// </summary>
    /// <param name="worksheet">Foaia în care se caută antetul.</param>
    /// <param name="headerMap">Dicționarul populat cu etichete de coloane.</param>
    /// <param name="lastRow">Ultimul rând utilizat în foaie.</param>
    /// <returns>Indexul rândului de antet sau -1 dacă nu a fost găsit.</returns>
    private int DetectHeaderRow(IXLWorksheet worksheet, Dictionary<string, int> headerMap, int lastRow)
        {
            var limitaInspectie = lastRow == 0 ? _optiuni.HeaderScanLimit : Math.Min(_optiuni.HeaderScanLimit, lastRow);
            for (var rowNumber = 1; rowNumber <= limitaInspectie; rowNumber++)
            {
                var row = worksheet.Row(rowNumber);
                var nonEmpty = row.CellsUsed().Select(c => c.GetString().Trim().ToLowerInvariant()).ToList();
                if (nonEmpty.Count == 0)
                {
                    continue;
                }

                var matches = nonEmpty.Count(cellText =>
                    CandidatiAntetOrdine.Any(h => cellText.Contains(h)) ||
                    CandidatiAntetSimbol.Any(h => cellText.Contains(h)) ||
                    CandidatiAntetDenumire.Any(h => cellText.Contains(h)) ||
                    CandidatiAntetUnitate.Any(h => cellText.Contains(h)) ||
                    CandidatiCantMateriale.Any(h => cellText.Contains(h)) ||
                    CandidatiPretMateriale.Any(h => cellText.Contains(h)) ||
                    CandidatiTotalLinie.Any(h => cellText.Contains(h)));

                if (matches >= 2)
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        var text = cell.GetString().Trim();
                        if (!headerMap.ContainsKey(text))
                        {
                            headerMap[text] = cell.Address.ColumnNumber;
                        }
                    }

                    return rowNumber;
                }
            }

            return -1;
        }

    /// <summary>
    /// Caută o coloană posibilă folosind termenii candidat și harta antetului.
    /// </summary>
    /// <param name="worksheet">Foaia analizată.</param>
    /// <param name="headerRow">Rândul de antet detectat.</param>
    /// <param name="headerMap">Maparea etichetelor către indexul coloanelor.</param>
    /// <param name="candidates">Termenii cheie care descriu coloana căutată.</param>
    /// <returns>Indexul coloanei sau <c>null</c> dacă nu a fost găsită.</returns>
    private int? FindColumn(IXLWorksheet worksheet, int headerRow, Dictionary<string, int> headerMap, params string[] candidates)
        {
            if (headerRow <= 0)
            {
                return null;
            }

            foreach (var cell in worksheet.Row(headerRow).CellsUsed())
            {
                var text = cell.GetString().Trim().ToLowerInvariant();
                foreach (var candidate in candidates)
                {
                    if (text.Contains(candidate.ToLowerInvariant()))
                    {
                        return cell.Address.ColumnNumber;
                    }
                }
            }

            foreach (var entry in headerMap)
            {
                var key = entry.Key.ToLowerInvariant();
                foreach (var candidate in candidates)
                {
                    if (key.Contains(candidate.ToLowerInvariant()))
                    {
                        return entry.Value;
                    }
                }
            }

            return null;
        }

    /// <summary>
    /// Determină dacă parserul trebuie să treacă pe fallback pozițional pe baza calității antetului detectat.
    /// </summary>
    /// <param name="foaie">Foaia evaluată.</param>
    /// <param name="randAntet">Indexul rândului de antet.</param>
    /// <param name="ultimulRand">Ultimul rând folosit în foaie.</param>
    /// <param name="colName">Indexul coloanei de denumire detectate.</param>
    /// <param name="colSymbol">Indexul coloanei de simbol detectate.</param>
    /// <returns><c>true</c> dacă este nevoie de fallback pozițional.</returns>
    private bool TrebuieFallbackPozitional(IXLWorksheet foaie, int randAntet, int ultimulRand, int? colName, int? colSymbol)
        {
            if (colName == null || colSymbol == null)
            {
                return true;
            }

            var sampleHasName = 0;
            var sampleTotal = 0;
            for (var row = randAntet + 1; row <= Math.Min(randAntet + 10, ultimulRand); row++)
            {
                sampleTotal++;
                var name = GetCellString(foaie, row, colName);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    sampleHasName++;
                }
            }

            return sampleTotal > 0 && sampleHasName * 1.0 / sampleTotal < 0.25;
        }

    /// <summary>
    /// Construiește maparea de fallback pentru coloane pe baza profilului și a configurării personalizate.
    /// </summary>
    /// <returns>Dicționarul rol→index folosit în mod pozițional.</returns>
    private Dictionary<DevizColumnRole, int> IaColoaneFallback()
        {
            if (_optiuni.CustomFallbackColumns != null)
            {
                return new Dictionary<DevizColumnRole, int>(_optiuni.CustomFallbackColumns);
            }

            return _optiuni.Profile switch
            {
                DevizParserProfile.Racsadia => new Dictionary<DevizColumnRole, int>
                {
                    { DevizColumnRole.Order, 1 },
                    { DevizColumnRole.Symbol, 2 },
                    { DevizColumnRole.Name, 4 },
                    { DevizColumnRole.UnitOfMeasure, 14 },
                    { DevizColumnRole.Quantity, 15 },
                    { DevizColumnRole.UnitPrice, 17 },
                    { DevizColumnRole.LineTotal, 19 },
                    { DevizColumnRole.MaterialsQuantity, 0 },
                    { DevizColumnRole.MaterialsUnitPrice, 0 },
                    { DevizColumnRole.LaborQuantity, 0 },
                    { DevizColumnRole.LaborUnitPrice, 0 },
                    { DevizColumnRole.EquipmentQuantity, 0 },
                    { DevizColumnRole.EquipmentUnitPrice, 0 },
                    { DevizColumnRole.TransportQuantity, 0 },
                    { DevizColumnRole.TransportUnitPrice, 0 }
                },
                DevizParserProfile.Deviz360 => new Dictionary<DevizColumnRole, int>
                {
                    { DevizColumnRole.Order, 1 },
                    { DevizColumnRole.Symbol, 2 },
                    { DevizColumnRole.Name, 4 },
                    { DevizColumnRole.UnitOfMeasure, 14 },
                    { DevizColumnRole.Quantity, 15 },
                    { DevizColumnRole.UnitPrice, 17 },
                    { DevizColumnRole.LineTotal, 19 },
                    { DevizColumnRole.MaterialsQuantity, 0 },
                    { DevizColumnRole.MaterialsUnitPrice, 0 },
                    { DevizColumnRole.LaborQuantity, 0 },
                    { DevizColumnRole.LaborUnitPrice, 0 },
                    { DevizColumnRole.EquipmentQuantity, 0 },
                    { DevizColumnRole.EquipmentUnitPrice, 0 },
                    { DevizColumnRole.TransportQuantity, 0 },
                    { DevizColumnRole.TransportUnitPrice, 0 }
                },
                _ => new Dictionary<DevizColumnRole, int>
                {
                    { DevizColumnRole.Order, 1 },
                    { DevizColumnRole.Symbol, 2 },
                    { DevizColumnRole.Name, 3 },
                    { DevizColumnRole.UnitOfMeasure, 7 },
                    { DevizColumnRole.Quantity, 8 },
                    { DevizColumnRole.UnitPrice, 9 },
                    { DevizColumnRole.LineTotal, 10 },
                    { DevizColumnRole.MaterialsQuantity, 0 },
                    { DevizColumnRole.MaterialsUnitPrice, 0 },
                    { DevizColumnRole.LaborQuantity, 0 },
                    { DevizColumnRole.LaborUnitPrice, 0 },
                    { DevizColumnRole.EquipmentQuantity, 0 },
                    { DevizColumnRole.EquipmentUnitPrice, 0 },
                    { DevizColumnRole.TransportQuantity, 0 },
                    { DevizColumnRole.TransportUnitPrice, 0 }
                }
            };
        }

    /// <summary>
    /// Obține indexul de coloană din maparea fallback pentru rolul specificat.
    /// </summary>
    /// <param name="fallback">Maparea rol→index folosită în mod pozițional.</param>
    /// <param name="role">Rolul căutat.</param>
    /// <returns>Indexul coloanei sau <c>null</c> dacă lipsește.</returns>
    private static int? IaFallback(Dictionary<DevizColumnRole, int> fallback, DevizColumnRole role)
        {
            if (!fallback.TryGetValue(role, out var index) || index <= 0)
            {
                return null;
            }

            return index;
        }

    /// <summary>
    /// Actualizează valorile pe categorii (materiale, manoperă etc.) pe baza unui rând de defalcare detectat.
    /// </summary>
    /// <param name="randBaza">Rândul principal căruia i se aplică defalcarea.</param>
    /// <param name="cheieCategorie">Categoria identificată (materials/labor/etc.).</param>
    /// <param name="celuleFolosite">Celulele ce conțin valorile de defalcare.</param>
    private void AplicaDefalcareCategorie(RowItem randBaza, string cheieCategorie, List<IXLCell> celuleFolosite)
        {
            var valoriGasite = ExtractNumericTokens(celuleFolosite);
            if (valoriGasite.Count == 0)
            {
                return;
            }

            var total = valoriGasite.Last();
            var pretUnitar = valoriGasite.Count >= 2 ? valoriGasite[^2] : 0m;
            var cantitate = valoriGasite.Count >= 3 ? valoriGasite[^3] : 0m;

            if (cantitate == 0m)
            {
                cantitate = randBaza.Quantity;
            }

            if (pretUnitar == 0m && cantitate != 0m && total != 0m)
            {
                pretUnitar = Math.Round(total / cantitate, 4);
            }

            switch (cheieCategorie.ToLowerInvariant())
            {
                case "materials":
                    randBaza.Categories.Materials.Total = total;
                    if (pretUnitar != 0m) randBaza.Categories.Materials.UnitPrice = pretUnitar;
                    if (cantitate != 0m) randBaza.Categories.Materials.Quantity = cantitate;
                    break;
                case "labor":
                    randBaza.Categories.Labor.Total = total;
                    if (pretUnitar != 0m) randBaza.Categories.Labor.UnitPrice = pretUnitar;
                    if (cantitate != 0m) randBaza.Categories.Labor.Quantity = cantitate;
                    break;
                case "equipment":
                    randBaza.Categories.Equipment.Total = total;
                    if (pretUnitar != 0m) randBaza.Categories.Equipment.UnitPrice = pretUnitar;
                    if (cantitate != 0m) randBaza.Categories.Equipment.Quantity = cantitate;
                    break;
                case "transport":
                    randBaza.Categories.Transport.Total = total;
                    if (pretUnitar != 0m) randBaza.Categories.Transport.UnitPrice = pretUnitar;
                    if (cantitate != 0m) randBaza.Categories.Transport.Quantity = cantitate;
                    break;
            }

            randBaza.LineTotal = randBaza.Categories.Materials.Total + randBaza.Categories.Labor.Total + randBaza.Categories.Equipment.Total + randBaza.Categories.Transport.Total;
        }

    /// <summary>
    /// Rulează verificările de consistență pentru fiecare rând și colectează posibilele abateri.
    /// </summary>
    /// <param name="rezultat">Rezultatul în care se vor adăuga informațiile de validare.</param>
    private void VerificaRanduri(ParseResult rezultat)
        {
            foreach (var rand in rezultat.Rows)
            {
                var totalCategorii = rand.Categories.Materials.Total + rand.Categories.Labor.Total + rand.Categories.Equipment.Total + rand.Categories.Transport.Total;

                rand.Validation.CategoriesTotal = totalCategorii;
                rand.Validation.ComputedLineTotal = rand.ComputedLineTotal;
                rand.Validation.SheetLineTotal = rand.SheetLineTotal;
                rand.Validation.DifferenceToLineTotal = totalCategorii - rand.LineTotal;
                rand.Validation.DifferenceToComputedLineTotal = totalCategorii - rand.ComputedLineTotal;
                rand.Validation.DifferenceToSheetLineTotal = rand.SheetLineTotal.HasValue ? totalCategorii - rand.SheetLineTotal.Value : null;
                rand.Validation.CategoriesMatchLineTotal = true;
                rand.Validation.ComputedMatchesSheet = true;

                if (!AreClose(totalCategorii, rand.LineTotal))
                {
                    rand.Validation.CategoriesMatchLineTotal = false;
                    var diferenta = totalCategorii - rand.LineTotal;
                    var mesaj = $"{DescribeRow(rand)}: totalurile pe categorii difera de totalul liniei cu {diferenta}";
                    rand.Validation.Issues.Add(mesaj);
                    rezultat.Validation.RowIssues.Add(new RowIssue
                    {
                        Order = rand.Order,
                        Symbol = rand.Symbol,
                        Name = rand.Name,
                        Message = "Totalurile pe categorii nu corespund cu totalul liniei"
                    });
                    AddValidationError(rezultat.Validation, mesaj);
                }

                if (rand.SheetLineTotal.HasValue && !AreClose(rand.SheetLineTotal.Value, rand.ComputedLineTotal))
                {
                    rand.Validation.ComputedMatchesSheet = false;
                    var diferenta = rand.SheetLineTotal.Value - rand.ComputedLineTotal;
                    var mesaj = $"{DescribeRow(rand)}: totalul din foaie difera de cantitate × pret unitar cu {diferenta}";
                    rand.Validation.Issues.Add(mesaj);
                    rezultat.Validation.RowIssues.Add(new RowIssue
                    {
                        Order = rand.Order,
                        Symbol = rand.Symbol,
                        Name = rand.Name,
                        Message = "Totalul din foaie difera de cantitate × pret unitar"
                    });
                    AddValidationError(rezultat.Validation, mesaj);
                }
            }
        }

    /// <summary>
    /// Compară totalurile calculate cu sumarul din foaie și populează structura de validare.
    /// </summary>
    /// <param name="result">Rezultatul complet ce va fi actualizat.</param>
    /// <param name="summary">Sumarul extras din zona de totaluri a foii.</param>
    private void ApplySummaryToValidation(ParseResult result, SheetSummary summary)
        {
            if (summary.GrandTotal?.Total.HasValue == true)
            {
                var sheetTotal = summary.GrandTotal.Total.Value;
                var check = new SummaryCheck
                {
                    Key = "grandTotal",
                    Label = summary.GrandTotal.Label,
                    RowIndex = summary.GrandTotal.RowIndex,
                    SheetTotal = sheetTotal,
                    ComputedTotal = result.ComputedTotals.GrandTotal,
                    TotalMatches = AreClose(result.ComputedTotals.GrandTotal, sheetTotal)
                };

                result.Validation.GrandTotalSummary = check;

                if (!result.Validation.GrandTotalFromSheet.HasValue)
                {
                    result.Validation.GrandTotalFromSheet = sheetTotal;
                    result.Validation.GrandTotalMatchesSheet = check.TotalMatches;
                }


                if (!check.TotalMatches)
                {
                    AddValidationError(result.Validation, $"Totalul general calculat {check.ComputedTotal} difera de totalul din foaie {sheetTotal}");
                }

                if (summary.GrandTotal.Quantity.HasValue)
                {
                    check.SheetQuantity = summary.GrandTotal.Quantity;
                    check.ComputedQuantity = result.ComputedTotals.OverallQuantity;
                    check.QuantityMatches = AreClose(result.ComputedTotals.OverallQuantity, summary.GrandTotal.Quantity.Value);
                    if (!check.QuantityMatches)
                    {
                        AddValidationError(result.Validation, $"Cantitatea totala calculata {check.ComputedQuantity} difera de cantitatea din foaie {summary.GrandTotal.Quantity.Value}");
                    }
                }
            }

            foreach (var entry in NumeAfisareCategorii)
            {
                summary.Categories.TryGetValue(entry.Key, out var summaryLine);
                var computedTotal = entry.Key switch
                {
                    "materials" => result.ComputedTotals.Materials,
                    "labor" => result.ComputedTotals.Labor,
                    "equipment" => result.ComputedTotals.Equipment,
                    "transport" => result.ComputedTotals.Transport,
                    _ => 0m
                };

                var computedQuantity = entry.Key switch
                {
                    "materials" => result.ComputedTotals.MaterialsQuantity,
                    "labor" => result.ComputedTotals.LaborQuantity,
                    "equipment" => result.ComputedTotals.EquipmentQuantity,
                    "transport" => result.ComputedTotals.TransportQuantity,
                    _ => 0m
                };

                var check = new SummaryCheck
                {
                    Key = entry.Key,
                    Label = summaryLine?.Label ?? entry.Value,
                    RowIndex = summaryLine?.RowIndex ?? 0,
                    SheetTotal = summaryLine?.Total,
                    SheetQuantity = summaryLine?.Quantity,
                    ComputedTotal = computedTotal,
                    ComputedQuantity = computedQuantity,
                    TotalMatches = summaryLine?.Total == null || AreClose(computedTotal, summaryLine.Total.Value),
                    QuantityMatches = summaryLine?.Quantity == null || AreClose(computedQuantity, summaryLine.Quantity.Value)
                };

                result.Validation.CategorySummaries[entry.Key] = check;

                if (summaryLine?.Total != null && !check.TotalMatches)
                {
                    AddValidationError(result.Validation, $"Total sumar pentru {entry.Value} calculat {computedTotal} difera de valoarea din foaie {summaryLine.Total.Value}");
                }

                if (summaryLine?.Quantity != null && !check.QuantityMatches)
                {
                    AddValidationError(result.Validation, $"Cantitatea sumara pentru {entry.Value} calculata {computedQuantity} difera de cantitatea din foaie {summaryLine.Quantity.Value}");
                }
            }

            if (summary.TotalQuantity?.Quantity.HasValue == true)
            {
                var sheetQuantity = summary.TotalQuantity.Quantity.Value;
                var check = new SummaryCheck
                {
                    Key = "overallQuantity",
                    Label = summary.TotalQuantity.Label,
                    RowIndex = summary.TotalQuantity.RowIndex,
                    SheetQuantity = sheetQuantity,
                    ComputedQuantity = result.ComputedTotals.OverallQuantity,
                    QuantityMatches = AreClose(result.ComputedTotals.OverallQuantity, sheetQuantity)
                };

                result.Validation.TotalQuantitySummary = check;

                if (!check.QuantityMatches)
                {
                    AddValidationError(result.Validation, $"Cantitatea totala calculata {check.ComputedQuantity} difera de cantitatea din foaie {sheetQuantity}");
                }
            }
        }

    /// <summary>
    /// Scanează zona de final a foii pentru a identifica rândurile de sumar (categorii și totaluri).
    /// </summary>
    /// <param name="worksheet">Foaia inspectată.</param>
    /// <param name="startRow">Rândul de la care începe căutarea.</param>
    /// <param name="lastRow">Ultimul rând utilizat.</param>
    /// <returns>Structura <see cref="SheetSummary"/> găsită.</returns>
    private SheetSummary ScanSummaryArea(IXLWorksheet worksheet, int startRow, int lastRow)
        {
            var summary = new SheetSummary();
            if (lastRow <= 0)
            {
                return summary;
            }

            startRow = Math.Max(1, Math.Min(startRow, lastRow));
            var categoryColumns = new Dictionary<int, string>();

            for (var rowIndex = startRow; rowIndex <= lastRow; rowIndex++)
            {
                var row = worksheet.Row(rowIndex);
                var usedCells = row.CellsUsed().ToList();
                if (usedCells.Count == 0)
                {
                    continue;
                }

                foreach (var cell in usedCells)
                {
                    var cellLabel = NormalizeLabel(cell.GetString().Trim());
                    foreach (var kvp in TokeniRezumatCategorii)
                    {
                        if (categoryColumns.ContainsKey(cell.Address.ColumnNumber))
                        {
                            continue;
                        }

                        if (kvp.Value.Any(token => cellLabel.Contains(token)))
                        {
                            categoryColumns[cell.Address.ColumnNumber] = kvp.Key;
                            break;
                        }
                    }
                }

                var labelText = usedCells.First().GetString().Trim();
                var normalizedLabel = NormalizeLabel(labelText);
                var rowText = string.Join(" ", usedCells.Select(c => NormalizeLabel(c.GetString().Trim())));
                if (string.IsNullOrWhiteSpace(rowText))
                {
                    continue;
                }

                var numbers = ExtractNumericTokens(usedCells);
                if (numbers.Count == 0)
                {
                    continue;
                }

                if (normalizedLabel.Contains("total deviz") && normalizedLabel.Contains("fara tva"))
                {
                    var key = summary.ExtraTotals.ContainsKey("totalDevizFaraTvaInitial") ? "totalDevizFaraTvaFinal" : "totalDevizFaraTvaInitial";
                    summary.ExtraTotals[key] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                if (normalizedLabel.Contains("total cheltuieli directe"))
                {
                    summary.ExtraTotals["totalCheltuieliDirecte"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };

                    var orderedCategories = new[] { "materials", "labor", "equipment", "transport" };
                    for (var idx = 0; idx < orderedCategories.Length && idx < numbers.Count - 1; idx++)
                    {
                        var categoryKey = orderedCategories[idx];
                        var labelForCategory = NumeAfisareCategorii.TryGetValue(categoryKey, out var displayName)
                            ? displayName
                            : categoryKey;

                        summary.Categories[categoryKey] = new SummaryLine
                        {
                            RowIndex = rowIndex,
                            Label = labelForCategory,
                            Total = numbers[idx]
                        };
                    }

                    CaptureCategoryTotalsFromRow(worksheet, rowIndex, categoryColumns, summary);
                    continue;
                }

                if (normalizedLabel.Contains("contributia asiguratorie"))
                {
                    summary.ExtraTotals["contributiaAsiguratorie"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                if (normalizedLabel.Contains("cheltuieli indirecte"))
                {
                    summary.ExtraTotals["cheltuieliIndirecte"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                if (normalizedLabel.Equals("profit"))
                {
                    summary.ExtraTotals["profit"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                if (normalizedLabel.Contains("total general") && normalizedLabel.Contains("fara tva"))
                {
                    summary.ExtraTotals["totalGeneralFaraTva"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                if (normalizedLabel.StartsWith("tva"))
                {
                    summary.ExtraTotals["tva"] = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = labelText,
                        Total = numbers.Last()
                    };
                    continue;
                }

                var matchedCategory = false;
                foreach (var kvp in TokeniRezumatCategorii)
                {
                    if (rowText.Contains("total") && kvp.Value.Any(token => rowText.Contains(token)))
                    {
                        summary.Categories[kvp.Key] = new SummaryLine
                        {
                            RowIndex = rowIndex,
                            Label = usedCells.First().GetString().Trim(),
                            Quantity = numbers.Count >= 2 ? numbers[^2] : (decimal?)null,
                            Total = numbers.Last()
                        };
                        matchedCategory = true;
                        break;
                    }
                }

                if (matchedCategory)
                {
                    continue;
                }

                if (rowText.Contains("total general") || rowText.Contains("total fișa") || rowText.Contains("total fisa") || rowText.Contains("total lucrare") || rowText.Trim().Equals("total"))
                {
                    summary.GrandTotal = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = usedCells.First().GetString().Trim(),
                        Total = numbers.Last(),
                        Quantity = numbers.Count >= 2 ? numbers[^2] : (decimal?)null
                    };
                    continue;
                }

                if ((rowText.Contains("total cant") || rowText.Contains("cantitate total") || rowText.Contains("cantități totale") || rowText.Contains("cantitati totale")) && numbers.Count > 0)
                {
                    summary.TotalQuantity = new SummaryLine
                    {
                        RowIndex = rowIndex,
                        Label = usedCells.First().GetString().Trim(),
                        Quantity = numbers.Last()
                    };
                }
            }

            return summary;
        }

    /// <summary>
    /// Extrage totalurile pe categorii dintr-un rând de sumar și le atașează structurii de sumar.
    /// </summary>
    /// <param name="worksheet">Foaia sursă.</param>
    /// <param name="rowIndex">Indexul rândului analizat.</param>
    /// <param name="categoryColumns">Maparea coloanelor către cheile de categorie.</param>
    /// <param name="summary">Sumarul care va fi completat.</param>
    private void CaptureCategoryTotalsFromRow(IXLWorksheet worksheet, int rowIndex, Dictionary<int, string> categoryColumns, SheetSummary summary)
        {
            foreach (var entry in categoryColumns)
            {
                var columnNumber = entry.Key;
                var categoryKey = entry.Value;
                var cellValue = GetCellDecimal(worksheet, rowIndex, columnNumber, allowEmptyAsNull: true);
                if (!cellValue.HasValue)
                {
                    continue;
                }

                var label = NumeAfisareCategorii.TryGetValue(categoryKey, out var displayName)
                    ? displayName
                    : categoryKey;

                summary.Categories[categoryKey] = new SummaryLine
                {
                    RowIndex = rowIndex,
                    Label = label,
                    Total = cellValue.Value
                };
            }
        }

    /// <summary>
    /// Încearcă să găsească un total general numeric în absența unei linii de sumar clare.
    /// </summary>
    /// <param name="worksheet">Foaia căutată.</param>
    /// <param name="lastRow">Indexul ultimului rând folosit.</param>
    /// <returns>Valoarea totalului general sau <c>null</c> dacă nu este găsită.</returns>
    private decimal? LocateFallbackGrandTotal(IXLWorksheet worksheet, int lastRow)
        {
            for (var row = Math.Max(1, lastRow - 50); row <= lastRow; row++)
            {
                foreach (var cell in worksheet.Row(row).CellsUsed())
                {
                    var text = cell.GetString().Trim().ToLowerInvariant();
                    if (text.Contains("total") || text.Contains("total general") || text.Contains("sumă") || text.Contains("sum"))
                    {
                        var right = cell.CellRight();
                        if (right != null && right.TryGetValue(out double value))
                        {
                            return Convert.ToDecimal(value);
                        }

                        var below = cell.CellBelow();
                        if (below != null && below.TryGetValue(out double valueBelow))
                        {
                            return Convert.ToDecimal(valueBelow);
                        }
                    }
                }
            }

            return null;
        }

    /// <summary>
    /// Extrage toate valorile numerice detectabile dintr-o colecție de celule.
    /// </summary>
    /// <param name="cells">Celulele analizate.</param>
    /// <returns>Lista valorilor numerice găsite.</returns>
    private List<decimal> ExtractNumericTokens(IEnumerable<IXLCell> cells)
        {
            var results = new List<decimal>();
            foreach (var cell in cells)
            {
                if (cell.TryGetValue(out double numericValue))
                {
                    results.Add(Convert.ToDecimal(numericValue));
                    continue;
                }

                var text = cell.GetString().Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                var cleaned = new string(text.Where(ch => char.IsDigit(ch) || ch == ',' || ch == '.' || ch == '-').ToArray());
                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    continue;
                }

                if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out var dec))
                {
                    results.Add(dec);
                    continue;
                }

                if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.GetCultureInfo("fr-FR"), out dec))
                {
                    results.Add(dec);
                }
            }

            return results;
        }

    /// <summary>
    /// Generează o descriere scurtă a rândului pentru mesaje de validare.
    /// </summary>
    /// <param name="row">Rândul despre care se generează descrierea.</param>
    /// <returns>Șirul descriptiv.</returns>
    private string DescribeRow(RowItem row)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(row.Order)) parts.Add($"poz {row.Order}");
            if (!string.IsNullOrWhiteSpace(row.Symbol)) parts.Add(row.Symbol);
            if (!string.IsNullOrWhiteSpace(row.Name)) parts.Add(row.Name);
            return string.Join(" - ", parts);
        }

    /// <summary>
    /// Verifică dacă două valori sunt suficient de apropiate conform toleranței din opțiuni.
    /// </summary>
    /// <param name="left">Prima valoare.</param>
    /// <param name="right">A doua valoare.</param>
    /// <returns><c>true</c> dacă diferența absolută este sub toleranță.</returns>
    private bool AreClose(decimal left, decimal right)
        {
            return Math.Abs(left - right) <= _optiuni.ValidationTolerance;
        }

    /// <summary>
    /// Adaugă un mesaj de validare în lista globală de erori.
    /// </summary>
    /// <param name="validation">Structura în care se stochează erorile.</param>
    /// <param name="message">Mesajul ce trebuie salvat.</param>
    private static void AddValidationError(ValidationSummary validation, string message)
        {
            if (!validation.Errors.Contains(message))
            {
                validation.Errors.Add(message);
            }
        }

    /// <summary>
    /// Calculează totalul liniei folosind cantitatea și prețul unitar, cu fallback la valoarea explicită dacă este validă.
    /// </summary>
    /// <param name="quantity">Cantitatea detectată.</param>
    /// <param name="unitPrice">Prețul unitar detectat.</param>
    /// <param name="explicitTotal">Totalul furnizat direct în foaie.</param>
    /// <returns>Totalul rezultat.</returns>
    private static decimal NormalizeTotal(decimal quantity, decimal unitPrice, decimal? explicitTotal)
        {
            if (explicitTotal.HasValue)
            {
                return explicitTotal.Value;
            }

            return Math.Round(quantity * unitPrice, 4);
        }

    /// <summary>
    /// Citește textul dintr-o celulă specificată, întorcând șirul gol dacă indexul lipsește.
    /// </summary>
    /// <param name="worksheet">Foaia din care se citește.</param>
    /// <param name="row">Indexul rândului.</param>
    /// <param name="column">Indexul coloanei sau <c>null</c> pentru lipsă.</param>
    /// <returns>Textul celulei sau șir gol.</returns>
    private static string GetCellString(IXLWorksheet worksheet, int row, int? column)
        {
            if (column == null)
            {
                return string.Empty;
            }

            var cell = worksheet.Cell(row, column.Value);
            return cell?.GetString().Trim() ?? string.Empty;
        }

    /// <summary>
    /// Citește o valoare numerică dintr-o celulă, cu opțiuni pentru default și tratarea celulelor goale.
    /// </summary>
    /// <param name="worksheet">Foaia de lucru.</param>
    /// <param name="row">Rândul vizat.</param>
    /// <param name="column">Coloana vizată.</param>
    /// <param name="defaultValue">Valoarea implicită dacă analiza numerică eșuează.</param>
    /// <param name="allowEmptyAsNull">Determină dacă celulele goale întorc <c>null</c>.</param>
    /// <returns>Valoarea numerică sau default/null conform opțiunilor.</returns>
    private static decimal? GetCellDecimal(IXLWorksheet worksheet, int row, int? column, decimal? defaultValue = 0m, bool allowEmptyAsNull = false)
        {
            if (column == null)
            {
                return allowEmptyAsNull ? null : defaultValue;
            }

            var cell = worksheet.Cell(row, column.Value);
            if (cell == null || cell.IsEmpty())
            {
                return allowEmptyAsNull ? null : defaultValue;
            }

            if (cell.TryGetValue(out double numericValue))
            {
                return Convert.ToDecimal(numericValue);
            }

            var text = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return allowEmptyAsNull ? null : defaultValue;
            }

            text = text.Replace(" ", string.Empty);
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out var dec))
            {
                return dec;
            }

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.GetCultureInfo("fr-FR"), out dec))
            {
                return dec;
            }

            var cleaned = new string(text.Where(ch => char.IsDigit(ch) || ch == ',' || ch == '.' || ch == '-').ToArray());
            if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out dec))
            {
                return dec;
            }

            if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.GetCultureInfo("fr-FR"), out dec))
            {
                return dec;
            }

            return allowEmptyAsNull ? null : defaultValue;
        }

    /// <summary>
    /// Verifică dacă rândul pare a fi nivel principal (nu subtotal sau linie de sumar).
    /// </summary>
    /// <param name="row">Rândul analizat.</param>
    /// <returns><c>true</c> dacă rândul pare principal.</returns>
    private static bool IsTopLevelRow(RowItem row)
        {
            if (row == null || string.IsNullOrWhiteSpace(row.Order))
            {
                return false;
            }

            return !row.Order.Contains('.');
        }

    /// <summary>
    /// Normalizează un text pentru comparații (lowercase fără spații suplimentare și diacritice de bază).
    /// </summary>
    /// <param name="text">Textul original.</param>
    /// <returns>Forma normalizată.</returns>
    private static string NormalizeLabel(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var lower = text.Trim().ToLowerInvariant();
            return lower
                .Replace("ă", "a")
                .Replace("â", "a")
                .Replace("î", "i")
                .Replace("ș", "s")
                .Replace("ş", "s")
                .Replace("ț", "t")
                .Replace("ţ", "t");
        }

    /// <summary>
    /// Completează totalurile suplimentare (cheltuieli directe, indirecte, profit, TVA etc.) pe baza sumarului extras.
    /// </summary>
    /// <param name="result">Rezultatul parsării ce va fi actualizat.</param>
    /// <param name="summary">Sumarul extras din foaie.</param>
    private void ApplyAdditionalSummaryTotals(ParseResult result, SheetSummary summary)
        {
            if (summary == null)
            {
                return;
            }

            if (summary.ExtraTotals.TryGetValue("contributiaAsiguratorie", out var contributie) && contributie.Total.HasValue)
            {
                result.ComputedTotals.OtherDirectCosts = contributie.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("totalCheltuieliDirecte", out var totalDirect) && totalDirect.Total.HasValue)
            {
                result.ComputedTotals.TotalCheltuieliDirecte = totalDirect.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("cheltuieliIndirecte", out var indirecte) && indirecte.Total.HasValue)
            {
                result.ComputedTotals.CheltuieliIndirecte = indirecte.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("profit", out var profit) && profit.Total.HasValue)
            {
                result.ComputedTotals.Profit = profit.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("totalDevizFaraTvaInitial", out var totalDevizInitial) && totalDevizInitial.Total.HasValue)
            {
                result.ComputedTotals.TotalDevizFaraTvaInitial = totalDevizInitial.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("totalDevizFaraTvaFinal", out var totalDevizFinal) && totalDevizFinal.Total.HasValue)
            {
                result.ComputedTotals.TotalDevizFaraTvaFinal = totalDevizFinal.Total.Value;
            }

            if (summary.ExtraTotals.TryGetValue("totalGeneralFaraTva", out var totalGeneralFaraTva) && totalGeneralFaraTva.Total.HasValue)
            {
                result.ComputedTotals.TotalGeneralFaraTva = totalGeneralFaraTva.Total.Value;
            }
            else if (result.ComputedTotals.TotalDevizFaraTvaFinal != 0m)
            {
                result.ComputedTotals.TotalGeneralFaraTva = result.ComputedTotals.TotalDevizFaraTvaFinal;
            }

            if (summary.ExtraTotals.TryGetValue("tva", out var tvaLine) && tvaLine.Total.HasValue)
            {
                result.ComputedTotals.Vat = tvaLine.Total.Value;
            }

            if (result.ComputedTotals.TotalCheltuieliDirecte == 0m && (result.ComputedTotals.DirectGrandTotal != 0m || result.ComputedTotals.OtherDirectCosts != 0m))
            {
                result.ComputedTotals.TotalCheltuieliDirecte = result.ComputedTotals.DirectGrandTotal + result.ComputedTotals.OtherDirectCosts;
            }

            if (result.ComputedTotals.OtherDirectCosts == 0m && result.ComputedTotals.TotalCheltuieliDirecte != 0m && result.ComputedTotals.DirectGrandTotal != 0m)
            {
                var residual = result.ComputedTotals.TotalCheltuieliDirecte - result.ComputedTotals.DirectGrandTotal;
                if (Math.Abs(residual) > _optiuni.ValidationTolerance)
                {
                    result.ComputedTotals.OtherDirectCosts = residual;
                }
            }

            if (result.ComputedTotals.TotalDevizFaraTvaFinal == 0m && (result.ComputedTotals.TotalCheltuieliDirecte != 0m || result.ComputedTotals.CheltuieliIndirecte != 0m || result.ComputedTotals.Profit != 0m))
            {
                result.ComputedTotals.TotalDevizFaraTvaFinal = result.ComputedTotals.TotalCheltuieliDirecte + result.ComputedTotals.CheltuieliIndirecte + result.ComputedTotals.Profit;
            }

            if (result.ComputedTotals.TotalGeneralFaraTva == 0m && result.ComputedTotals.TotalDevizFaraTvaFinal != 0m)
            {
                result.ComputedTotals.TotalGeneralFaraTva = result.ComputedTotals.TotalDevizFaraTvaFinal;
            }

            if (result.ComputedTotals.TotalGeneralFaraTva != 0m || result.ComputedTotals.Vat != 0m)
            {
                result.ComputedTotals.GrandTotal = result.ComputedTotals.TotalGeneralFaraTva + result.ComputedTotals.Vat;
            }
            else if (summary.GrandTotal?.Total.HasValue == true)
            {
                result.ComputedTotals.GrandTotal = summary.GrandTotal.Total.Value;
            }
        }
    }
}
