using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DevizParsing.Core.Excel;
using DevizParsing.Core.Models;
using DevizParsing.Core.Persistence;
using Newtonsoft.Json;

namespace Deviz360DevizToJson;

/// <summary>
/// Punctul de intrare pentru utilitarul CLI Deviz360 → JSON, responsabil de parsare și salvare opțională în SQL.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Metoda principală care procesează argumentele, rulează parserul și gestionează output-ul JSON / baza de date.
    /// </summary>
    private static async Task Main(string[] args)
    {
        var modDump = false;
        var forteazaPozitional = false;
        string? fisierColoane = null;
        decimal? tolerantaPersonalizata = null;
        string? sirConexiune = null;
        string? numeTabela = null;
    string? numeTabelaFlatten = null;
    var modStocare = "raw";
        var vreaBazaDate = false;
        var argumenteRamase = new List<string>();

        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index];
            if (string.Equals(argument, "--dump", StringComparison.OrdinalIgnoreCase))
            {
                modDump = true;
                continue;
            }

            if (string.Equals(argument, "--force-positional", StringComparison.OrdinalIgnoreCase))
            {
                forteazaPozitional = true;
                continue;
            }

            if (string.Equals(argument, "--save-db", StringComparison.OrdinalIgnoreCase))
            {
                vreaBazaDate = true;
                continue;
            }

            if (string.Equals(argument, "--connection", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --connection option.");
                    return;
                }

                sirConexiune = args[++index];
                vreaBazaDate = true;
                continue;
            }

            if (string.Equals(argument, "--table", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --table option.");
                    return;
                }

                numeTabela = args[++index];
                continue;
            }

            if (string.Equals(argument, "--flat-table", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --flat-table option.");
                    return;
                }

                numeTabelaFlatten = args[++index];
                continue;
            }

            if (string.Equals(argument, "--mode", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --mode option.");
                    return;
                }

                modStocare = args[++index];
                continue;
            }

            if (string.Equals(argument, "--columns", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --columns option.");
                    return;
                }

                fisierColoane = args[++index];
                continue;
            }

            if (string.Equals(argument, "--tolerance", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= args.Length)
                {
                    Console.WriteLine("Missing value for --tolerance option.");
                    return;
                }

                var toleranceValue = args[++index];
                if (!decimal.TryParse(toleranceValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedTolerance))
                {
                    Console.WriteLine($"Could not parse tolerance '{toleranceValue}'.");
                    return;
                }

                tolerantaPersonalizata = parsedTolerance;
                continue;
            }

            argumenteRamase.Add(argument);
        }

        if (argumenteRamase.Count == 0)
        {
            Console.WriteLine("Usage: Deviz360DevizToJson [options] <input.xlsx> [output.json]");
            Console.WriteLine("Options:");
            Console.WriteLine("  --dump [rows] [cols] [start] Dump the worksheet cells for inspection (default 30x12 from row 1)");
            Console.WriteLine("  --force-positional            Force positional fallback columns");
            Console.WriteLine("  --columns <columns.json>      Provide explicit column indexes for positional parsing");
            Console.WriteLine("  --tolerance <decimal>         Override validation tolerance (default 0.05)");
            Console.WriteLine("  --save-db                     Persist the resulting JSON to SQL Server (uses ENV DEVIZ_DB_CONNECTION if no connection provided)");
            Console.WriteLine("  --connection <conn>           Explicit SQL Server connection string for persistence");
            Console.WriteLine("  --table <schema.table>        Target document table (default dbo.DevizImportRaw; _Pozitii/_Categorii are implied)");
            Console.WriteLine("  --flat-table <schema.table>   Target staging table for flattened rows (default dbo.DevizImportStage)");
            Console.WriteLine("  --mode <raw|flat|both>        Choose persistence mode (default raw)");
            return;
        }

        modStocare = string.IsNullOrWhiteSpace(modStocare) ? "raw" : modStocare.Trim().ToLowerInvariant();
        if (modStocare != "raw" && modStocare != "flat" && modStocare != "both")
        {
            Console.WriteLine("Invalid --mode value. Allowed values: raw, flat, both.");
            return;
        }

        var inputPath = argumenteRamase[0];
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' does not exist.");
            return;
        }

        if (modDump)
        {
            var randuriDeAratat = argumenteRamase.Count > 1 && int.TryParse(argumenteRamase[1], out var parsedRows) ? parsedRows : 30;
            var coloaneDeAratat = argumenteRamase.Count > 2 && int.TryParse(argumenteRamase[2], out var parsedCols) ? parsedCols : 12;
            var randDeStart = argumenteRamase.Count > 3 && int.TryParse(argumenteRamase[3], out var parsedStart) ? Math.Max(1, parsedStart) : 1;

            using var dumpStream = File.Open(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var dumpWorkbook = new XLWorkbook(dumpStream);
            var dumpWorksheet = dumpWorkbook.Worksheets.First();
            DumpWorksheet(dumpWorksheet, randuriDeAratat, coloaneDeAratat, randDeStart);
            return;
        }

        var outputPath = argumenteRamase.Count > 1 ? argumenteRamase[1] : Path.ChangeExtension(inputPath, ".json");

        var optiuniParser = new DevizParserOptions
        {
            Profile = DevizParserProfile.Deviz360,
            ForcePositionalFallback = forteazaPozitional
        };

        if (tolerantaPersonalizata.HasValue)
        {
            optiuniParser.ValidationTolerance = tolerantaPersonalizata.Value;
        }

        if (!string.IsNullOrWhiteSpace(fisierColoane))
        {
            try
            {
                var textConfig = File.ReadAllText(fisierColoane);
                var configuratie = JsonConvert.DeserializeObject<Dictionary<string, int>>(textConfig);
                if (configuratie == null)
                {
                    Console.WriteLine($"Columns config '{fisierColoane}' is empty or invalid.");
                    return;
                }

                var coloaneMapate = new Dictionary<DevizColumnRole, int>();
                foreach (var intrare in configuratie)
                {
                    if (Enum.TryParse<DevizColumnRole>(intrare.Key, ignoreCase: true, out var rolGasit))
                    {
                        coloaneMapate[rolGasit] = intrare.Value;
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Unknown column role '{intrare.Key}' in config; ignoring.");
                    }
                }

                if (coloaneMapate.Count == 0)
                {
                    Console.WriteLine("Columns config does not contain any known roles.");
                    return;
                }

                optiuniParser.CustomFallbackColumns = coloaneMapate;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not load columns config: " + ex.Message);
                return;
            }
        }

        var parserSimplu = new DevizWorksheetParser(optiuniParser);
        var rezultat = parserSimplu.Parse(inputPath);

        if (rezultat.Validation.Errors.Count > 0)
        {
            Console.WriteLine("Validation warnings:");
            foreach (var mesaj in rezultat.Validation.Errors.Distinct())
            {
                Console.WriteLine(" - " + mesaj);
            }
        }

        var jsonFinal = JsonConvert.SerializeObject(rezultat, Formatting.Indented);
        File.WriteAllText(outputPath, jsonFinal);
        Console.WriteLine($"JSON written to {outputPath}");

        try
        {
            var randuriCuratate = rezultat.Rows
                .Where(r => !string.IsNullOrWhiteSpace(r.Order) || !string.IsNullOrWhiteSpace(r.Name) || r.LineTotal != 0m)
                .ToList();

            var rezultatCuratat = new ParseResult
            {
                SourceFile = rezultat.SourceFile,
                Sheet = rezultat.Sheet,
                Metadata = rezultat.Metadata,
                ComputedTotals = rezultat.ComputedTotals,
                Validation = rezultat.Validation,
                Rows = randuriCuratate
            };

            var caleCurata = Path.Combine(Path.GetDirectoryName(outputPath) ?? string.Empty, Path.GetFileNameWithoutExtension(outputPath) + "_clean" + Path.GetExtension(outputPath));
            File.WriteAllText(caleCurata, JsonConvert.SerializeObject(rezultatCuratat, Formatting.Indented));
            Console.WriteLine($"Clean JSON written to {caleCurata}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Could not write clean JSON: " + ex.Message);
        }

        if (vreaBazaDate)
        {
            sirConexiune ??= Environment.GetEnvironmentVariable("DEVIZ_DB_CONNECTION");
            if (string.IsNullOrWhiteSpace(sirConexiune))
            {
                Console.WriteLine("Database save skipped: no connection string provided (use --connection or set DEVIZ_DB_CONNECTION).");
                return;
            }

            try
            {
                var profil = optiuniParser.Profile.ToString();
                if (modStocare == "raw" || modStocare == "both")
                {
                    var writer = new ParseResultDatabaseWriter(sirConexiune, numeTabela);
                    var idInserat = await writer.SalveazaAsync(rezultat, profil);
                    if (idInserat.HasValue)
                    {
                        Console.WriteLine($"Raw record created with Id {idInserat.Value}.");
                    }
                    else
                    {
                        Console.WriteLine("Raw record created.");
                    }
                }

                if (modStocare == "flat" || modStocare == "both")
                {
                    var flatWriter = new ParseResultFlatDatabaseWriter(sirConexiune, numeTabelaFlatten);
                    var inserted = await flatWriter.SalveazaRanduriAsync(rezultat, profil);
                    Console.WriteLine($"Flattened rows inserted: {inserted}.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to persist data to database: " + ex.Message);
            }
        }
    }

    /// <summary>
    /// Listează în consolă un eșantion din foaia Excel pentru diagnostic (număr de rânduri/coloane specificat).
    /// </summary>
    /// <param name="worksheet">Foaia de lucru ce trebuie inspectată.</param>
    /// <param name="rows">Numărul de rânduri de afișat.</param>
    /// <param name="cols">Numărul de coloane de afișat.</param>
    /// <param name="startRow">Rândul de început (implicit 1).</param>
    private static void DumpWorksheet(IXLWorksheet worksheet, int rows, int cols, int startRow = 1)
    {
        rows = Math.Max(1, rows);
        cols = Math.Max(1, cols);

        var beginRow = Math.Max(1, startRow);
        var endRow = Math.Max(beginRow, beginRow + rows - 1);
        Console.WriteLine($"Dumping rows {beginRow} to {endRow} (count {rows}) and first {cols} columns from worksheet '{worksheet.Name}'");
        for (var r = beginRow; r <= endRow; r++)
        {
            var pieces = new List<string>();
            for (var c = 1; c <= cols; c++)
            {
                var cell = worksheet.Cell(r, c);
                string display;
                if (cell.IsEmpty())
                {
                    display = string.Empty;
                }
                else if (cell.DataType == XLDataType.Number)
                {
                    display = cell.GetDouble().ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    display = cell.GetString();
                }

                pieces.Add($"{c}:{display}");
            }

            Console.WriteLine($"{r,3}: {string.Join(" | ", pieces)}");
        }
    }
}
