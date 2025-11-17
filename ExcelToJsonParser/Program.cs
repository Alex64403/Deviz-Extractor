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

namespace ExcelToJsonParser;

/// <summary>
/// Utilitar CLI general pentru conversia foilor Excel în JSON, cu salvare opțională în SQL Server.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Punctul de intrare: gestionează argumentele, rulează parserul și scrie rezultatele.
    /// </summary>
    private static async Task Main(string[] args)
    {
        var argumenteRamase = args.ToList();
        var modDump = argumenteRamase.Count > 0 && string.Equals(argumenteRamase[0], "--dump", StringComparison.OrdinalIgnoreCase);
        if (modDump)
        {
            argumenteRamase.RemoveAt(0);
        }

        string? sirConexiune = null;
        string? numeTabela = null;
        string? numeTabelaFlatten = null;
        var modStocare = "raw";
        var vreaBazaDate = false;

        for (var index = 0; index < argumenteRamase.Count; index++)
        {
            var argument = argumenteRamase[index];
            if (!argument.StartsWith("--", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(argument, "--save-db", StringComparison.OrdinalIgnoreCase))
            {
                vreaBazaDate = true;
                argumenteRamase.RemoveAt(index--);
                continue;
            }

            if (string.Equals(argument, "--connection", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= argumenteRamase.Count)
                {
                    Console.WriteLine("Missing value for --connection option.");
                    return;
                }

                sirConexiune = argumenteRamase[index + 1];
                vreaBazaDate = true;
                argumenteRamase.RemoveAt(index + 1);
                argumenteRamase.RemoveAt(index--);
                continue;
            }

            if (string.Equals(argument, "--table", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= argumenteRamase.Count)
                {
                    Console.WriteLine("Missing value for --table option.");
                    return;
                }

                numeTabela = argumenteRamase[index + 1];
                argumenteRamase.RemoveAt(index + 1);
                argumenteRamase.RemoveAt(index--);
                continue;
            }

            if (string.Equals(argument, "--flat-table", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= argumenteRamase.Count)
                {
                    Console.WriteLine("Missing value for --flat-table option.");
                    return;
                }

                numeTabelaFlatten = argumenteRamase[index + 1];
                argumenteRamase.RemoveAt(index + 1);
                argumenteRamase.RemoveAt(index--);
                continue;
            }

            if (string.Equals(argument, "--mode", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 >= argumenteRamase.Count)
                {
                    Console.WriteLine("Missing value for --mode option.");
                    return;
                }

                modStocare = argumenteRamase[index + 1];
                argumenteRamase.RemoveAt(index + 1);
                argumenteRamase.RemoveAt(index--);
                continue;
            }
        }

        if (argumenteRamase.Count == 0)
        {
            Console.WriteLine("Usage: ExcelToJsonParser <input.xlsx> [output.json]");
            Console.WriteLine("       ExcelToJsonParser --dump <input.xlsx> [rows] [cols]");
            Console.WriteLine("       ExcelToJsonParser [options] <input.xlsx> [output.json]");
            Console.WriteLine("Options:");
            Console.WriteLine("  --save-db                     Persist the resulting JSON to SQL Server");
            Console.WriteLine("  --connection <conn>           Explicit SQL Server connection string");
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
            using var dumpStream = File.Open(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var dumpWorkbook = new XLWorkbook(dumpStream);
            var dumpWorksheet = dumpWorkbook.Worksheets.First();
            DumpWorksheet(dumpWorksheet, randuriDeAratat, coloaneDeAratat);
            return;
        }

        var outputPath = argumenteRamase.Count > 1 ? argumenteRamase[1] : Path.ChangeExtension(inputPath, ".json");

        var optiuniParser = new DevizParserOptions
        {
            Profile = DevizParserProfile.Intersoft
        };

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

            var cleanPath = Path.Combine(Path.GetDirectoryName(outputPath) ?? string.Empty, Path.GetFileNameWithoutExtension(outputPath) + "_clean" + Path.GetExtension(outputPath));
            File.WriteAllText(cleanPath, JsonConvert.SerializeObject(rezultatCuratat, Formatting.Indented));
            Console.WriteLine($"Clean JSON written to {cleanPath}");
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
    /// Afișează un dump de celule pentru a inspecta vizual structura foii Excel.
    /// </summary>
    /// <param name="worksheet">Foaia inspectată.</param>
    /// <param name="rows">Numărul de rânduri afișate.</param>
    /// <param name="cols">Numărul de coloane afișate.</param>
    private static void DumpWorksheet(IXLWorksheet worksheet, int rows, int cols)
    {
        Console.WriteLine($"Dumping first {rows} rows and {cols} columns from worksheet '{worksheet.Name}'");
        for (var r = 1; r <= rows; r++)
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
