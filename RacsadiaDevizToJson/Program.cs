using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DevizParsing.Core.Excel;
using DevizParsing.Core.Models;
using Newtonsoft.Json;

namespace RacsadiaDevizToJson;

/// <summary>
/// Utilitar CLI pentru conversia foilor Racsadia în JSON folosind parserul comun.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Metoda principală: parsează argumentele, rulează parserul și scrie fișierele JSON (complet și curățat).
    /// </summary>
    private static void Main(string[] args)
    {
        var argumenteRamase = args.ToList();
        var modDump = argumenteRamase.Count > 0 && string.Equals(argumenteRamase[0], "--dump", StringComparison.OrdinalIgnoreCase);
        if (modDump)
        {
            argumenteRamase.RemoveAt(0);
        }

        if (argumenteRamase.Count == 0)
        {
            Console.WriteLine("Usage: RacsadiaDevizToJson <input.xlsx> [output.json]");
            Console.WriteLine("       RacsadiaDevizToJson --dump <input.xlsx> [rows] [cols]");
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
            Profile = DevizParserProfile.Racsadia
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
    }

    /// <summary>
    /// Afișează un dump al celulelor din foaie pentru depanare (primele N rânduri/coloane).
    /// </summary>
    /// <param name="worksheet">Foaia de lucru inspectată.</param>
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
