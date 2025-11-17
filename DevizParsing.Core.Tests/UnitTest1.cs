using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DevizParsing.Core.Excel;

namespace DevizParsing.Core.Tests;

/// <summary>
/// Suite de teste pentru <see cref="DevizWorksheetParser"/> care verifică profilele și sumarizările.
/// </summary>
public class DevizWorksheetParserTests
{
    /// <summary>
    /// Returnează calea completă către fișierul de test din directorul TestData.
    /// </summary>
    /// <param name="fileName">Numele fișierului din TestData.</param>
    /// <returns>Calea completă către fișier.</returns>
    private static string GetTestDataPath(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "TestData", fileName);
    }

    /// <summary>
    /// Verifică faptul că workbook-ul Intersoft produce totalurile așteptate și raportează discrepanțele detectate.
    /// </summary>
    [Fact]
    public void Parse_IntersoftWorkbook_ComputedTotalsMatchExpected()
    {
        var caleFisier = GetTestDataPath("RACASDIA INTERSOFT (2).xlsx");
        Assert.True(File.Exists(caleFisier), $"Missing test workbook at {caleFisier}");

        var parserSimplu = new DevizWorksheetParser(new DevizParserOptions
        {
            Profile = DevizParserProfile.Intersoft,
            ValidationTolerance = 0.01m
        });

        var rezultat = parserSimplu.Parse(caleFisier);

        Assert.NotEmpty(rezultat.Rows);
    Assert.Equal("MODERNIZARE DJ 573D RACASDIA - VRANIUT", rezultat.Metadata.Obiectiv);
    Assert.Equal("LUCRARI DE DRUM", rezultat.Metadata.Obiect);
    Assert.Equal("DEVIZ LUCRARI", rezultat.Metadata.StadiuFizic);
        Assert.Equal(19025605.49m, rezultat.ComputedTotals.Materials);
        Assert.Equal(2694538.19m, rezultat.ComputedTotals.Labor);
        Assert.Equal(1007652.51m, rezultat.ComputedTotals.Equipment);
        Assert.Equal(2525742.85m, rezultat.ComputedTotals.Transport);
        Assert.Equal(25253539.04m, rezultat.ComputedTotals.GrandTotal);
        Assert.Equal(16398288.62m, rezultat.Validation.GrandTotalFromSheet);
        Assert.False(rezultat.Validation.GrandTotalMatchesSheet);
        Assert.Contains(rezultat.Validation.Errors, mesaj => mesaj.Contains("Computed grand total 25253539.04", StringComparison.Ordinal));

        var primRand = rezultat.Rows.First(r => r.Order == "1");
        Assert.Equal("RpDC13H%", primRand.Symbol);
        Assert.Equal(58902.05m, primRand.LineTotal);
        Assert.Equal(17322.05m, primRand.Categories.Materials.Total);
        Assert.False(primRand.Validation.ComputedMatchesSheet);
    }

    /// <summary>
    /// Confirmă că parserul calculează corect defalcările pe categorii dintr-o foaie construită în memorie.
    /// </summary>
    [Fact]
    public void ParseWorksheet_InMemorySheetWithSummaries_ComputesCategoryBreakdowns()
    {
        using var workbook = new XLWorkbook();
        var foaie = workbook.AddWorksheet("Sample");

        foaie.Cell(1, 1).Value = "Nr crt";
        foaie.Cell(1, 2).Value = "Cod";
        foaie.Cell(1, 3).Value = "Denumire";
        foaie.Cell(1, 4).Value = "U.M.";
        foaie.Cell(1, 5).Value = "Cantitate";
        foaie.Cell(1, 6).Value = "Pret unitar";
        foaie.Cell(1, 7).Value = "Valoare";

        foaie.Cell(2, 1).Value = "1";
        foaie.Cell(2, 2).Value = "TEST1";
        foaie.Cell(2, 3).Value = "Pozitie test";
        foaie.Cell(2, 4).Value = "mp";
        foaie.Cell(2, 5).Value = 10m;
        foaie.Cell(2, 6).Value = 25m;
        foaie.Cell(2, 7).Value = 250m;

        foaie.Cell(3, 3).Value = "Total materiale";
        foaie.Cell(3, 5).Value = 10m;
        foaie.Cell(3, 6).Value = 25m;
        foaie.Cell(3, 7).Value = 250m;

        foaie.Cell(4, 1).Value = "Total general";
        foaie.Cell(4, 5).Value = 10m;
        foaie.Cell(4, 7).Value = 250m;

        var parserSimplu = new DevizWorksheetParser(new DevizParserOptions
        {
            Profile = DevizParserProfile.Intersoft,
            ValidationTolerance = 0.001m
        });

        var rezultat = parserSimplu.ParseWorksheet(foaie, "InMemory.xlsx");

        Assert.NotEmpty(rezultat.Rows);
        var rand = Assert.Single(rezultat.Rows);
        Assert.Equal("Pozitie test", rand.Name);
        Assert.Equal(10m, rand.Quantity);
        Assert.Equal(25m, rand.UnitPrice);
        Assert.Equal(250m, rand.LineTotal);
        Assert.Equal(250m, rand.Categories.Materials.Total);
        Assert.True(rezultat.Validation.GrandTotalMatchesSheet);
        Assert.Equal(250m, rezultat.Validation.GrandTotalFromSheet);
        Assert.Equal(250m, rezultat.ComputedTotals.GrandTotal);
        Assert.Equal(250m, rezultat.ComputedTotals.Materials);
        Assert.Empty(rezultat.Validation.Errors);
    }

    /// <summary>
    /// Asigură că profilul Racsadia folosește coloanele poziționale configurate și generează avertismentele corespunzătoare.
    /// </summary>
    [Fact]
    public void ParseWorksheet_RacsadiaFallback_UsesConfiguredColumns()
    {
        using var workbook = new XLWorkbook();
        var foaie = workbook.AddWorksheet("Racsadia");

        foaie.Cell(1, 1).Value = "Placeholder";
        foaie.Cell(2, 1).Value = "1";
        foaie.Cell(2, 2).Value = "RC01";
        foaie.Cell(2, 4).Value = "Element Racsadia";
        foaie.Cell(2, 14).Value = "m";
        foaie.Cell(2, 15).Value = 5m;
        foaie.Cell(2, 17).Value = 100m;
        foaie.Cell(2, 19).Value = 500m;

        foaie.Cell(3, 1).Value = "Total general";
        foaie.Cell(3, 15).Value = 5m;
        foaie.Cell(3, 19).Value = 500m;

        var parserSimplu = new DevizWorksheetParser(new DevizParserOptions
        {
            Profile = DevizParserProfile.Racsadia,
            ForcePositionalFallback = true,
            ValidationTolerance = 0.001m
        });

        var rezultat = parserSimplu.ParseWorksheet(foaie, "Racsadia.xlsx");

        var rand = Assert.Single(rezultat.Rows);
        Assert.Equal("Element Racsadia", rand.Name);
        Assert.Equal("m", rand.UnitOfMeasure);
        Assert.Equal(5m, rand.Quantity);
        Assert.Equal(100m, rand.UnitPrice);
        Assert.Equal(500m, rand.LineTotal);
        Assert.Equal(500m, rand.Categories.Materials.Total);
        Assert.Equal(5m, rand.Categories.Materials.Quantity);
        Assert.Equal(100m, rand.Categories.Materials.UnitPrice);
        Assert.Equal(500m, rezultat.ComputedTotals.GrandTotal);
        Assert.True(rezultat.Validation.GrandTotalMatchesSheet);
        var avertisment = Assert.Single(rezultat.Validation.Errors);
        Assert.Contains("Header row not detected", avertisment, StringComparison.Ordinal);
    }

    /// <summary>
    /// Testează că profilul Deviz360 poate funcționa cu mapări fallback personalizate.
    /// </summary>
    [Fact]
    public void ParseWorksheet_Deviz360Profile_UsesCustomFallbackColumns()
    {
        using var workbook = new XLWorkbook();
        var foaie = workbook.AddWorksheet("Deviz360");

        foaie.Cell(1, 1).Value = "Some intro row";
        foaie.Cell(2, 1).Value = "1";
        foaie.Cell(2, 2).Value = "DV01";
        foaie.Cell(2, 5).Value = "Element Deviz 360";
        foaie.Cell(2, 6).Value = "m";
        foaie.Cell(2, 7).Value = 12m;
        foaie.Cell(2, 8).Value = 50m;
        foaie.Cell(2, 9).Value = 600m;

        foaie.Cell(3, 1).Value = "TOTAL GENERAL";
        foaie.Cell(3, 7).Value = 12m;
        foaie.Cell(3, 9).Value = 600m;

        var coloanePersonalizate = new Dictionary<DevizColumnRole, int>
        {
            { DevizColumnRole.Order, 1 },
            { DevizColumnRole.Symbol, 2 },
            { DevizColumnRole.Name, 5 },
            { DevizColumnRole.UnitOfMeasure, 6 },
            { DevizColumnRole.Quantity, 7 },
            { DevizColumnRole.UnitPrice, 8 },
            { DevizColumnRole.LineTotal, 9 }
        };

        var parserSimplu = new DevizWorksheetParser(new DevizParserOptions
        {
            Profile = DevizParserProfile.Deviz360,
            ForcePositionalFallback = true,
            CustomFallbackColumns = coloanePersonalizate,
            ValidationTolerance = 0.001m
        });

        var rezultat = parserSimplu.ParseWorksheet(foaie, "Deviz360.xlsx");

        var rand = Assert.Single(rezultat.Rows);
        Assert.Equal("Element Deviz 360", rand.Name);
        Assert.Equal("m", rand.UnitOfMeasure);
        Assert.Equal(12m, rand.Quantity);
        Assert.Equal(50m, rand.UnitPrice);
        Assert.Equal(600m, rand.LineTotal);
        Assert.Empty(rezultat.Validation.Errors);
        Assert.True(rezultat.Validation.GrandTotalMatchesSheet);
        Assert.Equal(600m, rezultat.ComputedTotals.GrandTotal);
    }
}
