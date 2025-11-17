using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DevizParsing.Core.Excel;
using DevizParsing.Core.Models;
using DevizParsing.Core.Persistence;
using Newtonsoft.Json;

namespace DevizApp.UI;

/// <summary>
/// Punctul de intrare pentru aplicația console interactivă Deviz.
///
/// Această clasă conține logica UI-ului text pentru:
/// - selectarea fișierului Excel,
/// - alegerea profilului de parsare,
/// - rularea parserului și scrierea rezultatelor ca JSON,
/// - afișarea avertizărilor de validare și persistarea opțională în SQL Server.
///
/// Clasele și metodele din acest fișier sunt intenționat simple și pot fi
/// reutilizate atunci când componenta este referențiată dintr-un proiect mai mare.
/// </summary>
internal static class Program
{
	/// <summary>
	/// Dicționar cu etichete prietenoase (românești) pentru fiecare profil de parsare.
	/// Cheia este <see cref="DevizParserProfile"/>, valoarea este textul afișat în UI.
	/// </summary>
	private static readonly Dictionary<DevizParserProfile, string> ProfileEtichete = new()
	{
		{ DevizParserProfile.Intersoft, "Intersoft" },
		{ DevizParserProfile.Racsadia, "Racsadia" },
		{ DevizParserProfile.Deviz360, "Deviz360" }
	};

	/// <summary>
	/// Punctul de intrare asincron al aplicației.
	///
	/// Fluxul principal:
	/// 1. Cere calea fișierului Excel și profilul de parsare.
	/// 2. Rulează parserul și scrie rezultatul ca JSON.
	/// 3. Afișează avertizările și permite salvarea în baza de date.
	/// </summary>
	/// <returns>Task asincron care se încheie când operațiile UI s-au terminat.</returns>
	private static async Task Main()
	{
		Console.WriteLine("========================================");
		Console.WriteLine("      Interfata interactiva Deviz       ");
		Console.WriteLine("========================================\n");

		var excelPath = PromptForExistingFile("Cale fisier Excel:");
		if (excelPath == null)
		{
			Console.WriteLine("Nu a fost selectat niciun fisier. Aplicatia se opreste.");
			return;
		}

		var profilAles = PromptForProfile();
		var caleIesire = PromptForOutputPath(excelPath);

		var optiuniParser = new DevizParserOptions
		{
			Profile = profilAles
		};

		var parserFoaie = new DevizWorksheetParser(optiuniParser);
		ParseResult rezultatParse;

		try
		{
			Console.WriteLine("\nSe proceseaza fisierul Excel...\n");
			rezultatParse = parserFoaie.Parse(excelPath);
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Nu s-a putut parsa fisierul: {ex.Message}");
			return;
		}

		var jsonComplet = JsonConvert.SerializeObject(rezultatParse, Formatting.Indented);
		File.WriteAllText(caleIesire, jsonComplet);
		Console.WriteLine($"JSON salvat in {caleIesire}");

		WriteCleanJson(rezultatParse, caleIesire);

		if (rezultatParse.Validation.Errors.Count > 0)
		{
			Console.WriteLine("\nAvertizari de validare:");
			foreach (var warning in rezultatParse.Validation.Errors.Distinct())
			{
				Console.WriteLine(" - " + warning);
			}
		}

		Console.WriteLine($"\nRanduri parsate: {rezultatParse.Rows.Count} (foaia '{rezultatParse.Sheet}').");
		Console.WriteLine($"Total general calculat: {rezultatParse.ComputedTotals.GrandTotal:N2}");

		if (!PromptYesNo("Salvati rezultatul in SQL Server?"))
		{
			Console.WriteLine("\nGata. Apasati Enter pentru a iesi.");
			Console.ReadLine();
			return;
		}

		var sirConexiune = Prompt("Sir de conexiune (lasati gol pentru a folosi DEVIZ_DB_CONNECTION):", allowEmpty: true);
		if (string.IsNullOrWhiteSpace(sirConexiune))
		{
			sirConexiune = Environment.GetEnvironmentVariable("DEVIZ_DB_CONNECTION");
		}

		if (string.IsNullOrWhiteSpace(sirConexiune))
		{
			Console.WriteLine("Nu exista un sir de conexiune. Salvarea este omisa.");
			Console.WriteLine("Apasati Enter pentru a iesi.");
			Console.ReadLine();
			return;
		}

		var numeTabela = Prompt("Tabel tinta (implicit dbo.DevizImportRaw):", allowEmpty: true);

		try
		{
			var writer = new ParseResultDatabaseWriter(sirConexiune!, string.IsNullOrWhiteSpace(numeTabela) ? null : numeTabela);
			var idSalvat = await writer.SalveazaAsync(rezultatParse, profilAles.ToString());
			Console.WriteLine(idSalvat.HasValue
				? $"Inregistrare salvata cu Id {idSalvat.Value}."
				: "Inregistrare salvata.");
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Salvarea in baza de date a esuat: {ex.Message}");
		}

		Console.WriteLine("\nTotul este gata. Apasati Enter pentru a iesi.");
		Console.ReadLine();
	}

	/// <summary>
	/// Cere utilizatorului o cale către un fișier existent.
	/// Dacă utilizatorul introduce o linie goală, returnează <c>null</c>.
	/// </summary>
	/// <param name="label">Textul etichetei afișate la prompt.</param>
	/// <returns>Calea extinsă a fișierului sau <c>null</c> dacă utilizatorul renunță.</returns>
	private static string? PromptForExistingFile(string label)
	{
		while (true)
		{
			var raspuns = Prompt($"{label}:");
			if (string.IsNullOrWhiteSpace(raspuns))
			{
				return null;
			}

			var caleGandita = ExpandPath(raspuns);
			if (File.Exists(caleGandita))
			{
				return caleGandita;
			}

			Console.WriteLine("Fisierul nu a fost gasit. Incercati din nou sau apasati Enter pentru a renunta.");
		}
	}

	/// <summary>
	/// Afișează lista de profile de parsare disponibile și citește o alegere validă.
	/// Acceptă index numeric, nume tehnic al enum-ului sau eticheta afișată.
	/// </summary>
	/// <returns>Profilul selectat de utilizator.</returns>
	private static DevizParserProfile PromptForProfile()
	{
		var profiluri = Enum.GetValues<DevizParserProfile>();
		Console.WriteLine("\nProfiluri de parsare disponibile:");
		for (var index = 0; index < profiluri.Length; index++)
		{
			Console.WriteLine($" {index + 1}. {GetProfileEticheta(profiluri[index])}");
		}

		while (true)
		{
			var raspuns = Prompt("Alegeti profilul (numar sau denumire):");
			if (int.TryParse(raspuns, out var numarProfil) && numarProfil >= 1 && numarProfil <= profiluri.Length)
			{
				return profiluri[numarProfil - 1];
			}

			if (Enum.TryParse<DevizParserProfile>(raspuns, true, out var profilDinEnum))
			{
				return profilDinEnum;
			}

			var profilPotrivit = profiluri.FirstOrDefault(p => string.Equals(GetProfileEticheta(p), raspuns, StringComparison.OrdinalIgnoreCase));
			if (string.Equals(GetProfileEticheta(profilPotrivit), raspuns, StringComparison.OrdinalIgnoreCase))
			{
				return profilPotrivit;
			}

			Console.WriteLine("Selectie invalida. Incercati din nou.");
		}
	}

	/// <summary>
	/// Solicită utilizatorului o cale de ieșire pentru fișierul JSON și
	/// returnează calea finală. Dacă utilizatorul apasă Enter se folosește
	/// sugestia implicită (aceeași cale ca Excel, dar cu extensia .json).
	/// </summary>
	/// <param name="excelPath">Calea fișierului Excel folosită pentru a crea sugestia.</param>
	/// <returns>Calea finală pentru fișierul JSON.</returns>
	private static string PromptForOutputPath(string excelPath)
	{
		var sugestie = Path.ChangeExtension(excelPath, ".json");
		Console.WriteLine($"\nApasati Enter pentru a folosi calea implicita sau introduceti o locatie alternativa.");
		var raspuns = Prompt($"Cale iesire (implicit {sugestie}):", allowEmpty: true);
		if (string.IsNullOrWhiteSpace(raspuns))
		{
			return sugestie;
		}

		var caleExtinsa = ExpandPath(raspuns);
		var directorTinta = Path.GetDirectoryName(caleExtinsa);
		if (!string.IsNullOrWhiteSpace(directorTinta) && !Directory.Exists(directorTinta))
		{
			Directory.CreateDirectory(directorTinta);
		}

		return caleExtinsa;
	}

	/// <summary>
	/// Construiește și scrie o versiune 'curățată' a rezultatului de parsare,
	/// eliminând rândurile fără cod/denumire și fără valoare (LineTotal == 0).
	/// Acest fișier este util pentru inspectare rapidă și debugging.
	/// </summary>
	/// <param name="result">Rezultatul complet al parserului.</param>
	/// <param name="outputPath">Calea fișierului JSON original pentru a genera numele *_clean.json.</param>
	private static void WriteCleanJson(ParseResult result, string outputPath)
	{
		try
		{
			var randuriCuratate = result.Rows
				.Where(r => !string.IsNullOrWhiteSpace(r.Order) || !string.IsNullOrWhiteSpace(r.Name) || r.LineTotal != 0m)
				.ToList();

			var rezultatCuratat = new ParseResult
			{
				SourceFile = result.SourceFile,
				Sheet = result.Sheet,
				Metadata = result.Metadata,
				ComputedTotals = result.ComputedTotals,
				Validation = result.Validation,
				Rows = randuriCuratate
			};

			var caleCurata = Path.Combine(Path.GetDirectoryName(outputPath) ?? string.Empty,
				Path.GetFileNameWithoutExtension(outputPath) + "_clean" + Path.GetExtension(outputPath));

			File.WriteAllText(caleCurata, JsonConvert.SerializeObject(rezultatCuratat, Formatting.Indented));
			Console.WriteLine($"JSON curatat salvat in {caleCurata}");
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Scrierea JSON-ului curatat a esuat: {ex.Message}");
		}
	}

	/// <summary>
	/// Prompt simplu pentru un răspuns Da/Nu (acceptă 'd/da' sau 'n/nu' și variante în engleză).
	/// Returnează <c>true</c> pentru Da și <c>false</c> pentru Nu (implicit Enter -> Nu).
	/// </summary>
	/// <param name="question">Întrebarea afișată utilizatorului (fără sufixul (d/n)).</param>
	/// <returns>True dacă utilizatorul a confirmat (da), altfel false.</returns>
	private static bool PromptYesNo(string question)
	{
		while (true)
		{
			var raspuns = Prompt(question + " (d/n):");
			if (string.IsNullOrWhiteSpace(raspuns))
			{
				return false;
			}

			raspuns = raspuns.Trim().ToLowerInvariant();
			if (raspuns is "y" or "yes" or "d" or "da")
			{
				return true;
			}

			if (raspuns is "n" or "no" or "nu")
			{
				return false;
			}

			Console.WriteLine("Raspundeti cu 'd' sau 'n'.");
		}
	}

	/// <summary>
	/// Afișează un prompt și citește o linie de la utilizator.
	/// Dacă <paramref name="allowEmpty"/> este <c>false</c>, prompt-ul se repetă până se primește o valoare nenulă.
	/// </summary>
	/// <param name="label">Textul afișat utilizatorului.</param>
	/// <param name="allowEmpty">Permite acceptarea unui input gol dacă este true.</param>
	/// <returns>Valoarea introdusă, eliminând spațiile de la început și sfârșit.</returns>
	private static string Prompt(string label, bool allowEmpty = false)
	{
		while (true)
		{
			Console.Write(label + " ");
			var linie = Console.ReadLine() ?? string.Empty;
			if (allowEmpty || !string.IsNullOrWhiteSpace(linie))
			{
				return linie.Trim();
			}

			Console.WriteLine("Valoare obligatorie.");
		}
	}

	/// <summary>
	/// Extinde variabilele de mediu dintr-o cale și normalizează referințele relative ('.' sau '..')
	/// într-o cale absolută.
	/// </summary>
	/// <param name="path">Calea introdusă de utilizator.</param>
	/// <returns>Calea extinsă și normalizată.</returns>
	private static string ExpandPath(string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return path;
		}

		var caleExpandata = Environment.ExpandEnvironmentVariables(path.Trim());
		if (caleExpandata.StartsWith(".") || caleExpandata.StartsWith(".."))
		{
			caleExpandata = Path.GetFullPath(caleExpandata);
		}

		return caleExpandata;
	}

	/// <summary>
	/// Returnează eticheta (text afișat) asociată unui profil de parsare.
	/// Dacă nu există o etichetă definită, se folosește numele enum-ului.
	/// </summary>
	/// <param name="profile">Profilul pentru care se cere eticheta.</param>
	/// <returns>Eticheta prietenoasă sau numele enum.</returns>
	private static string GetProfileEticheta(DevizParserProfile profile)
	{
		return ProfileEtichete.TryGetValue(profile, out var eticheta)
			? eticheta
			: profile.ToString();
	}
}
