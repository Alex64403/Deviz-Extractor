using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DevizParsing.Core.Models;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;

namespace DevizParsing.Core.Persistence
{
    /// <summary>
    /// Persistă instanțe <see cref="ParseResult"/> într-un tabel SQL Server pentru procesări ulterioare.
    /// </summary>
    /// <remarks>
    /// The writer expects the destination table to expose, at a minimum, the following columns
    /// (additional nullable columns are allowed and will be left as NULL). Related tables must
    /// exist with the suffixes <c>Pozitii</c> and <c>Categorii</c> to store normalized data.
    /// <code>
    /// CREATE TABLE dbo.DevizImportRaw (
    ///     Id            BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     FileName      NVARCHAR(400) NOT NULL,
    ///     ParserProfile NVARCHAR(50)  NOT NULL,
    ///     SheetName     NVARCHAR(255) NULL,
    ///     ImportedAtUtc DATETIME2(3)  NOT NULL DEFAULT SYSUTCDATETIME(),
    ///     [RowCount]    INT           NOT NULL,
    ///     RawJson       NVARCHAR(MAX) NOT NULL,
    ///     MetaJson      NVARCHAR(MAX) NULL,
    ///     SourceHash    VARBINARY(32) NULL,
    ///     Beneficiar    NVARCHAR(400) NULL,
    ///     Executant     NVARCHAR(400) NULL,
    ///     Proiectant    NVARCHAR(400) NULL,
    ///     Obiectiv      NVARCHAR(400) NULL,
    ///     Obiect        NVARCHAR(400) NULL,
    ///     Deviz         NVARCHAR(400) NULL,
    ///     StadiuFizic   NVARCHAR(200) NULL,
    ///     SectiuneTehnica     NVARCHAR(200) NULL,
    ///     SectiuneFinanciara  NVARCHAR(200) NULL,
    ///     DataDocument  NVARCHAR(100) NULL
    /// );
    ///
    /// CREATE TABLE dbo.DevizImportRawPozitii (
    ///     Id                 BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     DocumentId         BIGINT NOT NULL REFERENCES dbo.DevizImportRaw(Id) ON DELETE CASCADE,
    ///     [Order]            NVARCHAR(50) NULL,
    ///     Symbol             NVARCHAR(100) NULL,
    ///     Name               NVARCHAR(400) NOT NULL,
    ///     UnitOfMeasure      NVARCHAR(50) NULL,
    ///     Quantity           DECIMAL(18,4) NOT NULL,
    ///     UnitPrice          DECIMAL(18,4) NOT NULL,
    ///     LineTotal          DECIMAL(18,4) NOT NULL,
    ///     ComputedLineTotal  DECIMAL(18,4) NOT NULL,
    ///     SheetLineTotal     DECIMAL(18,4) NULL,
    ///     Notes              NVARCHAR(MAX) NULL
    /// );
    ///
    /// CREATE TABLE dbo.DevizImportRawPozitiiMateriale (
    ///     Id           BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     ActivityId  BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
    ///     Quantity    DECIMAL(18,4) NOT NULL,
    ///     UnitPrice   DECIMAL(18,4) NOT NULL,
    ///     Total       DECIMAL(18,4) NOT NULL
    /// );
    ///
    /// CREATE TABLE dbo.DevizImportRawPozitiiManopera (
    ///     Id           BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     ActivityId  BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
    ///     Quantity    DECIMAL(18,4) NOT NULL,
    ///     UnitPrice   DECIMAL(18,4) NOT NULL,
    ///     Total       DECIMAL(18,4) NOT NULL
    /// );
    ///
    /// CREATE TABLE dbo.DevizImportRawPozitiiUtilaje (
    ///     Id           BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     ActivityId  BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
    ///     Quantity    DECIMAL(18,4) NOT NULL,
    ///     UnitPrice   DECIMAL(18,4) NOT NULL,
    ///     Total       DECIMAL(18,4) NOT NULL
    /// );
    ///
    /// CREATE TABLE dbo.DevizImportRawPozitiiTransport (
    ///     Id           BIGINT IDENTITY(1,1) PRIMARY KEY,
    ///     ActivityId  BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
    ///     Quantity    DECIMAL(18,4) NOT NULL,
    ///     UnitPrice   DECIMAL(18,4) NOT NULL,
    ///     Total       DECIMAL(18,4) NOT NULL
    /// );
    /// </code>
    /// </remarks>
    public sealed class ParseResultDatabaseWriter
    {
    private readonly string _sirConexiune;
    private readonly string _numeTabelaCuSchema;
    private readonly string _numeTabelaPozitii;
    private readonly string _numeTabelaMateriale;
    private readonly string _numeTabelaManopera;
    private readonly string _numeTabelaUtilaje;
    private readonly string _numeTabelaTransport;

        public ParseResultDatabaseWriter(string connectionString, string? tableName = null)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("Connection string is required.", nameof(connectionString));
            }

            _sirConexiune = connectionString;
            var baza = string.IsNullOrWhiteSpace(tableName) ? "DevizImportRaw" : tableName!;
            _numeTabelaCuSchema = QualifyTableName(baza);
            _numeTabelaPozitii = QualifyTableName(AppendSuffix(baza, "Pozitii"));
            _numeTabelaMateriale = QualifyTableName(AppendSuffix(baza, "PozitiiMateriale"));
            _numeTabelaManopera = QualifyTableName(AppendSuffix(baza, "PozitiiManopera"));
            _numeTabelaUtilaje = QualifyTableName(AppendSuffix(baza, "PozitiiUtilaje"));
            _numeTabelaTransport = QualifyTableName(AppendSuffix(baza, "PozitiiTransport"));
        }

    /// <summary>
    /// Salvează rezultatul de parsare și returnează identificatorul generat (dacă tabela expune OUTPUT INSERTED.Id).
    /// </summary>
    /// <param name="result">Rezultatul ce trebuie persistat.</param>
    /// <param name="profile">Numele profilului de parsare folosit.</param>
    /// <param name="cancellationToken">Token opțional pentru anularea operației asincrone.</param>
    /// <returns>Id-ul inserat sau <c>null</c> dacă tabela nu îl furnizează.</returns>
    public async Task<int?> SalveazaAsync(ParseResult result, string profile, CancellationToken cancellationToken = default)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            if (string.IsNullOrWhiteSpace(profile))
            {
                throw new ArgumentException("Profile is required.", nameof(profile));
            }

            var jsonDeSalvat = JsonConvert.SerializeObject(result, Formatting.None);
            var metaJson = JsonConvert.SerializeObject(result.Metadata ?? new DevizMetadata(), Formatting.None);
            await using var connection = new SqlConnection(_sirConexiune);
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

            await using var tranzactie = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                var sqlInsert = $@"INSERT INTO {_numeTabelaCuSchema}
    (FileName, ParserProfile, SheetName, ImportedAtUtc, [RowCount], RawJson, MetaJson, SourceHash, Beneficiar, Executant, Proiectant, Obiectiv, Obiect, Deviz, StadiuFizic, SectiuneTehnica, SectiuneFinanciara, DataDocument)
OUTPUT INSERTED.Id
VALUES
    (@FileName, @ParserProfile, @SheetName, SYSUTCDATETIME(), @RowCount, @RawJson, @MetaJson, @SourceHash, @Beneficiar, @Executant, @Proiectant, @Obiectiv, @Obiect, @Deviz, @StadiuFizic, @SectiuneTehnica, @SectiuneFinanciara, @DataDocument);";

                await using var comanda = new SqlCommand(sqlInsert, connection, tranzactie);
                comanda.Parameters.Add(new SqlParameter("@FileName", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)result.SourceFile ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@ParserProfile", SqlDbType.NVarChar, 50)
                {
                    Value = profile
                });
                comanda.Parameters.Add(new SqlParameter("@SheetName", SqlDbType.NVarChar, 255)
                {
                    Value = (object?)result.Sheet ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@RowCount", SqlDbType.Int)
                {
                    Value = result.Rows?.Count ?? 0
                });
                comanda.Parameters.Add(new SqlParameter("@RawJson", SqlDbType.NVarChar, -1)
                {
                    Value = jsonDeSalvat
                });
                comanda.Parameters.Add(new SqlParameter("@MetaJson", SqlDbType.NVarChar, -1)
                {
                    Value = metaJson
                });
                comanda.Parameters.Add(new SqlParameter("@SourceHash", SqlDbType.VarBinary, 32)
                {
                    Value = DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Beneficiar", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Beneficiar) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Executant", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Executant) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Proiectant", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Proiectant) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Obiectiv", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Obiectiv) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Obiect", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Obiect) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@Deviz", SqlDbType.NVarChar, 400)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.Deviz) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@StadiuFizic", SqlDbType.NVarChar, 200)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.StadiuFizic) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@SectiuneTehnica", SqlDbType.NVarChar, 200)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.SectiuneTehnica) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@SectiuneFinanciara", SqlDbType.NVarChar, 200)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.SectiuneFinanciara) ?? DBNull.Value
                });
                comanda.Parameters.Add(new SqlParameter("@DataDocument", SqlDbType.NVarChar, 100)
                {
                    Value = (object?)NullIfEmpty(result.Metadata?.DataDocument) ?? DBNull.Value
                });

                var idInserat = await comanda.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);
                var documentId = idInserat switch
                {
                    int id => (long)id,
                    long longId => longId,
                    null => 0L,
                    _ => Convert.ToInt64(idInserat, CultureInfo.InvariantCulture)
                };

                if (documentId == 0)
                {
                    await tranzactie.RollbackAsync(cancellationToken).ConfigureAwait(false);
                    return null;
                }

                await InserarePozitiiAsync(connection, tranzactie, documentId, result, cancellationToken).ConfigureAwait(false);

                await tranzactie.CommitAsync(cancellationToken).ConfigureAwait(false);

                return documentId <= int.MaxValue ? (int?)documentId : null;
            }
            catch
            {
                await tranzactie.RollbackAsync(CancellationToken.None).ConfigureAwait(false);
                throw;
            }
        }

    /// <summary>
    /// Normalizează numele tabelei într-un format sigur cu identificatori între paranteze pătrate.
    /// </summary>
    /// <param name="rawName">Numele introdus de utilizator (poate include schema).</param>
    /// <returns>Numele complet calificat potrivit pentru SQL Server.</returns>
    private static string QualifyTableName(string rawName)
    {
            var parts = rawName
                .Split('.', StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim().Trim('[', ']'))
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToArray();

            if (parts.Length == 0)
            {
                throw new ArgumentException("Table name must contain at least one identifier.", nameof(rawName));
            }

            return string.Join('.', parts.Select(p => $"[{p}]").ToArray());
        }

        private static string AppendSuffix(string rawName, string suffix)
        {
            if (string.IsNullOrWhiteSpace(rawName))
            {
                throw new ArgumentException("Base table name must not be empty.", nameof(rawName));
            }

            var parts = rawName.Split('.', StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim().Trim('[', ']')).ToArray();
            if (parts.Length == 0)
            {
                throw new ArgumentException("Base table name must contain at least one identifier.", nameof(rawName));
            }

            var last = parts[^1] + suffix;
            parts[^1] = last;
            return string.Join('.', parts.Select(p => $"[{p}]").ToArray());
        }

        private static string? NullIfEmpty(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private async Task InserarePozitiiAsync(SqlConnection connection, SqlTransaction tranzactie, long documentId, ParseResult result, CancellationToken cancellationToken)
        {
            if (result.Rows == null || result.Rows.Count == 0)
            {
                return;
            }

            var sqlPozitie = $@"INSERT INTO {_numeTabelaPozitii}
    (DocumentId, [Order], Symbol, Name, UnitOfMeasure, Quantity, UnitPrice, LineTotal, ComputedLineTotal, SheetLineTotal, Notes)
OUTPUT INSERTED.Id
VALUES
    (@DocumentId, @Order, @Symbol, @Name, @UnitOfMeasure, @Quantity, @UnitPrice, @LineTotal, @ComputedLineTotal, @SheetLineTotal, @Notes);";

            await using var comandaPozitii = new SqlCommand(sqlPozitie, connection, tranzactie);
            var paramDocumentId = comandaPozitii.Parameters.Add("@DocumentId", SqlDbType.BigInt);
            var paramOrder = comandaPozitii.Parameters.Add("@Order", SqlDbType.NVarChar, 50);
            var paramSymbol = comandaPozitii.Parameters.Add("@Symbol", SqlDbType.NVarChar, 100);
            var paramName = comandaPozitii.Parameters.Add("@Name", SqlDbType.NVarChar, 400);
            var paramUnit = comandaPozitii.Parameters.Add("@UnitOfMeasure", SqlDbType.NVarChar, 50);
            var paramQuantity = comandaPozitii.Parameters.Add("@Quantity", SqlDbType.Decimal);
            paramQuantity.Precision = 18;
            paramQuantity.Scale = 4;
            var paramUnitPrice = comandaPozitii.Parameters.Add("@UnitPrice", SqlDbType.Decimal);
            paramUnitPrice.Precision = 18;
            paramUnitPrice.Scale = 4;
            var paramLineTotal = comandaPozitii.Parameters.Add("@LineTotal", SqlDbType.Decimal);
            paramLineTotal.Precision = 18;
            paramLineTotal.Scale = 4;
            var paramComputedLine = comandaPozitii.Parameters.Add("@ComputedLineTotal", SqlDbType.Decimal);
            paramComputedLine.Precision = 18;
            paramComputedLine.Scale = 4;
            var paramSheetLine = comandaPozitii.Parameters.Add("@SheetLineTotal", SqlDbType.Decimal);
            paramSheetLine.Precision = 18;
            paramSheetLine.Scale = 4;
            var paramNotes = comandaPozitii.Parameters.Add("@Notes", SqlDbType.NVarChar, -1);

            paramDocumentId.Value = documentId;

            var materialeCmd = CreateCategoryCommand(connection, tranzactie, _numeTabelaMateriale);
            await using var comandaMateriale = materialeCmd.Command;
            var paramMatActivity = materialeCmd.Activity;
            var paramMatQuantity = materialeCmd.Quantity;
            var paramMatUnitPrice = materialeCmd.UnitPrice;
            var paramMatTotal = materialeCmd.Total;

            var manoperaCmd = CreateCategoryCommand(connection, tranzactie, _numeTabelaManopera);
            await using var comandaManopera = manoperaCmd.Command;
            var paramManActivity = manoperaCmd.Activity;
            var paramManQuantity = manoperaCmd.Quantity;
            var paramManUnitPrice = manoperaCmd.UnitPrice;
            var paramManTotal = manoperaCmd.Total;

            var utilajeCmd = CreateCategoryCommand(connection, tranzactie, _numeTabelaUtilaje);
            await using var comandaUtilaje = utilajeCmd.Command;
            var paramUtilActivity = utilajeCmd.Activity;
            var paramUtilQuantity = utilajeCmd.Quantity;
            var paramUtilUnitPrice = utilajeCmd.UnitPrice;
            var paramUtilTotal = utilajeCmd.Total;

            var transportCmd = CreateCategoryCommand(connection, tranzactie, _numeTabelaTransport);
            await using var comandaTransport = transportCmd.Command;
            var paramTransActivity = transportCmd.Activity;
            var paramTransQuantity = transportCmd.Quantity;
            var paramTransUnitPrice = transportCmd.UnitPrice;
            var paramTransTotal = transportCmd.Total;

            foreach (var rand in result.Rows)
            {
                paramOrder.Value = (object?)NullIfEmpty(rand.Order) ?? DBNull.Value;
                paramSymbol.Value = (object?)NullIfEmpty(rand.Symbol) ?? DBNull.Value;
                paramName.Value = (object?)NullIfEmpty(rand.Name) ?? DBNull.Value;
                paramUnit.Value = (object?)NullIfEmpty(rand.UnitOfMeasure) ?? DBNull.Value;
                paramQuantity.Value = rand.Quantity;
                paramUnitPrice.Value = rand.UnitPrice;
                paramLineTotal.Value = rand.LineTotal;
                paramComputedLine.Value = rand.ComputedLineTotal;
                paramSheetLine.Value = rand.SheetLineTotal.HasValue ? rand.SheetLineTotal.Value : (object)DBNull.Value;
                paramNotes.Value = (object?)NullIfEmpty(rand.Notes) ?? DBNull.Value;

                var pozIdObj = await comandaPozitii.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);
                var pozId = pozIdObj switch
                {
                    int id => (long)id,
                    long longId => longId,
                    null => 0L,
                    _ => Convert.ToInt64(pozIdObj, CultureInfo.InvariantCulture)
                };

                if (pozId == 0)
                {
                    continue;
                }

                await InserareCategorieAsync(rand.Categories.Materials, pozId, paramMatActivity, paramMatQuantity, paramMatUnitPrice, paramMatTotal, comandaMateriale, cancellationToken).ConfigureAwait(false);
                await InserareCategorieAsync(rand.Categories.Labor, pozId, paramManActivity, paramManQuantity, paramManUnitPrice, paramManTotal, comandaManopera, cancellationToken).ConfigureAwait(false);
                await InserareCategorieAsync(rand.Categories.Equipment, pozId, paramUtilActivity, paramUtilQuantity, paramUtilUnitPrice, paramUtilTotal, comandaUtilaje, cancellationToken).ConfigureAwait(false);
                await InserareCategorieAsync(rand.Categories.Transport, pozId, paramTransActivity, paramTransQuantity, paramTransUnitPrice, paramTransTotal, comandaTransport, cancellationToken).ConfigureAwait(false);
            }
        }
        private static async Task InserareCategorieAsync(Category categorie, long pozId, SqlParameter paramActivityId, SqlParameter paramQuantity, SqlParameter paramUnitPrice, SqlParameter paramTotal, SqlCommand comandaCategorie, CancellationToken cancellationToken)
        {
            if (!ArTrebuiaPersistataCategorie(categorie))
            {
                return;
            }

            paramActivityId.Value = pozId;
            paramQuantity.Value = categorie.Quantity;
            paramUnitPrice.Value = categorie.UnitPrice;
            paramTotal.Value = categorie.Total;
            await comandaCategorie.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }

        private static (SqlCommand Command, SqlParameter Activity, SqlParameter Quantity, SqlParameter UnitPrice, SqlParameter Total) CreateCategoryCommand(SqlConnection connection, SqlTransaction tranzactie, string tableName)
        {
            var sqlCategorie = $@"INSERT INTO {tableName}
    (ActivityId, Quantity, UnitPrice, Total)
VALUES
    (@ActivityId, @Quantity, @UnitPrice, @Total);";

            var comanda = new SqlCommand(sqlCategorie, connection, tranzactie);
            var paramActivityId = comanda.Parameters.Add("@ActivityId", SqlDbType.BigInt);
            var paramQuantity = comanda.Parameters.Add("@Quantity", SqlDbType.Decimal);
            paramQuantity.Precision = 18;
            paramQuantity.Scale = 4;
            var paramUnitPrice = comanda.Parameters.Add("@UnitPrice", SqlDbType.Decimal);
            paramUnitPrice.Precision = 18;
            paramUnitPrice.Scale = 4;
            var paramTotal = comanda.Parameters.Add("@Total", SqlDbType.Decimal);
            paramTotal.Precision = 18;
            paramTotal.Scale = 4;

            return (comanda, paramActivityId, paramQuantity, paramUnitPrice, paramTotal);
        }

        private static bool ArTrebuiaPersistataCategorie(Category categorie)
        {
            return categorie.Quantity != 0m || categorie.UnitPrice != 0m || categorie.Total != 0m;
        }
    }
}
