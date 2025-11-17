using System;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DevizParsing.Core.Models;
using Microsoft.Data.SqlClient;

namespace DevizParsing.Core.Persistence
{
    /// <summary>
    /// Persistă fiecare rând parse într-un tabel SQL Server denormalizat (staging) cu coloane explicite.
    /// </summary>
    public sealed class ParseResultFlatDatabaseWriter
    {
        private readonly string _connectionString;
        private readonly string _destinationTable;

        public ParseResultFlatDatabaseWriter(string connectionString, string? tableName = null)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("Connection string is required.", nameof(connectionString));
            }

            _connectionString = connectionString;
            _destinationTable = QualifyTableName(string.IsNullOrWhiteSpace(tableName) ? "DevizImportStage" : tableName!);
        }

        public async Task<int> SalveazaRanduriAsync(ParseResult result, string profile, CancellationToken cancellationToken = default)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            if (string.IsNullOrWhiteSpace(profile))
            {
                throw new ArgumentException("Profile is required.", nameof(profile));
            }

            var table = BuildDataTable(result, profile);
            if (table.Rows.Count == 0)
            {
                return 0;
            }

            await using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

            using var bulkCopy = new SqlBulkCopy(connection)
            {
                DestinationTableName = _destinationTable,
                BulkCopyTimeout = 120
            };

            foreach (DataColumn column in table.Columns)
            {
                bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
            }

            await bulkCopy.WriteToServerAsync(table, cancellationToken).ConfigureAwait(false);
            return table.Rows.Count;
        }

        private static DataTable BuildDataTable(ParseResult result, string profile)
        {
            var table = new DataTable();
            table.Columns.Add("SourceFile", typeof(string));
            table.Columns.Add("SheetName", typeof(string));
            table.Columns.Add("ParserProfile", typeof(string));
            table.Columns.Add("ImportedAtUtc", typeof(DateTime));
            table.Columns.Add("Beneficiar", typeof(string));
            table.Columns.Add("Executant", typeof(string));
            table.Columns.Add("Proiectant", typeof(string));
            table.Columns.Add("Obiectiv", typeof(string));
            table.Columns.Add("Obiect", typeof(string));
            table.Columns.Add("Deviz", typeof(string));
            table.Columns.Add("StadiuFizic", typeof(string));
            table.Columns.Add("DataDocument", typeof(string));
            table.Columns.Add("RowNumber", typeof(string));
            table.Columns.Add("Symbol", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("UnitOfMeasure", typeof(string));
            table.Columns.Add("Quantity", typeof(decimal));
            table.Columns.Add("UnitPrice", typeof(decimal));
            table.Columns.Add("LineTotal", typeof(decimal));
            table.Columns.Add("ComputedLineTotal", typeof(decimal));
            table.Columns.Add("SheetLineTotal", typeof(decimal));
            table.Columns.Add("MaterialsQuantity", typeof(decimal));
            table.Columns.Add("MaterialsUnitPrice", typeof(decimal));
            table.Columns.Add("MaterialsTotal", typeof(decimal));
            table.Columns.Add("LaborQuantity", typeof(decimal));
            table.Columns.Add("LaborUnitPrice", typeof(decimal));
            table.Columns.Add("LaborTotal", typeof(decimal));
            table.Columns.Add("EquipmentQuantity", typeof(decimal));
            table.Columns.Add("EquipmentUnitPrice", typeof(decimal));
            table.Columns.Add("EquipmentTotal", typeof(decimal));
            table.Columns.Add("TransportQuantity", typeof(decimal));
            table.Columns.Add("TransportUnitPrice", typeof(decimal));
            table.Columns.Add("TransportTotal", typeof(decimal));
            table.Columns.Add("Notes", typeof(string));

            var importedAt = DateTime.UtcNow;
            var metadata = result.Metadata ?? new DevizMetadata();

            foreach (var row in result.Rows)
            {
                var dataRow = table.NewRow();
                dataRow["SourceFile"] = DbValueOrNull(result.SourceFile);
                dataRow["SheetName"] = DbValueOrNull(result.Sheet);
                dataRow["ParserProfile"] = profile;
                dataRow["ImportedAtUtc"] = importedAt;
                dataRow["Beneficiar"] = DbValueOrNull(metadata.Beneficiar);
                dataRow["Executant"] = DbValueOrNull(metadata.Executant);
                dataRow["Proiectant"] = DbValueOrNull(metadata.Proiectant);
                dataRow["Obiectiv"] = DbValueOrNull(metadata.Obiectiv);
                dataRow["Obiect"] = DbValueOrNull(metadata.Obiect);
                dataRow["Deviz"] = DbValueOrNull(string.IsNullOrWhiteSpace(metadata.Deviz) ? metadata.StadiuFizic : metadata.Deviz);
                dataRow["StadiuFizic"] = DbValueOrNull(metadata.StadiuFizic);
                dataRow["DataDocument"] = DbValueOrNull(metadata.DataDocument);
                dataRow["RowNumber"] = DbValueOrNull(row.Order);
                dataRow["Symbol"] = DbValueOrNull(row.Symbol);
                dataRow["Name"] = DbValueOrNull(row.Name);
                dataRow["UnitOfMeasure"] = DbValueOrNull(row.UnitOfMeasure);
                dataRow["Quantity"] = row.Quantity;
                dataRow["UnitPrice"] = row.UnitPrice;
                dataRow["LineTotal"] = row.LineTotal;
                dataRow["ComputedLineTotal"] = row.ComputedLineTotal;
                dataRow["SheetLineTotal"] = row.SheetLineTotal.HasValue ? row.SheetLineTotal.Value : (object)DBNull.Value;
                dataRow["MaterialsQuantity"] = row.Categories.Materials.Quantity;
                dataRow["MaterialsUnitPrice"] = row.Categories.Materials.UnitPrice;
                dataRow["MaterialsTotal"] = row.Categories.Materials.Total;
                dataRow["LaborQuantity"] = row.Categories.Labor.Quantity;
                dataRow["LaborUnitPrice"] = row.Categories.Labor.UnitPrice;
                dataRow["LaborTotal"] = row.Categories.Labor.Total;
                dataRow["EquipmentQuantity"] = row.Categories.Equipment.Quantity;
                dataRow["EquipmentUnitPrice"] = row.Categories.Equipment.UnitPrice;
                dataRow["EquipmentTotal"] = row.Categories.Equipment.Total;
                dataRow["TransportQuantity"] = row.Categories.Transport.Quantity;
                dataRow["TransportUnitPrice"] = row.Categories.Transport.UnitPrice;
                dataRow["TransportTotal"] = row.Categories.Transport.Total;
                dataRow["Notes"] = DbValueOrNull(row.Notes);
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static object DbValueOrNull(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? DBNull.Value : value.Trim();
        }

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
    }
}
