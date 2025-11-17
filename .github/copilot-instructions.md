## Copilot / AI agent instructions — DevizParsing repo

Purpose (short)
- Help an AI coding agent become productive quickly in this repository: understand the architecture, developer workflows, important files, conventions and integration points. Use examples from the codebase when possible.

Big picture (what this codebase does and why)
- This repo parses Excel "deviz" (bill-of-quantities) files into a structured `ParseResult`, validates totals and optionally persists both the raw JSON and a normalized representation into SQL Server.
- Design choices to be aware of:
  - Keep a raw document row (`DevizImportRaw`) that stores `RawJson` + `MetaJson` for traceability and debugging.
  - Normalize rows into two related tables with suffixes: `<TableName>_Pozitii` (positions) and `<TableName>_Categorii` (category breakdowns). This is implemented transactionally so an import is atomic.

Key files and where to look first
- `DevizParsing.Core\Excel\DevizWorksheetParser.cs` — main parser: header detection, metadata extraction, fallback positional column mapping, category splitting, numeric parsing, and validation logic. Read this to understand parsing behavior and heuristics.
- `DevizParsing.Core\Persistence\ParseResultDatabaseWriter.cs` — DB writer: transactional insert into `DevizImportRaw`, `_Pozitii` and `_Categorii` tables. Shows SQL used and naming conventions (suffixing). Good for changes to persistence.
- `db_setup.sql` — schema creation/ALTER statements. Use it to verify database expectations (table and column names).
- CLI projects (ex: `ExcelToJsonParser`, `Deviz360DevizToJson`, `RacsadiaDevizToJson`) — thin wrappers that call the core parser and optionally persist. Their `Program.cs` shows CLI options (`--file`, `--save-db`, `--table`, `--db-connection`).

Developer workflows (quick commands)
- Build all projects:

```powershell
dotnet build
```

- Run the parser CLI (parse + save to DB). Replace file path and connection string:

```powershell
dotnet run --project .\ExcelToJsonParser -- --file "C:\cale\la\deviz.xlsx" --save-db --db-connection "Server=.\SQLEXPRESS;Database=Deviz;Integrated Security=True;TrustServerCertificate=True"
```

- Run DB setup script (ensure tables/columns exist):

```powershell
sqlcmd -S .\SQLEXPRESS -i .\db_setup.sql
```

- Quick DB check to list repo tables:

```powershell
sqlcmd -S .\SQLEXPRESS -d Deviz -Q "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE 'DevizImportRaw%';"
```

Project-specific conventions and patterns
- Table naming: default document table is `dbo.DevizImportRaw`. The writer appends `_Pozitii` and `_Categorii` to the base name passed via `--table`. If you change table names, update both writer and schema.
- Parser heuristics: header detection tries to match column label tokens (see arrays near top of `DevizWorksheetParser.cs`); if not confident it falls back to positional mapping based on `DevizParserOptions.Profile` or `CustomFallbackColumns`.
- Metadata extraction: parser scans up to `MetadataScanLimit` (50) rows above header and cleans keys via `NormalizeMetadataKey`. Extra unknown key/value pairs go into `ParseResult.Metadata.Extra`.
- Numeric parsing: helper methods try cell numeric value, then parse text using en-US and fr-FR cultures and finally strip non-digit characters — expect tolerant numeric parsing but watch locale edge cases.
- Validation: computed totals are stored in `ParseResult.ComputedTotals` and validation issues in `ParseResult.Validation`. The CLI outputs these in JSON; use them to triage mismatches.

Integration points and dependencies
- ClosedXML — reads Excel workbooks (`XLWorkbook`), used in the parser.
- Newtonsoft.Json — JSON serialization/deserialization of `ParseResult` and metadata.
- SQL Server — persistence target. Connection strings often use Integrated Security and `TrustServerCertificate=True` in examples.

Where to change behavior safely
- Parsing tweaks (column matching, category tokens, numeric parsing): update `DevizWorksheetParser.cs`.
- Persistence behavior (transactional logic, table suffixes, column mappings): update `ParseResultDatabaseWriter.cs` and mirror changes in `db_setup.sql`.
- CLI flags / help text: update `Program.cs` files under CLI projects.

Examples of common changes and where to implement them
- Add a new metadata field to `DevizImportRaw`: add property in `DevizMetadata` model, ensure `DevizWorksheetParser.ExtractMetadata` populates it, update `db_setup.sql` to `ALTER TABLE ... ADD ...`, and update `ParseResultDatabaseWriter` to persist it.
- Change fallback column mapping for a profile: edit `IaColoaneFallback()` in `DevizWorksheetParser.cs` or provide `CustomFallbackColumns` through `DevizParserOptions`.

Testing and verification notes
- There are no explicit unit tests in the repository (search for `*.Tests` or test projects). For quick verification, use a known Excel file and run CLI with `--out` to produce JSON, then inspect `ParseResult.Validation` and `ComputedTotals`.
- When changing DB schema, run `db_setup.sql` and then import a small test file to validate persistent writes succeed and counts match expectations (e.g., positions and categories).

If you modify this file
- Keep the file concise; include only repository-specific guidance and examples as above. Do not repeat generic coding advice.

Next steps / feedback
- If any area is unclear, tell me which file or workflow you want expanded and I will iterate.
