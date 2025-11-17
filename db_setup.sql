IF OBJECT_ID('dbo.DevizImportRaw', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRaw (
        Id               BIGINT IDENTITY(1,1) PRIMARY KEY,
        FileName         NVARCHAR(400) NOT NULL,
        ParserProfile    NVARCHAR(50)  NOT NULL,
        SheetName        NVARCHAR(255) NULL,
        ImportedAtUtc    DATETIME2(3)  NOT NULL DEFAULT SYSUTCDATETIME(),
        [RowCount]       INT           NOT NULL,
        RawJson          NVARCHAR(MAX) NOT NULL,
        MetaJson         NVARCHAR(MAX) NULL,
        SourceHash       VARBINARY(32) NULL,
        Beneficiar       NVARCHAR(400) NULL,
        Executant        NVARCHAR(400) NULL,
        Proiectant       NVARCHAR(400) NULL,
        Obiectiv         NVARCHAR(400) NULL,
        Obiect           NVARCHAR(400) NULL,
        Deviz            NVARCHAR(400) NULL,
        StadiuFizic      NVARCHAR(200) NULL,
        SectiuneTehnica  NVARCHAR(200) NULL,
        SectiuneFinanciara NVARCHAR(200) NULL,
        DataDocument     NVARCHAR(100) NULL
    );
END;

IF COL_LENGTH('dbo.DevizImportRaw', 'Beneficiar') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Beneficiar NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'Executant') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Executant NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'Proiectant') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Proiectant NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'Obiectiv') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Obiectiv NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'Obiect') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Obiect NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'Deviz') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD Deviz NVARCHAR(400) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'StadiuFizic') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD StadiuFizic NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'SectiuneTehnica') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD SectiuneTehnica NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'SectiuneFinanciara') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD SectiuneFinanciara NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.DevizImportRaw', 'DataDocument') IS NULL
    ALTER TABLE dbo.DevizImportRaw ADD DataDocument NVARCHAR(100) NULL;

IF OBJECT_ID('dbo.DevizImportRawPozitii', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRawPozitii (
        Id                BIGINT IDENTITY(1,1) PRIMARY KEY,
        DocumentId        BIGINT NOT NULL REFERENCES dbo.DevizImportRaw(Id) ON DELETE CASCADE,
        [Order]           NVARCHAR(50) NULL,
        Symbol            NVARCHAR(100) NULL,
        Name              NVARCHAR(400) NOT NULL,
        UnitOfMeasure     NVARCHAR(50) NULL,
        Quantity          DECIMAL(18,4) NOT NULL,
        UnitPrice         DECIMAL(18,4) NOT NULL,
        LineTotal         DECIMAL(18,4) NOT NULL,
        ComputedLineTotal DECIMAL(18,4) NOT NULL,
        SheetLineTotal    DECIMAL(18,4) NULL,
        Notes             NVARCHAR(MAX) NULL
    );
END;

IF OBJECT_ID('dbo.DevizImportRawPozitiiMateriale', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRawPozitiiMateriale (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizImportRawPozitiiManopera', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRawPozitiiManopera (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizImportRawPozitiiUtilaje', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRawPozitiiUtilaje (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizImportRawPozitiiTransport', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizImportRawPozitiiTransport (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizImportRawPozitii(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizDocument', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizDocument (
        Id                 BIGINT IDENTITY(1,1) PRIMARY KEY,
        ImportId           BIGINT NULL REFERENCES dbo.DevizImportRaw(Id) ON DELETE SET NULL,
        FileName           NVARCHAR(400) NOT NULL,
        ParserProfile      NVARCHAR(50)  NULL,
        SheetName          NVARCHAR(255) NULL,
        Beneficiar         NVARCHAR(400) NULL,
        Obiectiv           NVARCHAR(400) NULL,
        Obiect             NVARCHAR(400) NULL,
        Executant          NVARCHAR(400) NULL,
        Proiectant         NVARCHAR(400) NULL,
        Deviz              NVARCHAR(400) NULL,
        StadiuFizic        NVARCHAR(200) NULL,
        SectiuneTehnica    NVARCHAR(200) NULL,
        SectiuneFinanciara NVARCHAR(200) NULL,
        DataDocument       NVARCHAR(100) NULL,
        CreatedAtUtc       DATETIME2(3)  NOT NULL DEFAULT SYSUTCDATETIME()
    );
END;

IF OBJECT_ID('dbo.DevizActivitate', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizActivitate (
        Id                BIGINT IDENTITY(1,1) PRIMARY KEY,
        DocumentId        BIGINT NOT NULL REFERENCES dbo.DevizDocument(Id) ON DELETE CASCADE,
        [Order]           NVARCHAR(50) NULL,
        Symbol            NVARCHAR(100) NULL,
        Name              NVARCHAR(400) NOT NULL,
        UnitOfMeasure     NVARCHAR(50) NULL,
        Quantity          DECIMAL(18,4) NOT NULL,
        UnitPrice         DECIMAL(18,4) NOT NULL,
        LineTotal         DECIMAL(18,4) NOT NULL,
        ComputedLineTotal DECIMAL(18,4) NOT NULL,
        SheetLineTotal    DECIMAL(18,4) NULL,
        Notes             NVARCHAR(MAX) NULL
    );
END;

IF OBJECT_ID('dbo.DevizActivitateMateriale', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizActivitateMateriale (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizActivitate(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizActivitateManopera', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizActivitateManopera (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizActivitate(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizActivitateUtilaje', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizActivitateUtilaje (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizActivitate(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizActivitateTransport', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizActivitateTransport (
        Id         BIGINT IDENTITY(1,1) PRIMARY KEY,
        ActivityId BIGINT NOT NULL REFERENCES dbo.DevizActivitate(Id) ON DELETE CASCADE,
        Quantity   DECIMAL(18,4) NOT NULL,
        UnitPrice  DECIMAL(18,4) NOT NULL,
        Total      DECIMAL(18,4) NOT NULL
    );
END;

IF OBJECT_ID('dbo.DevizDocumentTotaluri', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DevizDocumentTotaluri (
        DocumentId  BIGINT NOT NULL PRIMARY KEY REFERENCES dbo.DevizDocument(Id) ON DELETE CASCADE,
        TotalMateriale DECIMAL(18,4) NOT NULL DEFAULT 0,
        TotalManopera  DECIMAL(18,4) NOT NULL DEFAULT 0,
        TotalUtilaje   DECIMAL(18,4) NOT NULL DEFAULT 0,
        TotalTransport DECIMAL(18,4) NOT NULL DEFAULT 0,
        TotalGeneral   DECIMAL(18,4) NOT NULL DEFAULT 0
    );
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_DevizActivitate_DocumentId'
      AND object_id = OBJECT_ID('dbo.DevizActivitate')
)
BEGIN
    CREATE INDEX IX_DevizActivitate_DocumentId ON dbo.DevizActivitate(DocumentId);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_DevizActivitateMateriale_ActivityId'
      AND object_id = OBJECT_ID('dbo.DevizActivitateMateriale')
)
BEGIN
    CREATE INDEX IX_DevizActivitateMateriale_ActivityId ON dbo.DevizActivitateMateriale(ActivityId);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_DevizActivitateManopera_ActivityId'
      AND object_id = OBJECT_ID('dbo.DevizActivitateManopera')
)
BEGIN
    CREATE INDEX IX_DevizActivitateManopera_ActivityId ON dbo.DevizActivitateManopera(ActivityId);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_DevizActivitateUtilaje_ActivityId'
      AND object_id = OBJECT_ID('dbo.DevizActivitateUtilaje')
)
BEGIN
    CREATE INDEX IX_DevizActivitateUtilaje_ActivityId ON dbo.DevizActivitateUtilaje(ActivityId);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_DevizActivitateTransport_ActivityId'
      AND object_id = OBJECT_ID('dbo.DevizActivitateTransport')
)
BEGIN
    CREATE INDEX IX_DevizActivitateTransport_ActivityId ON dbo.DevizActivitateTransport(ActivityId);
END;
