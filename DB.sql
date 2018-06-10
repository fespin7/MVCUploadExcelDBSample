IF EXISTS(SELECT 1 FROM sys.tables WHERE object_id = OBJECT_ID('People'))
BEGIN;
    DROP TABLE [People];
END;
GO

CREATE TABLE [People] (
    [Identification] VARCHAR(13) NULL,
    [Name] VARCHAR(255) NULL,
    [BirthDate] DATETIME NULL,
    [Phone] VARCHAR(100) NULL,
    [Email] VARCHAR(255) NULL,
    [City] VARCHAR(255) NULL,
    [ZipCode] VARCHAR(10) NULL,
    [Region] VARCHAR(50) NULL,
    [Country] VARCHAR(100) NULL,
    [Company] VARCHAR(255) NULL,
    [Assets] DECIMAL(18,4) NULL
);
GO
