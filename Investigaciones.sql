CREATE TABLE Investigaciones (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Prompt NVARCHAR(MAX),
    Respuesta NVARCHAR(MAX),
    Fecha DATETIME
);