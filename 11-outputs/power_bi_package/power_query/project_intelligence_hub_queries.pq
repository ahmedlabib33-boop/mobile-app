// Power Query M - update WorkspaceRoot to your local Project Intelligence Hub folder.
let
    WorkspaceRoot = "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\11-outputs\power_bi_package\data\",
    Projects = Csv.Document(File.Contents(WorkspaceRoot & "projects_registry.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    ProjectsPromoted = Table.PromoteHeaders(Projects, [PromoteAllScalars=true]),
    SourceFiles = Csv.Document(File.Contents(WorkspaceRoot & "source_file_register.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    SourceFilesPromoted = Table.PromoteHeaders(SourceFiles, [PromoteAllScalars=true]),
    Claims = Csv.Document(File.Contents(WorkspaceRoot & "claims_intelligence_summary.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    ClaimsPromoted = Table.PromoteHeaders(Claims, [PromoteAllScalars=true]),
    Clauses = Csv.Document(File.Contents(WorkspaceRoot & "contract_clauses.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    ClausesPromoted = Table.PromoteHeaders(Clauses, [PromoteAllScalars=true])
in
    [Projects=ProjectsPromoted, SourceFiles=SourceFilesPromoted, Claims=ClaimsPromoted, Clauses=ClausesPromoted]
