namespace FiscalM_AImport.Models
{
    public class AppSettings
    {
        public DataverseSettings Dataverse { get; set; } = new DataverseSettings();
        public ImportSettings Import { get; set; } = new ImportSettings();
    }

    public class DataverseSettings
    {
        public string ConnectionString { get; set; } = string.Empty;
    }

    public class ImportSettings
    {
        public int FieldNamesRow { get; set; } = 2;
        public string ExcelFile { get; set; } = "import.xlsx";
    }
}
