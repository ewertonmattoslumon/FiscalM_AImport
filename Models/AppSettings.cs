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
        public bool ImportLead { get; set; }
        public bool ImportContact { get; set; }
        public bool ImportAccount { get; set; }
        public string LeadExcelFile { get; set; } = "leads.xlsx";
        public string ContactExcelFile { get; set; } = "contacts.xlsx";
        public string AccountExcelFile { get; set; } = "accounts.xlsx";
    }
}
