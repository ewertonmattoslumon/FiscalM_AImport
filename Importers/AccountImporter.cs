using Microsoft.PowerPlatform.Dataverse.Client;

namespace FiscalM_AImport.Importers
{
    public class AccountImporter : BaseEntityImporter
    {
        protected override string EntityLogicalName => "account";

        public AccountImporter(ServiceClient serviceClient, string baseDir, string excelFileName)
            : base(serviceClient, baseDir, excelFileName)
        {
        }
    }
}
