using Microsoft.PowerPlatform.Dataverse.Client;

namespace FiscalM_AImport.Importers
{
    public class ContactImporter : BaseEntityImporter
    {
        protected override string EntityLogicalName => "contact";

        public ContactImporter(ServiceClient serviceClient, string baseDir, string excelFileName)
            : base(serviceClient, baseDir, excelFileName)
        {
        }
    }
}
