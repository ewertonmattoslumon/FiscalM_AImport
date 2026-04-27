using Microsoft.PowerPlatform.Dataverse.Client;

namespace FiscalM_AImport.Importers
{
    public class LeadImporter : BaseEntityImporter
    {
        protected override string EntityLogicalName => "lead";

        public LeadImporter(ServiceClient serviceClient, string baseDir, string excelFileName, int fieldNamesRow)
            : base(serviceClient, baseDir, excelFileName, fieldNamesRow)
        {
        }
    }
}
