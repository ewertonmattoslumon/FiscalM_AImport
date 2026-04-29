# Fiscal M&A Importer (Lead, Contact and Account)

## Key design decisions:

- Bypass automation — every CreateRequest carries BypassCustomPluginExecution = true (blocks plugins and classic workflows) and SuppressCallbackRegistrationExpanderJob = true (blocks Power Automate). Your user must have the prvBypassCustomPlugins privilege in Dynamics 365.
- Field type handling — on first run, BaseEntityImporter calls RetrieveEntityRequest to download the entity's attribute metadata. Each cell value is then cast to the correct SDK type (OptionSetValue, Money, EntityReference, DateTime, etc.) so Dataverse accepts it without coercion errors. For lookup fields, the Excel cell must contain the target record's GUID.
- Resume support — the GeneratedId column is added automatically (with "Generated ID" as the display-name header and GeneratedId as the logical-name header). On restart, any row that already has a value there is skipped. The workbook is saved after every successful row.
- Excel file location — place your .xlsx files in the project directory and run with dotnet run, or place them next to the compiled exe and run it from that folder.

To get started, update appsettings.json:
1. Set Dataverse.ConnectionString to your org's connection string.
2. Flip the ImportLead / ImportContact / ImportAccount flags on appsettings.json.
3. Set the matching *ExcelFile names on appsettings.json.
