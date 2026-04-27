using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace FiscalM_AImport.Importers
{
    public abstract class BaseEntityImporter : IEntityImporter
    {
        private const string GeneratedIdFieldName = "GeneratedId";

        protected readonly ServiceClient _serviceClient;
        protected readonly string _baseDir;
        protected readonly string _excelFileName;

        protected abstract string EntityLogicalName { get; }

        private Dictionary<string, AttributeMetadata>? _attributeMetadata;

        protected BaseEntityImporter(ServiceClient serviceClient, string baseDir, string excelFileName)
        {
            _serviceClient = serviceClient;
            _baseDir = baseDir;
            _excelFileName = excelFileName;
        }

        public void Import()
        {
            var filePath = Path.Combine(_baseDir, _excelFileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"[{EntityLogicalName}] Excel file not found: {filePath}");
                return;
            }

            Console.WriteLine($"[{EntityLogicalName}] Loading entity metadata...");
            var metadata = LoadAttributeMetadata();
            Console.WriteLine($"[{EntityLogicalName}] Loaded {metadata.Count} attribute definitions.");

            Console.WriteLine($"[{EntityLogicalName}] Starting import from '{_excelFileName}'...");

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.Worksheet(1);

            int lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            if (lastCol == 0 || lastRow < 2)
            {
                Console.WriteLine($"[{EntityLogicalName}] Sheet must have at least 2 rows (row 1: display names, row 2: logical field names).");
                return;
            }

            // Build column → logical field name mapping from row 2
            var columnMap = new Dictionary<int, string>(lastCol);
            int generatedIdCol = 0;

            for (int col = 1; col <= lastCol; col++)
            {
                var fieldName = worksheet.Cell(2, col).GetString().Trim();
                if (string.IsNullOrWhiteSpace(fieldName))
                    continue;

                columnMap[col] = fieldName;

                if (string.Equals(fieldName, GeneratedIdFieldName, StringComparison.OrdinalIgnoreCase))
                    generatedIdCol = col;
            }

            // Add GeneratedId column if not present
            if (generatedIdCol == 0)
            {
                generatedIdCol = lastCol + 1;
                worksheet.Cell(1, generatedIdCol).Value = "Generated ID";
                worksheet.Cell(2, generatedIdCol).Value = GeneratedIdFieldName;
                columnMap[generatedIdCol] = GeneratedIdFieldName;
                lastCol = generatedIdCol;
                workbook.Save();
                Console.WriteLine($"[{EntityLogicalName}] Added '{GeneratedIdFieldName}' column at position {generatedIdCol}.");
            }

            // Refresh lastRow in case header additions changed it
            lastRow = worksheet.LastRowUsed()?.RowNumber() ?? lastRow;

            int imported = 0, skipped = 0, errors = 0;

            for (int row = 3; row <= lastRow; row++)
            {
                // Skip rows already imported
                var existingId = worksheet.Cell(row, generatedIdCol).GetString();
                if (!string.IsNullOrWhiteSpace(existingId))
                {
                    skipped++;
                    continue;
                }

                try
                {
                    var entity = new Entity(EntityLogicalName);
                    bool hasAnyField = false;

                    foreach (var kvp in columnMap)
                    {
                        int col = kvp.Key;
                        string fieldName = kvp.Value;

                        if (string.Equals(fieldName, GeneratedIdFieldName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var cellValue = worksheet.Cell(row, col).Value;
                        if (cellValue.IsBlank)
                            continue;

                        if (!metadata.TryGetValue(fieldName, out var attrMeta))
                        {
                            Console.WriteLine($"[{EntityLogicalName}] Row {row}: Field '{fieldName}' not found in entity metadata — column skipped.");
                            continue;
                        }

                        var typedValue = ConvertToDataverseValue(cellValue, attrMeta, row);
                        if (typedValue != null)
                        {
                            entity[fieldName] = typedValue;
                            hasAnyField = true;
                        }
                    }

                    if (!hasAnyField)
                    {
                        skipped++;
                        continue;
                    }

                    var createRequest = new CreateRequest { Target = entity };
                    // Bypass plugins, classic workflows, and Power Automate flows
                    createRequest.Parameters["BypassCustomPluginExecution"] = true;
                    createRequest.Parameters["SuppressCallbackRegistrationExpanderJob"] = true;

                    var createResponse = (CreateResponse)_serviceClient.Execute(createRequest);
                    var recordId = createResponse.id;

                    imported++;
                    Console.WriteLine($"[{EntityLogicalName}] Row {row}: Imported. ID = {recordId}");

                    worksheet.Cell(row, generatedIdCol).Value = recordId.ToString();
                    try
                    {
                        workbook.Save();
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"[{EntityLogicalName}] Row {row}: WARNING — Record created (ID: {recordId}) but Excel could not be saved: {saveEx.Message}. Row may be re-imported on next run.");
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    Console.WriteLine($"[{EntityLogicalName}] Row {row}: ERROR — {ex.Message}");
                }
            }

            Console.WriteLine($"[{EntityLogicalName}] Done. Imported: {imported} | Skipped: {skipped} | Errors: {errors}");
        }

        private Dictionary<string, AttributeMetadata> LoadAttributeMetadata()
        {
            if (_attributeMetadata != null)
                return _attributeMetadata;

            var request = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = EntityLogicalName
            };

            var response = (RetrieveEntityResponse)_serviceClient.Execute(request);

            _attributeMetadata = new Dictionary<string, AttributeMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (var attr in response.EntityMetadata.Attributes)
                _attributeMetadata[attr.LogicalName] = attr;

            return _attributeMetadata;
        }

        private object? ConvertToDataverseValue(XLCellValue cellValue, AttributeMetadata attrMeta, int rowNum)
        {
            if (cellValue.IsBlank) return null;

            var stringRep = GetCellAsString(cellValue);
            if (string.IsNullOrWhiteSpace(stringRep)) return null;

            try
            {
                if (!attrMeta.AttributeType.HasValue)
                    return stringRep;

                switch (attrMeta.AttributeType.Value)
                {
                    case AttributeTypeCode.String:
                    case AttributeTypeCode.Memo:
                    case AttributeTypeCode.EntityName:
                        return stringRep;

                    case AttributeTypeCode.Integer:
                        return cellValue.Type == XLDataType.Number
                            ? (int)Math.Round(cellValue.GetNumber())
                            : int.Parse(stringRep, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.BigInt:
                        return cellValue.Type == XLDataType.Number
                            ? (long)Math.Round(cellValue.GetNumber())
                            : long.Parse(stringRep, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Double:
                        return cellValue.Type == XLDataType.Number
                            ? cellValue.GetNumber()
                            : double.Parse(stringRep, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Decimal:
                        return cellValue.Type == XLDataType.Number
                            ? (decimal)cellValue.GetNumber()
                            : decimal.Parse(stringRep, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Money:
                        var moneyAmt = cellValue.Type == XLDataType.Number
                            ? (decimal)cellValue.GetNumber()
                            : decimal.Parse(stringRep, CultureInfo.InvariantCulture);
                        return new Money(moneyAmt);

                    case AttributeTypeCode.Boolean:
                        if (cellValue.Type == XLDataType.Boolean)
                            return cellValue.GetBoolean();
                        if (int.TryParse(stringRep, out var boolInt))
                            return boolInt != 0;
                        return bool.Parse(stringRep);

                    case AttributeTypeCode.DateTime:
                        if (cellValue.Type == XLDataType.DateTime)
                            return cellValue.GetDateTime();
                        // Excel sometimes stores dates as OADate numbers
                        if (cellValue.Type == XLDataType.Number)
                            return DateTime.FromOADate(cellValue.GetNumber());
                        return DateTime.Parse(stringRep, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Picklist:
                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Status:
                        var optionVal = cellValue.Type == XLDataType.Number
                            ? (int)Math.Round(cellValue.GetNumber())
                            : int.Parse(stringRep, CultureInfo.InvariantCulture);
                        return new OptionSetValue(optionVal);

                    case AttributeTypeCode.Lookup:
                    case AttributeTypeCode.Customer:
                    case AttributeTypeCode.Owner:
                        if (!Guid.TryParse(stringRep, out var lookupGuid))
                            return null;
                        if (attrMeta is LookupAttributeMetadata lookupMeta && lookupMeta.Targets?.Length > 0)
                            return new EntityReference(lookupMeta.Targets[0], lookupGuid);
                        return null;

                    case AttributeTypeCode.Uniqueidentifier:
                        return Guid.TryParse(stringRep, out var guidVal) ? guidVal : (object?)null;

                    default:
                        return stringRep;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{EntityLogicalName}] Row {rowNum}: Warning — could not convert '{stringRep}' for field '{attrMeta.LogicalName}' (type: {attrMeta.AttributeType}): {ex.Message}");
                return null;
            }
        }

        private static string GetCellAsString(XLCellValue cellValue)
        {
            return cellValue.Type switch
            {
                XLDataType.Text     => cellValue.GetText(),
                XLDataType.Number   => cellValue.GetNumber().ToString(CultureInfo.InvariantCulture),
                XLDataType.Boolean  => cellValue.GetBoolean().ToString(),
                XLDataType.DateTime => cellValue.GetDateTime().ToString("o", CultureInfo.InvariantCulture),
                XLDataType.TimeSpan => cellValue.GetTimeSpan().ToString(),
                _                   => string.Empty
            };
        }
    }
}
