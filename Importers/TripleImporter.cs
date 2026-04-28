using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace FiscalM_AImport.Importers
{
    public class TripleImporter
    {
        private const string LeadEntity    = "lead";
        private const string ContactEntity = "contact";
        private const string AccountEntity = "account";

        private const string ColGeneratedLeadId    = "GeneratedLeadId";
        private const string ColGeneratedContactId = "GeneratedContactId";
        private const string ColGeneratedAccountId = "GeneratedAccountId";

        private readonly ServiceClient _serviceClient;
        private readonly string _baseDir;
        private readonly string _excelFileName;
        private readonly int _fieldNamesRow;

        public TripleImporter(ServiceClient serviceClient, string baseDir, string excelFileName, int fieldNamesRow)
        {
            _serviceClient = serviceClient;
            _baseDir       = baseDir;
            _excelFileName = excelFileName;
            _fieldNamesRow = fieldNamesRow;
        }

        public void Import()
        {
            var filePath = Path.Combine(_baseDir, _excelFileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Excel file not found: {filePath}");
                return;
            }

            Console.WriteLine("Loading entity metadata...");
            var leadMeta    = LoadMetadata(LeadEntity);
            var contactMeta = LoadMetadata(ContactEntity);
            var accountMeta = LoadMetadata(AccountEntity);
            Console.WriteLine();

            Console.WriteLine($"Opening '{_excelFileName}'...");
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.Worksheet(1);

            int lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            if (lastCol == 0 || lastRow < _fieldNamesRow)
            {
                Console.WriteLine($"Sheet must have at least {_fieldNamesRow} row(s).");
                return;
            }

            // ── Parse column definitions from the field-names row ─────────────────────
            var columns = new List<ColumnDef>();
            for (int col = 1; col <= lastCol; col++)
            {
                var rawName = worksheet.Cell(_fieldNamesRow, col).GetString().Trim();
                if (!string.IsNullOrWhiteSpace(rawName))
                    columns.Add(ParseColumn(col, rawName));
            }

            int generatedLeadIdCol    = columns.FirstOrDefault(c => c.Role == ColumnRole.GeneratedLeadId)?.Index    ?? 0;
            int generatedContactIdCol = columns.FirstOrDefault(c => c.Role == ColumnRole.GeneratedContactId)?.Index ?? 0;
            int generatedAccountIdCol = columns.FirstOrDefault(c => c.Role == ColumnRole.GeneratedAccountId)?.Index ?? 0;
            int secondContactCol      = columns.FirstOrDefault(c => c.Role == ColumnRole.SecondContact)?.Index      ?? 0;

            // If "Second Contact" column wasn't found via logical name, scan the display-names row
            if (secondContactCol == 0 && _fieldNamesRow > 1)
            {
                for (int col = 1; col <= lastCol; col++)
                {
                    var displayName = worksheet.Cell(_fieldNamesRow - 1, col).GetString().Trim();
                    if (string.Equals(displayName, "Second Contact", StringComparison.OrdinalIgnoreCase))
                    {
                        secondContactCol = col;
                        break;
                    }
                }
            }

            // ── Add missing tracking columns ──────────────────────────────────────────
            int displayRow = _fieldNamesRow > 1 ? _fieldNamesRow - 1 : 0;

            if (generatedLeadIdCol == 0)
            {
                generatedLeadIdCol = ++lastCol;
                if (displayRow > 0) worksheet.Cell(displayRow, generatedLeadIdCol).Value = "Generated Lead ID";
                worksheet.Cell(_fieldNamesRow, generatedLeadIdCol).Value = ColGeneratedLeadId;
                columns.Add(new ColumnDef(generatedLeadIdCol, ColGeneratedLeadId, null, ColGeneratedLeadId, ColumnRole.GeneratedLeadId));
                Console.WriteLine($"Added '{ColGeneratedLeadId}' column at position {generatedLeadIdCol}.");
            }

            if (generatedContactIdCol == 0)
            {
                generatedContactIdCol = ++lastCol;
                if (displayRow > 0) worksheet.Cell(displayRow, generatedContactIdCol).Value = "Generated Contact ID";
                worksheet.Cell(_fieldNamesRow, generatedContactIdCol).Value = ColGeneratedContactId;
                columns.Add(new ColumnDef(generatedContactIdCol, ColGeneratedContactId, null, ColGeneratedContactId, ColumnRole.GeneratedContactId));
                Console.WriteLine($"Added '{ColGeneratedContactId}' column at position {generatedContactIdCol}.");
            }

            if (generatedAccountIdCol == 0)
            {
                generatedAccountIdCol = ++lastCol;
                if (displayRow > 0) worksheet.Cell(displayRow, generatedAccountIdCol).Value = "Generated Account ID";
                worksheet.Cell(_fieldNamesRow, generatedAccountIdCol).Value = ColGeneratedAccountId;
                columns.Add(new ColumnDef(generatedAccountIdCol, ColGeneratedAccountId, null, ColGeneratedAccountId, ColumnRole.GeneratedAccountId));
                Console.WriteLine($"Added '{ColGeneratedAccountId}' column at position {generatedAccountIdCol}.");
            }

            workbook.Save();
            lastRow = worksheet.LastRowUsed()?.RowNumber() ?? lastRow;

            int processed = 0, skipped = 0, errors = 0;
            Console.WriteLine();

            for (int row = _fieldNamesRow + 1; row <= lastRow; row++)
            {
                var existingLeadId    = worksheet.Cell(row, generatedLeadIdCol).GetString();
                var existingContactId = worksheet.Cell(row, generatedContactIdCol).GetString();
                var existingAccountId = worksheet.Cell(row, generatedAccountIdCol).GetString();

                // Row fully processed — nothing to do
                if (!string.IsNullOrWhiteSpace(existingLeadId) &&
                    !string.IsNullOrWhiteSpace(existingContactId) &&
                    !string.IsNullOrWhiteSpace(existingAccountId))
                {
                    skipped++;
                    continue;
                }

                bool isSecondContact = secondContactCol > 0 &&
                    string.Equals(
                        worksheet.Cell(row, secondContactCol).GetString().Trim(),
                        "Yes", StringComparison.OrdinalIgnoreCase);

                bool leadFailed = false, contactFailed = false, accountFailed = false;
                Guid leadId = Guid.Empty;

                // ── 1. LEAD ───────────────────────────────────────────────────────────
                if (!string.IsNullOrWhiteSpace(existingLeadId))
                {
                    Guid.TryParse(existingLeadId, out leadId);
                }
                else
                {
                    try
                    {
                        var entity = BuildEntity(LeadEntity, "Lead", leadMeta, columns, row, worksheet, secondContactCol);

                        if (entity.Attributes.Count == 0)
                        {
                            skipped++;
                            continue;
                        }

                        var resp = (CreateResponse)_serviceClient.Execute(BypassRequest(entity));
                        leadId = resp.id;

                        worksheet.Cell(row, generatedLeadIdCol).Value = leadId.ToString();
                        TrySave(workbook, row, "Lead");
                        Console.WriteLine($"  Row {row} Lead    created: {leadId}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Row {row} Lead    ERROR: {ex.Message}");
                        leadFailed = true;
                    }
                }

                // Cannot proceed without a valid Lead ID
                if (leadFailed) { errors++; continue; }

                // ── 2. CONTACT ────────────────────────────────────────────────────────
                if (isSecondContact)
                {
                    if (string.IsNullOrWhiteSpace(existingContactId))
                    {
                        // Mark as intentionally skipped so the row is not re-evaluated each run
                        worksheet.Cell(row, generatedContactIdCol).Value = "N/A";
                        TrySave(workbook, row, "Contact-N/A");
                    }
                    Console.WriteLine($"  Row {row} Contact skipped (Second Contact = Yes).");
                }
                else if (!string.IsNullOrWhiteSpace(existingContactId))
                {
                    // Already imported in a previous run — nothing to do
                }
                else
                {
                    try
                    {
                        var entity = BuildEntity(ContactEntity, "Contact", contactMeta, columns, row, worksheet, secondContactCol);
                        entity["chl_leadori"] = new EntityReference(LeadEntity, leadId);

                        var resp = (CreateResponse)_serviceClient.Execute(BypassRequest(entity));
                        var contactId = resp.id;

                        worksheet.Cell(row, generatedContactIdCol).Value = contactId.ToString();
                        TrySave(workbook, row, "Contact");
                        Console.WriteLine($"  Row {row} Contact created: {contactId}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Row {row} Contact ERROR: {ex.Message}");
                        contactFailed = true;
                    }
                }

                // ── 3. ACCOUNT ────────────────────────────────────────────────────────
                // Account links to Lead, not to Contact — proceed even if Contact failed
                if (!string.IsNullOrWhiteSpace(existingAccountId))
                {
                    // Already imported in a previous run — nothing to do
                }
                else
                {
                    try
                    {
                        var entity = BuildEntity(AccountEntity, "Account", accountMeta, columns, row, worksheet, secondContactCol);
                        entity["originatingleadid"] = new EntityReference(LeadEntity, leadId);

                        var resp = (CreateResponse)_serviceClient.Execute(BypassRequest(entity));
                        var accountId = resp.id;

                        worksheet.Cell(row, generatedAccountIdCol).Value = accountId.ToString();
                        TrySave(workbook, row, "Account");
                        Console.WriteLine($"  Row {row} Account created: {accountId}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Row {row} Account ERROR: {ex.Message}");
                        accountFailed = true;
                    }
                }

                if (contactFailed || accountFailed)
                    errors++;
                else
                    processed++;
            }

            Console.WriteLine();
            Console.WriteLine($"Done. Rows completed: {processed} | Skipped (already done): {skipped} | Rows with errors: {errors}");
        }

        // ── Entity builder ────────────────────────────────────────────────────────────

        private static Entity BuildEntity(
            string entityName, string entityPrefix,
            Dictionary<string, AttributeMetadata> metadata,
            List<ColumnDef> columns, int row,
            IXLWorksheet worksheet, int secondContactCol)
        {
            var entity = new Entity(entityName);

            foreach (var col in columns)
            {
                if (col.Role != ColumnRole.Normal) continue;
                if (col.Index == secondContactCol)  continue;

                // Include shared columns (no prefix) or those explicitly for this entity
                bool include = col.EntityPrefix == null ||
                    string.Equals(col.EntityPrefix, entityPrefix, StringComparison.OrdinalIgnoreCase);
                if (!include) continue;

                var cell = worksheet.Cell(row, col.Index).Value;
                if (cell.IsBlank) continue;

                if (!metadata.TryGetValue(col.FieldName, out var attrMeta))
                {
                    Console.WriteLine($"  [{entityPrefix}] Row {row}: '{col.FieldName}' not in metadata — skipped.");
                    continue;
                }

                var typed = ConvertValue(cell, attrMeta, row, entityPrefix);
                if (typed != null)
                    entity[col.FieldName] = typed;
            }

            return entity;
        }

        // ── Helpers ───────────────────────────────────────────────────────────────────

        private static CreateRequest BypassRequest(Entity entity)
        {
            var req = new CreateRequest { Target = entity };
            req.Parameters["BypassCustomPluginExecution"]          = true;
            req.Parameters["SuppressCallbackRegistrationExpanderJob"] = true;
            return req;
        }

        private static void TrySave(XLWorkbook workbook, int row, string context)
        {
            try { workbook.Save(); }
            catch (Exception ex)
            {
                Console.WriteLine($"  Row {row} WARNING [{context}]: record created but Excel save failed — {ex.Message}");
            }
        }

        private Dictionary<string, AttributeMetadata> LoadMetadata(string entityName)
        {
            Console.WriteLine($"  Loading metadata for '{entityName}'...");
            var req  = new RetrieveEntityRequest { EntityFilters = EntityFilters.Attributes, LogicalName = entityName };
            var resp = (RetrieveEntityResponse)_serviceClient.Execute(req);

            var dict = new Dictionary<string, AttributeMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (var attr in resp.EntityMetadata.Attributes)
                dict[attr.LogicalName] = attr;

            Console.WriteLine($"  Loaded {dict.Count} attributes.");
            return dict;
        }

        private static ColumnDef ParseColumn(int index, string rawName)
        {
            if (string.Equals(rawName, ColGeneratedLeadId,    StringComparison.OrdinalIgnoreCase)) return new ColumnDef(index, rawName, null, rawName, ColumnRole.GeneratedLeadId);
            if (string.Equals(rawName, ColGeneratedContactId, StringComparison.OrdinalIgnoreCase)) return new ColumnDef(index, rawName, null, rawName, ColumnRole.GeneratedContactId);
            if (string.Equals(rawName, ColGeneratedAccountId, StringComparison.OrdinalIgnoreCase)) return new ColumnDef(index, rawName, null, rawName, ColumnRole.GeneratedAccountId);

            if (string.Equals(rawName, "secondcontact",  StringComparison.OrdinalIgnoreCase) ||
                string.Equals(rawName, "second_contact", StringComparison.OrdinalIgnoreCase))
                return new ColumnDef(index, rawName, null, rawName, ColumnRole.SecondContact);

            var dot = rawName.IndexOf('.');
            if (dot > 0 && dot < rawName.Length - 1)
            {
                var prefix = rawName.Substring(0, dot);
                var field  = rawName.Substring(dot + 1);
                return new ColumnDef(index, rawName, prefix, field, ColumnRole.Normal);
            }

            return new ColumnDef(index, rawName, null, rawName, ColumnRole.Normal);
        }

        // ── Value converter ───────────────────────────────────────────────────────────

        private static object? ConvertValue(XLCellValue cell, AttributeMetadata attrMeta, int row, string entityName)
        {
            if (cell.IsBlank) return null;

            var str = CellToString(cell);
            if (string.IsNullOrWhiteSpace(str)) return null;

            try
            {
                if (!attrMeta.AttributeType.HasValue) return str;

                switch (attrMeta.AttributeType.Value)
                {
                    case AttributeTypeCode.String:
                    case AttributeTypeCode.Memo:
                    case AttributeTypeCode.EntityName:
                        return str;

                    case AttributeTypeCode.Integer:
                        return cell.Type == XLDataType.Number
                            ? (int)Math.Round(cell.GetNumber())
                            : int.Parse(str, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.BigInt:
                        return cell.Type == XLDataType.Number
                            ? (long)Math.Round(cell.GetNumber())
                            : long.Parse(str, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Double:
                        return cell.Type == XLDataType.Number
                            ? cell.GetNumber()
                            : double.Parse(str, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Decimal:
                        return cell.Type == XLDataType.Number
                            ? (decimal)cell.GetNumber()
                            : decimal.Parse(str, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Money:
                        var moneyAmt = cell.Type == XLDataType.Number
                            ? (decimal)cell.GetNumber()
                            : decimal.Parse(str, CultureInfo.InvariantCulture);
                        return new Money(moneyAmt);

                    case AttributeTypeCode.Boolean:
                        if (cell.Type == XLDataType.Boolean) return cell.GetBoolean();
                        if (int.TryParse(str, out var boolInt)) return boolInt != 0;
                        return bool.Parse(str);

                    case AttributeTypeCode.DateTime:
                        if (cell.Type == XLDataType.DateTime) return cell.GetDateTime();
                        if (cell.Type == XLDataType.Number)   return DateTime.FromOADate(cell.GetNumber());
                        return DateTime.Parse(str, CultureInfo.InvariantCulture);

                    case AttributeTypeCode.Picklist:
                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Status:
                        var optVal = cell.Type == XLDataType.Number
                            ? (int)Math.Round(cell.GetNumber())
                            : int.Parse(str, CultureInfo.InvariantCulture);
                        return new OptionSetValue(optVal);

                    case AttributeTypeCode.Lookup:
                    case AttributeTypeCode.Customer:
                    case AttributeTypeCode.Owner:
                        if (!Guid.TryParse(str, out var lookupGuid)) return null;
                        if (attrMeta is LookupAttributeMetadata lm && lm.Targets?.Length > 0)
                            return new EntityReference(lm.Targets[0], lookupGuid);
                        return null;

                    case AttributeTypeCode.Uniqueidentifier:
                        return Guid.TryParse(str, out var guidVal) ? guidVal : (object?)null;

                    default:
                        return str;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  [{entityName}] Row {row}: Warning — cannot convert '{str}' for '{attrMeta.LogicalName}' ({attrMeta.AttributeType}): {ex.Message}");
                return null;
            }
        }

        private static string CellToString(XLCellValue v) => v.Type switch
        {
            XLDataType.Text     => v.GetText(),
            XLDataType.Number   => v.GetNumber().ToString(CultureInfo.InvariantCulture),
            XLDataType.Boolean  => v.GetBoolean().ToString(),
            XLDataType.DateTime => v.GetDateTime().ToString("o", CultureInfo.InvariantCulture),
            XLDataType.TimeSpan => v.GetTimeSpan().ToString(),
            _                   => string.Empty
        };

        // ── Column model ──────────────────────────────────────────────────────────────

        private enum ColumnRole
        {
            Normal,
            GeneratedLeadId,
            GeneratedContactId,
            GeneratedAccountId,
            SecondContact
        }

        private class ColumnDef
        {
            public int        Index        { get; }
            public string     RawName      { get; }
            public string?    EntityPrefix { get; }   // null = shared across all entities
            public string     FieldName    { get; }   // logical name without the entity prefix
            public ColumnRole Role         { get; }

            public ColumnDef(int index, string rawName, string? entityPrefix, string fieldName, ColumnRole role)
            {
                Index        = index;
                RawName      = rawName;
                EntityPrefix = entityPrefix;
                FieldName    = fieldName;
                Role         = role;
            }
        }
    }
}
