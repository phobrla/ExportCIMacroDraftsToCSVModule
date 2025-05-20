# ExportCIMacroDraftsToCSV

This VBA macro automates the export of data from the "CI Macro Drafts" Excel worksheet to a properly formatted CSV file for use with PhraseExpress and ServiceNow macros.

## Features

- **Exports from Table:** Automatically reads from the `CIMacroDrafts` table in the `CI Macro Drafts` worksheet.
- **Field Formatting:** Produces a CSV with three columns:
  1. Description
  2. Macro Text (`txt`)
  3. Constant label (`_SNOW Macros (SERVICES)`)
- **Macro Text Format:** The `txt` field is constructed to match ServiceNow and PhraseExpress macro requirements, with placeholders like `{#TAB}`, `{#DEL -count 15}`, and `{#ENTER}`.
- **Smart New Line Handling:** All `{#ENTER}` to the left of `{#insert -id ... SNOW Acknowledged}` are kept as `{#ENTER}`; all after are `{#ENTER -variablename New Line}`.
- **Comma and Quote Handling:** Double quotes are escaped as `""`. The `txt` field is wrapped in quotes only if it contains a comma.
- **No Trailing Blank Line:** Output does not end with an extra blank line.
- **Status Update:** Successfully exported rows are marked as "Yes" in the "Exported" column.

## Usage

1. **Open your Excel file** and ensure you have a worksheet named `CI Macro Drafts` with a table named `CIMacroDrafts`.
2. **Add the macro** from [`ExportCIMacroDraftsToCSV.vba`](./ExportCIMacroDraftsToCSV.vba) to your VBA editor.
3. **Run the macro**.  
   - The macro will create a CSV file in your PhraseExpress Documents folder, named with a timestamp.
   - You will receive a message box on completion indicating the file path and number of rows exported.

## Customization

- **Special Characters/Keywords:**  
  - You can adjust the special character placeholders in the `specialChars` dictionary at the top of the script.
- **Output Path:**  
  - Modify the `FilePath` variable in the macro to change where files are saved.

## Example Macro Text

Example output for the `txt` field:

```
MICROSOFT ACTIVE DIRECTORY{#TAB}{#TAB}{#TAB}{#TAB}{#TAB}End user access{#sleep 1000}{#TAB}Account locked{#sleep 1000}{#TAB}{#TAB}{#TAB}{#TAB}{#TAB}{#DEL -count 15}TSG_TSC1{#sleep 1000}{#ENTER}{#TAB}{#TAB}Hobrla, Phil (Phil){#ENTER}{#sleep 1000}{#TAB}{#insert -id 1F4D85EA-7001-48CF-88F7-F9E7012C27FE -variablename SNOW Acknowledged}Locked Account{#TAB}User indicates they are locked out of their Active Directory Account.
```

## Notes

- Ensure all sheet, table, and column names match what the macro expects.
- The macro ensures compliance with CSV standards for double-quote and comma handling.

## License

This script is provided as-is for internal automation and workflow improvement.
