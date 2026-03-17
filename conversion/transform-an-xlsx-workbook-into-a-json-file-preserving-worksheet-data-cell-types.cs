using System;
using Aspose.Cells;

class WorkbookToJson
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Path where the resulting JSON file will be saved
        string jsonPath = "output.json";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve hierarchy and cell types
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,

            // Preserve parent‑child hierarchy (nested structure)
            ExportNestedStructure = true,

            // Include empty cells as null values
            ExportEmptyCells = true,

            // Keep original cell data types (do not force everything to string)
            ExportAsString = false
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonPath, saveOptions);
    }
}