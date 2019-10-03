using OfficeOpenXml;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordExporter.Core.WorkItems;

namespace WordExporter.Core.ExcelManipulation
{
    /// <summary>
    /// Todo: probably it is better to remove completely the part where 
    /// this component is filling dictionaries with parameter and let
    /// this part to be completed outside. (Also fix <see cref="WordManipulator"/>
    /// accordingly.
    /// </summary>
    public class ExcelManipulator : IDisposable
    {
        private readonly ExcelPackage _excelPackage;

        public ExcelManipulator(String excelFileName)
        {
            _excelPackage = new ExcelPackage(new FileInfo(excelFileName));
        }

        public void FillWorkItems(IEnumerable<LinkedWorkItem> workItems)
        {
            var sheet = _excelPackage.Workbook.Worksheets[1];
            Dictionary<Int32, String> fieldMapping = new Dictionary<int, string>();
            Int32 col = 1;
            while (sheet.Cells[2, col].Value != null)
            {
                string trimmedValue = sheet.Cells[2, col].Value.ToString().Trim('{', '}');
                fieldMapping[col] = trimmedValue;
                col++;
            }

            sheet.InsertRow(3, workItems.Count());
            Int32 row = 3;
            foreach (var workItem in workItems)
            {
                Log.Logger.Debug("Filling work item {id} in row {rownum}", workItem.WorkItem.Id, row);
                Dictionary<String, Object> propertyDictionary = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                LinkedWorkItem current = workItem;
                while (current != null)
                {
                    current.WorkItem.FillPropertyDictionary(
                        propertyDictionary, 
                        current.WorkItem.Type.Name.ToLower() + '.',
                        true);
                    current = current.Parent;
                }

                sheet.Cells[2, 1, 2, col].Copy(sheet.Cells[row, 1]);
                //now for each mapping for current row we need to substitute the real value
                foreach (var colInfo in fieldMapping)
                {
                    sheet.Cells[row, colInfo.Key].Value = "";
                    if (propertyDictionary.TryGetValue(colInfo.Value, out var value))
                    {
                        sheet.Cells[row, colInfo.Key].Value = value?.ToString() ?? "";
                    }
                }
                row++;
            }

            sheet.DeleteRow(2);
        }

        public void Dispose()
        {
            _excelPackage.Save();
            _excelPackage.Dispose();
        }
    }
}
