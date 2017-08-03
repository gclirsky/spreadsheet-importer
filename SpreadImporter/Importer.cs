using OfficeOpenXml;
using SpreadImporter.Mapper;
using System;
using System.Data;
using System.IO;

namespace SpreadImporter
{
    public class Importer
    {
        private static ExcelPackage excelPackage;
        private readonly DbMapper mapper;
        private readonly FileInfo spreadsheet;
        private string importTableName;

        private int bodyStartRow, bodyEndRow, headerRow;
        private int bodyStartCol, bodyEndCol, headerStartCol, headerEndCol;

        public ExcelWorksheet WorkSheet { get; private set; }

        public Importer(DbMapper mapper, string sheetPath)
        {
            if (mapper == null)
            {
                throw new ArgumentNullException(nameof(mapper));
            }

            if (!File.Exists(sheetPath))
            {
                throw new FileNotFoundException("Spreadsheet not found");
            }

            this.mapper = mapper;
            this.spreadsheet = new FileInfo(sheetPath);
        }

        public Importer getExcelPackage()
        {
            if (excelPackage == null)
            {
                excelPackage = new ExcelPackage(spreadsheet);
            }

            return this;
        }

        public void setImportTableName(string name)
        {
            this.importTableName = name;
        }

        public Importer getExcelSheet(string sheetName = "")
        {
            if (excelPackage == null)
            {
                throw new NullReferenceException("SpreadSheet Package not yet been loaded, need to get Excel Package firstly");
            }

            var workbook = excelPackage.Workbook;

            try
            {
                if (string.IsNullOrEmpty(sheetName))
                {
                    WorkSheet = workbook.Worksheets[1];  // 1-based index
                }
                else
                {
                    WorkSheet = workbook.Worksheets[sheetName];
                }

                bodyStartRow = mapper.DbInfo.SheetFormat.BodyStartRow;
                bodyEndRow = WorkSheet.Dimension.End.Row;
                bodyStartCol = WorkSheet.Dimension.Start.Column;
                bodyEndCol = WorkSheet.Dimension.End.Column;

                headerRow = mapper.DbInfo.SheetFormat.HeaderRow;
                headerStartCol = bodyStartCol;
                headerEndCol = bodyEndCol;
            }
            catch (IndexOutOfRangeException ex)
            {
                throw new InvalidOperationException("Invalid worksheet name", ex);
            }

            return this;
        }


        public DataTable readCells()
        {
            if (WorkSheet == null)
            {
                throw new NullReferenceException("Excel worksheet not yet been loaded, need to get Excel Sheet with expected sheet name or sheet index firstly");
            }

            var dataTable = createDbTableSchema(WorkSheet);
            fillDataTable(WorkSheet, dataTable);

            return dataTable;
        }

        private DataTable createDbTableSchema(ExcelWorksheet worksheet)
        {
            if (string.IsNullOrEmpty(importTableName))
            {
                throw new InvalidOperationException("The import table name not yet been setup");
            }

            var mapperTable = mapper.DbInfo.getDbTable(importTableName);
            if (mapperTable == null)
            {
                throw new InvalidOperationException(string.Format(@"Table '{0}' not yet configured in DbMapper xml", importTableName));
            }

            DataTable table = new DataTable(mapperTable.Name);
            using (ExcelRange rng = worksheet.Cells[headerRow, headerStartCol, headerRow, headerEndCol])
            {
                foreach (var item in rng)
                {
                    table.Columns.Add(item.Value.ToString());
                }
            }

            return table;
        }

        private void fillDataTable(ExcelWorksheet worksheet, DataTable dataTable)
        {
            for (int row = bodyStartRow; row <= bodyEndRow; row++)
            {
                DataRow dtRow = dataTable.NewRow();
                using (ExcelRange rng = worksheet.Cells[row, bodyStartCol, row, bodyEndCol])
                {
                    // TODO: will failed at here as DbMapper.xml not yet fully configured, so cannot the data after the column exceeds 3
                    foreach (var item in rng)
                    {
                        var colIndex = item.Start.Column;
                        var dbField = mapper.DbInfo.getDbTable(importTableName).getDbField(colIndex);
                        dtRow[item.Columns] = dbField.dataFormatting(item.Value);
                    }
                }

                dataTable.Rows.Add(dtRow);
            }
        }
    }
}
