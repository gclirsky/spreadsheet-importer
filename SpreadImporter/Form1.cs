using SpreadImporter.Mapper;
using System;
using System.Data;
using System.Windows.Forms;

namespace SpreadImporter
{
    public partial class Form1 : Form
    {
        private DbMapper mapper;
        private Importer importer;

        const string sheetPath = "sample.xlsx";
        const string importTable = "staff";

        public Form1()
        {
            InitializeComponent();
            InitializeDbMapper();
            InitializeImporter();
        }

        private void InitializeImporter()
        {
            importer = new Importer(mapper, sheetPath);
            importer.setImportTableName(importTable);
        }

        private void InitializeDbMapper()
        {
            mapper = new DbMapper();
            mapper.init();
            mapper.loadLookupsFromDb();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable table = importer.getExcelPackage().getExcelSheet().readCells();

                using (var connection = mapper.openDbConnection())
                {
                    var executionCount = mapper.insert(connection, table);
                    MessageBox.Show(string.Format("{0} rows affected by insert", executionCount));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
