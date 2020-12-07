using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UtnCreateUpdateMaxnettoScript
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                LoadExcelFile(openFileDialog1.FileName);
            }
        }

        private void LoadExcelFile(string fileName)
        {
            Cursor = Cursors.WaitCursor;
            dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            try
            {
                var wb = xl.Workbooks.Open(fileName, 0, true);
                try
                {
                    var sheet = wb.Sheets[1];
                    var xlCellTypeLastCell = 11;
                    var lastCell = sheet.Cells.SpecialCells(xlCellTypeLastCell);
                    var colIndexName = IndexToAbc(lastCell.Column);
                    var arrData = (object[,])sheet.Range[$"A1:{colIndexName}{lastCell.Row}"].Value;
                    var rowCount = arrData.GetLength(0);
                    var columnCount = arrData.GetLength(1);
                    using (var gr = lvTable.CreateGraphics())
                    {
                        var font = lvTable.Font;
                        lvTable.Columns.Clear();
                        float width = 0f;
                        var dict = new Dictionary<string, float>();
                        var names = new List<string> { "(stub)" };
                        for (var col = 1; col <= columnCount; col++)
                        {
                            var colName = $"{arrData[2, col]}";
                            names.Add(colName);
                            width = gr.MeasureString(colName, font).Width;
                            dict.Add(colName, width);
                            lvTable.Columns.Add(colName);
                        }
                        lvTable.Items.Clear();
                        for (var row = 3; row <= rowCount; row++)
                        {
                            var rowValue = $"{arrData[row, 1]}";
                            if (DateTime.TryParse(rowValue, out DateTime dt))
                                rowValue = $"{dt:dd.MM.yy HH:mm:ss}";
                            width = gr.MeasureString(rowValue, font).Width;
                            if (width > dict[names[1]]) dict[names[1]] = width;
                            var lvi = new ListViewItem(rowValue);
                            for (var col = 2; col <= columnCount; col++)
                            {
                                rowValue = $"{arrData[row, col]}";
                                width = gr.MeasureString(rowValue, font).Width + 10f;
                                if (width > dict[names[col]]) dict[names[col]] = width;
                                lvi.SubItems.Add(rowValue);
                            }
                            lvTable.Items.Add(lvi);
                        }
                        for (var col = 0; col < lvTable.Columns.Count; col++)
                        {
                            var column = lvTable.Columns[col];
                            column.Width = Convert.ToInt32(dict[names[col + 1]]);
                        }
                    }
                }
                finally
                {
                    wb.Close(false);
                }
            }
            finally
            {
                xl.Quit();
                Cursor = Cursors.Default;
            }
        }

        private static string IndexToAbc(int index)
        {
            const string s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var letters = s.ToCharArray();
            var result = "";
            if (index <= letters.Length)
            {
                result = $"{letters[index - 1]}";
                return result;
            }
            while (true)
            {
                var big = (index - 1) % 26;
                result = $"{letters[big]}{result}";
                index = (index - 1) / 26;
                if (index <= letters.Length)
                    return result;
            }
        }
    }
}
