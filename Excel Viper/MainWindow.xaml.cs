using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Viper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private delegate void UpdateProgressBarDelegate(
        System.Windows.DependencyProperty dp, Object value);


        public MainWindow()
        {
            InitializeComponent();

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel (.xls, .xlsx)|*.xls;*.xlsx";
            dlg.ShowDialog();

            filePath.Text = dlg.FileName;
            Load(null,null);
        }

        private void Start_Format(object sender, RoutedEventArgs e)
        {
            prog.Minimum = 0;
            prog.Value = 0;
            prog.Maximum = Convert.ToInt32(lastRow.Text);

            UpdateProgressBarDelegate updatePbDelegate =
            new UpdateProgressBarDelegate(prog.SetValue);


            // Open File
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath.Text);

            // Get old Worksheet
            Excel.Worksheet oldWorksheet = workbook.Worksheets[worksheetName.Text];

            // Create new Worksheet
            Excel.Worksheet newWorksheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            if (!sheetExist(newWorksheetName.Text, ref workbook))
                newWorksheet.Name = newWorksheetName.Text;

            if (Convert.ToInt32(startRow.Text) > 1)
            {
                int max = Convert.ToInt32(startFields.Text);

                for (int col = 1; col <= max; col++)
                {
                    newWorksheet.Cells[1, col] = oldWorksheet.Cells[1, col];
                }
            }

            int actRow = Convert.ToInt32(startRow.Text);
            // Every Row
            for (int row = Convert.ToInt32(startRow.Text); row <= Convert.ToInt32(lastRow.Text); row++)
            {
                // Write Startfields
                actRow = addProdFields(row, ref oldWorksheet, ref newWorksheet, actRow);
                //prog.Value = row;
                Dispatcher.Invoke(updatePbDelegate,System.Windows.Threading.DispatcherPriority.Background,new object[] { ProgressBar.ValueProperty, (double)row});
            }

            // Save & Close
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            MessageBox.Show("Formatierung abgeschlossen!", "Fertig");
        }

        void addStartFields(int newRow, int oldRow, ref Excel.Worksheet oldWorksheet, ref Excel.Worksheet newWorksheet)
        {
            for (int col = 1; col <= Convert.ToInt32(startFields.Text); col++)
            {
                newWorksheet.Cells[newRow, col] = oldWorksheet.Cells[oldRow, col];
            }
        }

        int addProdFields(int row, ref Excel.Worksheet oldWorksheet, ref Excel.Worksheet newWorksheet, int actRow = -1)
        {
            int max = Convert.ToInt32(startFields.Text) + (Convert.ToInt32(filedsPerProduct.Text) * Convert.ToInt32(prodCount.Text));
            if (actRow == -1)
                actRow = row;
            int counter = 0;

            for (int col = Convert.ToInt32(startFields.Text) + 1; col <= max; col += Convert.ToInt32(filedsPerProduct.Text))
            {
                int actCol =  Convert.ToInt32(startFields.Text) + 1;

                if (!isClear(row, col, ref oldWorksheet))
                {
                    var oT = oldWorksheet.Cells[Convert.ToInt32(startRow.Text) - 1, col].Text;
                    newWorksheet.Cells[actRow, actCol] = oT + oldWorksheet.Cells[row, col].Text;
                    newWorksheet.Cells[actRow, actCol + 1] = oldWorksheet.Cells[row, col + 1];
                    newWorksheet.Cells[actRow, actCol + 2] = oldWorksheet.Cells[row, col + 2];

                    // Write Startfields
                    addStartFields(actRow, row, ref oldWorksheet, ref newWorksheet);

                    actRow++;
                }
                counter++;
            }

            return actRow;
        }

        bool isClear(int row, int col, ref Excel.Worksheet oldWorksheet)
        {
            return String.IsNullOrWhiteSpace(oldWorksheet.Cells[row, col].Text) && String.IsNullOrWhiteSpace(oldWorksheet.Cells[row, col + 1].Text) && String.IsNullOrWhiteSpace(oldWorksheet.Cells[row, col + 2].Text);
        }

        private void Load(object sender, RoutedEventArgs e)
        {
            try
            {
                // Open File
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath.Text);

                worksheetName.Items.Clear();

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    worksheetName.Items.Add(sheet.Name);

                workbook.Close();
                excelApp.Quit();
                MessageBox.Show("Laden abgeschlossen!", "Fertig");
            }
            catch { 
            }

        }

        private void worksheetName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                newWorksheetName.Text = e.AddedItems[0].ToString() + " Viper";
            }
            catch { }

            try
            {
                // Open File
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath.Text);

                Excel.Worksheet sheet = workbook.Worksheets[e.AddedItems[0].ToString()];
                
                bool isEmpty = false;
                int counter = 1;

                while (!isEmpty)
                {
                    isEmpty = String.IsNullOrWhiteSpace(sheet.Cells[counter, 1].Text);
                    counter++;
                }

                lastRow.Text = (counter - 2).ToString();

                workbook.Close();
                excelApp.Quit();
            }
            catch { }
        }

        bool sheetExist(string name, ref Excel.Workbook book)
        {
            bool found = false;

            foreach (Excel.Worksheet sheet in book.Worksheets)
                if (sheet.Name == name)
                    found = true;

            return found;
        }

    }
}
