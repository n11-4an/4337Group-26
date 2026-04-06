using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        private DataTable importedData;

        private string connectionString = @"Server=.\SQLEXPRESS;Database=Lab3_4337;Integrated Security=True;";

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Гильманова Азиза");
        }
        

        private void ShowInfo_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Гильманова Азиза\nГруппа: 4337\nВозраст: 18\nВариант 8", "Об авторе");
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.Title = "Выберите файл 4.xlsx";

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    importedData = LoadFromExcel(filePath);
                    SaveToDatabase(importedData);
                    MessageBox.Show($"Импортировано {importedData.Rows.Count} записей!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (importedData == null || importedData.Rows.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте данные!", "Предупреждение");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "export_po_dolzhnostyam.xlsx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    GroupAndExportToExcel(importedData, filePath);
                    MessageBox.Show("Экспорт завершен!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private DataTable LoadFromExcel(string filePath)
        {
            DataTable dt = new DataTable();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    dt.Columns.Add(worksheet.Cells[1, col].Text);

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    dt.Rows.Add(dataRow);
                }
            }
            return dt;
        }

        private void SaveToDatabase(DataTable data)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("DELETE FROM Clients", conn))
                    cmd.ExecuteNonQuery();

                foreach (DataRow row in data.Rows)
                {
                    string sql = @"INSERT INTO Clients (Код_клиента, ФИО, Логин, Должность) 
                                   VALUES (@code, @fio, @login, @position)";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@code", row[0].ToString());
                        cmd.Parameters.AddWithValue("@fio", row[1].ToString());
                        cmd.Parameters.AddWithValue("@login", row[2].ToString());
                        cmd.Parameters.AddWithValue("@position", row[3].ToString());
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private void GroupAndExportToExcel(DataTable data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var groups = data.AsEnumerable().GroupBy(row => row[3].ToString()).OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    string positionName = string.IsNullOrEmpty(group.Key) ? "Без должности" : group.Key;
                    if (positionName.Length > 31)
                        positionName = positionName.Substring(0, 31);

                    var worksheet = package.Workbook.Worksheets.Add(positionName);

                    worksheet.Cells[1, 1].Value = "Код клиента";
                    worksheet.Cells[1, 2].Value = "ФИО";
                    worksheet.Cells[1, 3].Value = "Логин";

                    int row = 2;
                    foreach (var item in group)
                    {
                        worksheet.Cells[row, 1].Value = item[0].ToString();
                        worksheet.Cells[row, 2].Value = item[1].ToString();
                        worksheet.Cells[row, 3].Value = item[2].ToString();
                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}