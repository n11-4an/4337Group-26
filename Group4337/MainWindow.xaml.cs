using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        private DataTable? importedData;
        private List<Employee>? employees;

        // Строка подключения к SQL Server Express
        private string connectionString = @"Server=.\SQLEXPRESS;Database=Lab3_4337;Integrated Security=True;";

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Гильманова Азиза");
            employees = new List<Employee>();
        }

        public class Employee
        {
            public string? Код_сотрудника { get; set; }
            public string? ФИО { get; set; }
            public string? Логин { get; set; }
            public string? Должность { get; set; }
        }

        private void ShowInfo_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Гильманова Азиза\nГруппа: 4337\nВариант 8\nЛР3 + ЛР4", "Об авторе");
        }

        // ==================== EXCEL ====================
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
                    MessageBox.Show($"Импортировано {importedData?.Rows.Count} записей!", "Успех");
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
                    MessageBox.Show("Сначала импортируйте данные из Excel!", "Предупреждение");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "export_po_dolzhnostyam.xlsx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    GroupAndExportToExcel(importedData, filePath);
                    MessageBox.Show("Экспорт в Excel завершен!", "Успех");
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
                    if (positionName.Length > 31) positionName = positionName.Substring(0, 31);
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

        // ==================== JSON + WORD ====================
        private void ImportJsonButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON files (*.json)|*.json";
                openFileDialog.Title = "Выберите файл 4.json";

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    string jsonContent = File.ReadAllText(filePath);
                    employees = JsonSerializer.Deserialize<List<Employee>>(jsonContent);
                    SaveEmployeesToDatabase(employees);
                    MessageBox.Show($"Импортировано {employees?.Count} сотрудников из JSON!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте JSON: {ex.Message}", "Ошибка");
            }
        }

        private void SaveEmployeesToDatabase(List<Employee>? employees)
        {
            if (employees == null) return;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand(@"
                    IF EXISTS (SELECT * FROM sys.tables WHERE name = 'Employees')
                        DELETE FROM Employees", conn))
                    cmd.ExecuteNonQuery();

                using (SqlCommand cmd = new SqlCommand(@"
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Employees')
                        CREATE TABLE Employees (
                            Id INT IDENTITY(1,1) PRIMARY KEY,
                            Код_сотрудника NVARCHAR(50),
                            ФИО NVARCHAR(200),
                            Логин NVARCHAR(100),
                            Должность NVARCHAR(100)
                        )", conn))
                    cmd.ExecuteNonQuery();

                foreach (var emp in employees)
                {
                    string sql = @"INSERT INTO Employees (Код_сотрудника, ФИО, Логин, Должность) 
                                   VALUES (@code, @fio, @login, @position)";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@code", emp.Код_сотрудника ?? "");
                        cmd.Parameters.AddWithValue("@fio", emp.ФИО ?? "");
                        cmd.Parameters.AddWithValue("@login", emp.Логин ?? "");
                        cmd.Parameters.AddWithValue("@position", emp.Должность ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private void ExportWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employees == null || employees.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте JSON данные!", "Предупреждение");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word files (*.docx)|*.docx";
                saveFileDialog.FileName = "export_po_dolzhnostyam.docx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    GroupAndExportToWord(employees, filePath);
                    MessageBox.Show("Экспорт в Word завершен!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка");
            }
        }

        private void GroupAndExportToWord(List<Employee> employees, string filePath)
        {
            var groups = employees.GroupBy(e => e.Должность).OrderBy(g => g.Key);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                foreach (var group in groups)
                {
                    // Заголовок
                    Paragraph titlePara = new Paragraph();
                    Run titleRun = new Run();
                    titleRun.AppendChild(new Text($"=== {group.Key} ==="));
                    titleRun.RunProperties = new RunProperties(new Bold());
                    titlePara.AppendChild(titleRun);
                    body.AppendChild(titlePara);

                    body.AppendChild(new Paragraph(new Run(new Break())));

                    // Таблица
                    Table table = new Table();
                    TableProperties tblProp = new TableProperties(
                        new TableBorders(
                            new TopBorder() { Val = BorderValues.Single, Size = 1 },
                            new BottomBorder() { Val = BorderValues.Single, Size = 1 },
                            new LeftBorder() { Val = BorderValues.Single, Size = 1 },
                            new RightBorder() { Val = BorderValues.Single, Size = 1 },
                            new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 1 },
                            new InsideVerticalBorder() { Val = BorderValues.Single, Size = 1 }
                        )
                    );
                    table.AppendChild(tblProp);

                    // Заголовки таблицы
                    TableRow headerRow = new TableRow();
                    AddCell(headerRow, "Код сотрудника", "2000");
                    AddCell(headerRow, "ФИО", "4000");
                    AddCell(headerRow, "Логин", "3000");
                    table.AppendChild(headerRow);

                    // Данные
                    foreach (var emp in group)
                    {
                        TableRow dataRow = new TableRow();
                        AddCell(dataRow, emp.Код_сотрудника ?? "", "2000");
                        AddCell(dataRow, emp.ФИО ?? "", "4000");
                        AddCell(dataRow, emp.Логин ?? "", "3000");
                        table.AppendChild(dataRow);
                    }

                    body.AppendChild(table);

                    // Количество сотрудников
                    Paragraph countPara = new Paragraph();
                    countPara.AppendChild(new Run(new Text($"Всего сотрудников: {group.Count()}")));
                    body.AppendChild(countPara);

                    // Разрыв страницы
                    body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                mainPart.Document.Save();
            }
        }

        private void AddCell(TableRow row, string text, string width)
        {
            TableCell cell = new TableCell();
            cell.AppendChild(new Paragraph(new Run(new Text(text))));
            cell.TableCellProperties = new TableCellProperties(new TableCellWidth() { Width = width });
            row.AppendChild(cell);
        }
    }
}