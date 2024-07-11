using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text; // Aggiungi questa direttiva
using System.Windows.Forms;
using OfficeOpenXml;
namespace Excel_Baxter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        private void btnCreateExcel_Click(object sender, EventArgs e)
        {
            using (var package = new ExcelPackage())
            {
                // Crea un nuovo foglio di lavoro
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Aggiungi dati al foglio di lavoro
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Age";

                worksheet.Cells[2, 1].Value = 1;
                worksheet.Cells[2, 2].Value = "John Doe";
                worksheet.Cells[2, 3].Value = 30;

                worksheet.Cells[3, 1].Value = 2;
                worksheet.Cells[3, 2].Value = "Jane Doe";
                worksheet.Cells[3, 3].Value = 25;

                // Salva il file Excel
                var file = new FileInfo("output.xlsx");
                package.SaveAs(file);

                MessageBox.Show("File Excel creato con successo!");
            }
        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Seleziona un file Excel";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog1.FileName;
                var fileInfo = new FileInfo(filePath);

                // Verifica se il file esiste
                if (!fileInfo.Exists)
                {
                    MessageBox.Show("Il file selezionato non esiste.");
                    return;
                }

                try
                {
                    // Leggere il file selezionato dall'utente
                    DataTable selectedFileData = ReadExcelFile(fileInfo);

                    // Stampa i nomi delle colonne per debug
                    if (selectedFileData != null)
                    {
                        string columnNames = string.Join(", ", selectedFileData.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                        MessageBox.Show($"Colonne nel file selezionato: {columnNames}");
                    }

                    // Leggere il file anagrafica
                    string anagraficaPath = @"C:\Users\r.pierelli\Downloads\Baxter\Anagrafica di conversione Bill to\anagrafica.xlsx";
                    var anagraficaInfo = new FileInfo(anagraficaPath);

                    if (!anagraficaInfo.Exists)
                    {
                        MessageBox.Show("Il file anagrafica non esiste.");
                        return;
                    }

                    DataTable anagraficaData = ReadExcelFile(anagraficaInfo);

                    // Stampa i nomi delle colonne per debug
                    if (anagraficaData != null)
                    {
                        string columnNames = string.Join(", ", anagraficaData.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                        MessageBox.Show($"Colonne nel file anagrafica: {columnNames}");
                    }

                    // Creare un nuovo file Excel con i dati elaborati
                    CreateNewExcelWithProcessedData(selectedFileData, anagraficaData);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Errore durante la lettura del file Excel: {ex.Message}");
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private DataTable ReadExcelFile(FileInfo fileInfo)
        {
            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        MessageBox.Show($"Il file {fileInfo.Name} non contiene fogli di lavoro!");
                        return null;
                    }

                    var worksheet = package.Workbook.Worksheets[0];

                    // Controlla se il foglio di lavoro è vuoto
                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show($"Il foglio di lavoro nel file {fileInfo.Name} è vuoto!");
                        return null;
                    }

                    var dataTable = new DataTable();
                    var columnNames = new HashSet<string>();

                    // Verifica la prima riga e leggi i nomi delle colonne
                    int columns = worksheet.Dimension.Columns;
                    int rows = worksheet.Dimension.Rows;
                    for (int col = 1; col <= columns; col++)
                    {
                        var columnName = worksheet.Cells[1, col].Text;
                        if (string.IsNullOrWhiteSpace(columnName))
                        {
                            columnName = $"Column{col}";
                        }

                        // Gestisce i nomi delle colonne duplicate
                        if (columnNames.Contains(columnName))
                        {
                            int duplicateCount = 1;
                            string newColumnName;
                            do
                            {
                                newColumnName = $"{columnName}_{duplicateCount}";
                                duplicateCount++;
                            } while (columnNames.Contains(newColumnName));
                            columnName = newColumnName;
                        }

                        columnNames.Add(columnName);
                        dataTable.Columns.Add(columnName);
                    }

                    // Aggiungi righe alla DataTable
                    for (int row = 2; row <= rows; row++)
                    {
                        var newRow = dataTable.NewRow();
                        for (int col = 1; col <= columns; col++)
                        {
                            newRow[col - 1] = worksheet.Cells[row, col].Text;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    return dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante la lettura del file {fileInfo.Name}: {ex.Message}");
                return null;
            }
        }

        private void CreateNewExcelWithProcessedData(DataTable selectedFileData, DataTable anagraficaData)
        {
            if (selectedFileData == null || anagraficaData == null)
            {
                MessageBox.Show("Errore nella lettura dei file Excel.");
                return;
            }

            // Verifica che entrambe le DataTable contengano le colonne necessarie
            if (!selectedFileData.Columns.Contains("cutcode"))
            {
                MessageBox.Show("Il file selezionato non contiene la colonna 'cutcode'.");
                return;
            }

            if (!anagraficaData.Columns.Contains("cutcode") || !anagraficaData.Columns.Contains("VTV Customer"))
            {
                MessageBox.Show("Il file anagrafica non contiene le colonne necessarie 'cutcode' o 'VTV Customer'.");
                return;
            }

            if (!selectedFileData.Columns.Contains("Price BIll to"))
            {
                MessageBox.Show("Il file selezionato non contiene la colonna 'Price BIll to'.");
                return;
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ProcessedData");

                // Copia le colonne del foglio selezionato
                for (int col = 0; col < selectedFileData.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = selectedFileData.Columns[col].ColumnName;
                }

                int newRow = 2;

                foreach (DataRow row in selectedFileData.Rows)
                {
                    string cutcode = row["cutcode"].ToString().Trim();

                    if (string.IsNullOrEmpty(cutcode))
                    {
                        // Se cutcode è vuoto, copia la riga senza fare la ricerca
                        for (int col = 0; col < selectedFileData.Columns.Count; col++)
                        {
                            worksheet.Cells[newRow, col + 1].Value = row[col];
                        }
                        newRow++;
                        continue;
                    }

                    var matchingRows = anagraficaData.AsEnumerable()
                        .Where(r => r.Field<string>("cutcode") == cutcode);

                    foreach (var anagraficaRow in matchingRows)
                    {
                        for (int col = 0; col < selectedFileData.Columns.Count; col++)
                        {
                            worksheet.Cells[newRow, col + 1].Value = row[col];
                        }
                        worksheet.Cells[newRow, selectedFileData.Columns["Price BIll to"].Ordinal + 1].Value = anagraficaRow["VTV Customer"];
                        newRow++;
                    }
                }

                // Salva il nuovo file Excel
                var file = new FileInfo("ProcessedData.xlsx");
                package.SaveAs(file);

                MessageBox.Show("File Excel elaborato creato con successo!");
            }
        }

        private void confronto_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    var filePaths = Directory.GetFiles(selectedPath, "*.xlsx");

                    if (filePaths.Length == 0)
                    {
                        MessageBox.Show("Nessun file Excel trovato nella cartella selezionata.");
                        return;
                    }

                    string referenceFilePath = @"C:\Users\r.pierelli\Downloads\Baxter\R5640722_KCEUIT01_1569260_PDF - Copia.xlsx";
                    var referenceFileInfo = new FileInfo(referenceFilePath);

                    if (!referenceFileInfo.Exists)
                    {
                        MessageBox.Show("Il file di riferimento non esiste.");
                        return;
                    }

                    DataTable referenceData = ReadExcelFile(referenceFileInfo);
                    if (referenceData == null)
                    {
                        MessageBox.Show("Errore nella lettura del file di riferimento.");
                        return;
                    }

                    List<DataTable> dataTables = new List<DataTable>();
                    foreach (var filePath in filePaths)
                    {
                        var fileInfo = new FileInfo(filePath);
                        DataTable dataTable = ReadExcelFile(fileInfo);
                        if (dataTable != null)
                        {
                            dataTables.Add(dataTable);
                        }
                    }

                    string outputFolderPath = GetOutputFolderPath();
                    if (string.IsNullOrEmpty(outputFolderPath))
                    {
                        MessageBox.Show("Cartella di destinazione non selezionata.");
                        return;
                    }

                    string logFilePath = Path.Combine(outputFolderPath, "ConfrontoLog.txt");
                    CompareAndGenerateReports(dataTables, referenceData, outputFolderPath, logFilePath);
                }
            }
        }

        private string GetOutputFolderPath()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Seleziona la cartella di destinazione per i report";
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return null;
        }

        private void CompareAndGenerateReports(List<DataTable> dataTables, DataTable referenceData, string outputFolderPath, string logFilePath)
        {
            var matchingRows = new List<DataRow>();
            var nonMatchingRows = new List<(DataRow, string)>();

            foreach (var dataTable in dataTables)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    string billto = null;
                    if (row.Table.Columns.Contains("billtobaxter") && !string.IsNullOrWhiteSpace(row["billtobaxter"].ToString()))
                    {
                        billto = row["billtobaxter"].ToString().Trim();
                    }
                    else if (row.Table.Columns.Contains("billtovantive") && !string.IsNullOrWhiteSpace(row["billtovantive"].ToString()))
                    {
                        billto = row["billtovantive"].ToString().Trim();
                    }

                    string numerogara = row.Table.Columns.Contains("numerogara") ? row["numerogara"].ToString().Trim() : "";
                    string codiceprodotto = row.Table.Columns.Contains("codiceprodotto") ? row["codiceprodotto"].ToString().Trim() : "";

                    var referenceRow = referenceData.AsEnumerable()
                        .FirstOrDefault(r =>
                            (r.Table.Columns.Contains("billtobaxter") && r.Field<string>("billtobaxter") == billto) ||
                            (r.Table.Columns.Contains("billtovantive") && r.Field<string>("billtovantive") == billto) &&
                            (string.IsNullOrEmpty(numerogara) || r.Field<string>("numerogara") == numerogara) &&
                            (string.IsNullOrEmpty(codiceprodotto) || r.Field<string>("codiceprodotto") == codiceprodotto));

                    if (referenceRow != null)
                    {
                        matchingRows.Add(row);
                    }
                    else
                    {
                        nonMatchingRows.Add((row, "Key not found in reference file"));
                    }
                }
            }

            CreateExcelReport(Path.Combine(outputFolderPath, "MatchingRows.xlsx"), matchingRows);
            CreateNonMatchingExcelReport(Path.Combine(outputFolderPath, "NonMatchingRows.xlsx"), nonMatchingRows);

            MessageBox.Show("Confronto completato. I report sono stati generati.");
        }

        private void LogRowDetails(DataRow row, string reason, string logFilePath)
        {
            string logMessage = $"Row Details: {string.Join(", ", row.ItemArray)} - Reason: {reason}";
            Console.WriteLine(logMessage);

            // Scrivi nel file di log
            using (StreamWriter sw = new StreamWriter(logFilePath, true))
            {
                sw.WriteLine(logMessage);
            }
        }
        private bool RowMatches(DataRow row1, DataRow row2, out string mismatchReason)
        {
            mismatchReason = null;
            StringBuilder noteBuilder = new StringBuilder();

            for (int i = 0; i < row1.Table.Columns.Count; i++)
            {
                if (!row1.Table.Columns[i].ColumnName.Equals(row2.Table.Columns[i].ColumnName))
                {
                    mismatchReason = $"Column name mismatch at column index {i}";
                    return false;
                }

                if (!row1[i].Equals(row2[i]))
                {
                    noteBuilder.Append($"Mismatch at {row1.Table.Columns[i].ColumnName}: {row1[i]} != {row2[i]}. ");
                }
            }

            mismatchReason = noteBuilder.ToString();
            return string.IsNullOrEmpty(mismatchReason);
        }
        private void CreateNonMatchingExcelReport(string fileName, List<(DataRow Row, string Reason)> rows)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Report");

                if (rows.Count > 0)
                {
                    // Aggiungi le intestazioni delle colonne
                    for (int col = 0; col < rows[0].Row.Table.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = rows[0].Row.Table.Columns[col].ColumnName;
                    }
                    worksheet.Cells[1, rows[0].Row.Table.Columns.Count + 1].Value = "Note";

                    // Aggiungi le righe
                    for (int row = 0; row < rows.Count; row++)
                    {
                        for (int col = 0; col < rows[row].Row.Table.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = rows[row].Row[col];
                        }
                        worksheet.Cells[row + 2, rows[row].Row.Table.Columns.Count + 1].Value = rows[row].Reason;
                    }
                }

                var file = new FileInfo(fileName);
                package.SaveAs(file);
            }
        }




        private void CreateExcelReport(string fileName, List<DataRow> rows)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Report");

                if (rows.Count > 0)
                {
                    // Aggiungi le intestazioni delle colonne
                    for (int col = 0; col < rows[0].Table.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = rows[0].Table.Columns[col].ColumnName;
                    }

                    // Aggiungi le righe
                    for (int row = 0; row < rows.Count; row++)
                    {
                        for (int col = 0; col < rows[row].Table.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = rows[row][col];
                        }
                    }
                }

                var file = new FileInfo(fileName);
                package.SaveAs(file);
            }
        }
    }
}
