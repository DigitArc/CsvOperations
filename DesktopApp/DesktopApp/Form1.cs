using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.TypeConversion;

namespace DesktopApp
{
    public partial class Form1 : Form
    {
        private string _selectedCsvFileName;
        public Form1()
        {
            InitializeComponent();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Text = "Csv Operations v1";
        }
        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Multiselect = false, Filter = "CSV files (*.csv)|*.csv" };
            DialogResult result = openFileDialog.ShowDialog();
            if (result != DialogResult.OK) return;
            try
            {
                _selectedCsvFileName = openFileDialog.FileName;
                using var reader = new StreamReader(_selectedCsvFileName);
                using var csv = new CsvReader(reader);
                int locationX = 14;
                int locationY = 120;
                csv.Read();
                string headerField = csv.GetField(0);
                int counter = 0;
                string[] headers = headerField.Split(",");
                foreach (var header in headers)
                {
                    var textBox = new TextBox
                    {
                        Location = new Point(locationX, locationY),
                        Text = header,
                        Name = "txt",
                        Size = new Size(145, 23)
                    };
                    Controls.Add(textBox);
                    counter++;
                    if (counter % 4 != 0)
                    {
                        locationX += 170;
                    }
                    else
                    {
                        locationX = 14;
                        locationY += 40;
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{exception.Message}");
            }
        }
        private void SetLoading(bool displayLoader)
        {
            if (displayLoader)
            {
                Invoke((MethodInvoker)delegate
                {
                    Cursor = Cursors.WaitCursor;
                });
            }
            else
            {
                Invoke((MethodInvoker)delegate
                {
                    Cursor = Cursors.Default;
                });
            }
        }

        private void WriteHeader(IWriter csvWriter)
        {
            var textBoxes = Controls.Find("txt", true);
            foreach (var item in textBoxes)
            {
                csvWriter.WriteField(item.Text);
            }
            csvWriter.NextRecord();
        }

        private static void WriteContents(IReader csvReader, IWriter csvWriter)
        {
            csvReader.Read();
            while (csvReader.Read())
            {
                var fields = csvReader.GetField(0);
                var fieldSplit = fields.Split(",");
                foreach (var str in fieldSplit)
                {
                    if (DateTime.TryParse(str, out DateTime date))
                    {
                        csvWriter.WriteField(date);
                    }
                    else if (double.TryParse(str, out double value))
                    {
                        csvWriter.WriteField(value);
                    }
                    else
                        csvWriter.WriteField(str);
                }
                csvWriter.NextRecord();
            }
        }
        private void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                SetLoading(true);
                SetDisableControls(true);
                var fileName = $"{Guid.NewGuid()}.csv";
                using var streamReader = new StreamReader(_selectedCsvFileName);
                using var csvReader = new CsvReader(streamReader);
                using var fileStream = File.Create(fileName);
                using var streamWriter = new StreamWriter(fileStream) { AutoFlush = true };
                using var csvWriter = new CsvWriter(streamWriter);
                var options = new TypeConverterOptions { Formats = new[] { "MM/dd/yyyy hh:mm" } };
                csvWriter.Configuration.TypeConverterOptionsCache.AddOptions<DateTime>(options);
                WriteHeader(csvWriter);
                WriteContents(csvReader, csvWriter);
                SetLoading(false);
                SetDisableControls(false);
                OpenCreatedFile(fileName);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                SetLoading(false);
                SetDisableControls(false);
            }
        }
        private static void OpenCreatedFile(string fileName)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                Arguments = $"{AppDomain.CurrentDomain.BaseDirectory}{fileName}",
                FileName = "explorer.exe"
            };
            Process.Start(startInfo);

        }
        private void SetDisableControls(bool state)
        {
            foreach (Control c in Controls)
            {
                if (c.GetType() == typeof(TextBox))
                {
                    c.Visible = !state;
                }
                c.Enabled = !state;
            }
        }
    }

}
