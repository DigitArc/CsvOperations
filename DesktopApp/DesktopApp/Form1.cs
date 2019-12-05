using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.TypeConversion;

namespace DesktopApp
{

    public partial class Form1 : Form
    {
        public static string[] BannedWords =
        {
            "Bad",
            "Konfig",
            "Error",
            "Not Connect",
            "var7",
            "Invalid Expression Syntax",
            "#WERT",
            "#NAME?",
            "Open Auto",
            "Stopped Manual"
        };
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
                        var formattedDate = date.ToString("dd/MM/yyyy HH:mm:ss");
                        csvWriter.WriteField(formattedDate);
                    }
                    else if (double.TryParse(str, out double value))
                    {
                        csvWriter.WriteField(value);
                    }
                    else
                    {
                        var field = ClearField(str);
                        field = field.Replace("ON", 1.ToString()).Replace("OFF", 0.ToString());
                        csvWriter.WriteField(field);
                    }
                }
                csvWriter.NextRecord();
            }
        }

        private static string ClearField(string str)
        {
            foreach (var bannedWord in BannedWords)
            {
                if (str.Contains(bannedWord))
                {
                    str = str.Replace(bannedWord, string.Empty);
                }
            }

            return str;
        }

        private void ClearTextInput()
        {
            var textBoxes = Controls.Find("txt", true);
            foreach (var item in textBoxes)
            {
                Controls.Remove(item);
            }
        }
        private void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_selectedCsvFileName))
                {
                    MessageBox.Show("Please select file");
                    return;
                }
                var button = (Button) sender;
                button.Visible = false;
                ClearTextInput();
                SetLoading(true);
                var fileName = $"{Guid.NewGuid()}.csv";
                using var fileStream = File.Create(fileName);
                using var streamReader = new StreamReader(_selectedCsvFileName);
                using var streamWriter = new StreamWriter(fileStream) { AutoFlush = true };
                using var csvWriter = new CsvWriter(streamWriter);
                using var csvReader = new CsvReader(streamReader);
                var options = new TypeConverterOptions { Formats = new[] { "MM/dd/yyyy hh:mm" } };
                csvWriter.Configuration.TypeConverterOptionsCache.AddOptions<DateTime>(options);
                WriteHeader(csvWriter);
                WriteContents(csvReader, csvWriter);
                SetLoading(false);
                OpenCreatedFile(fileName);
                _selectedCsvFileName = string.Empty;
                button.Visible = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                SetLoading(false);
                ClearTextInput();
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
