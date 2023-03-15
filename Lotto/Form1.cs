using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LottoNumberGenerator
{
    public partial class Form1 : Form
    {
        private string filePath = "";
        private Dictionary<int, int> numberFrequencies = new Dictionary<int, int>();
        private List<int> numbers = new List<int>();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenExcel_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Select an Excel file";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                lblFilePath.Text = filePath;

                // Read numbers from Excel file
                ReadNumbersFromExcel(filePath);
            }
        }

        private void ReadNumbersFromExcel(string filePath)
        {
            // Clear existing numbers
            numberFrequencies.Clear();
            numbers.Clear();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            int number = 0;
                            if (int.TryParse(reader.GetValue(i).ToString(), out number))
                            {
                                if (!numberFrequencies.ContainsKey(number))
                                {
                                    numberFrequencies.Add(number, 0);
                                }

                                numberFrequencies[number]++;
                            }
                        }
                    }
                }
            }

            // Calculate the total number of occurrences
            int totalOccurrences = numberFrequencies.Values.Sum();

            // Calculate the frequency of each number and add it to the list
            foreach (int number in numberFrequencies.Keys)
            {
                int frequency = (int)((double)numberFrequencies[number] / totalOccurrences * 100);
                for (int i = 0; i < frequency; i++)
                {
                    numbers.Add(number);
                }
            }
        }

        private void btnGenerateNumbers_Click(object sender, System.EventArgs e)
        {
            // Generate 5 sets of numbers
            for (int i = 0; i < 5; i++)
            {
                List<int> set = new List<int>();
                for (int j = 0; j < 7; j++)
                {
                    // Choose a random number from the list
                    int index = new Random().Next(0, numbers.Count);
                    int number = numbers[index];

                    // Add the number to the set and remove it from the list
                    set.Add(number);
                    numbers.RemoveAt(index);
                }

                // Sort the set and add it to the list box
                set.Sort();
                listBoxNumbers.Items.Add(string.Join(", ", set));
            }

            // Update the status label
            lblStatus.Text = "Numbers generated.";
        }
    }
}
