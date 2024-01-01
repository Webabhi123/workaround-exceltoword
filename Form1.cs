using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
//using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;
using WordprocessingText = DocumentFormat.OpenXml.Wordprocessing.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using System.Linq.Expressions;
using System.IO;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        private System.Windows.Forms.RichTextBox ResultsRichTextBox;
        public Form1()
        {
            InitializeComponent();
            ResultsRichTextBox = new RichTextBox();

        }

        private void ProcessButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Open Excel file dialog
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                    openFileDialog.Title = "Select an Excel File";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string excelFilePath = openFileDialog.FileName;

                        // Read Excel data
                        List<Exceltoword.ExcelRowData> excelData = ReadExcelUsingEpplus(excelFilePath);

                        // Open Word file dialog
                        using (OpenFileDialog openWordDialog = new OpenFileDialog())
                        {
                            openWordDialog.Filter = "Word Files|*.docx";
                            openWordDialog.Title = "Select a Word Document";

                            if (openWordDialog.ShowDialog() == DialogResult.OK)
                            {
                                string wordTemplatePath = openWordDialog.FileName;

                                // Create a copy of the original Word document
                                string copyPath = $"{Path.GetDirectoryName(wordTemplatePath)}\\Copy_{Path.GetFileName(wordTemplatePath)}";
                                File.Copy(wordTemplatePath, copyPath, true);

                                // Process each row of Excel data
                                int rowIndex = 1; // Add a counter to generate a unique identifier for each row

                                foreach (var excelRow in excelData)
                                {
                                    // Load the copied Word document for each row
                                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(copyPath, true))
                                    {
                                        // Access Word document content
                                        Body body = wordDoc.MainDocumentPart.Document.Body;

                                        //ModifyCookstoveModel(body,"Cookstove model", excelRow.CookstoveModel);
                                        // Replace placeholders in Word document based on Excel data
                                        ReplacePlaceholderInWordDocument(body, "Name", excelRow.Name);
                                        ReplacePlaceholderInWordDocument(body, "Contact", excelRow.Contact);
                                        ReplacePlaceholderInWordDocument(body, "District", excelRow.District);
                                        ReplacePlaceholderInWordDocument(body, "Village", excelRow.Village);
                                        ReplacePlaceholderInWordDocument(body, "Date of installation", excelRow.DateofInstallation);
                                        ReplacePlaceholderInWordDocument(body, "Stove serial number", excelRow.StoveSerialNumber);
                                        ReplacePlaceholderInWordDocument(body, "Number of household members", excelRow.TotalHouseHoldMember);
                                        ReplacePlaceholderInWordDocument(body, "Adults", excelRow.Adults);
                                        ReplacePlaceholderInWordDocument(body, "Children (up to 14 years)", excelRow.NumberofChildren);
                                        //ReplacePlaceholderInWordDocument(body, "Date", excelRow.Date);

                                        // Modify Cookstove model directly
                                        ReplaceWordInWordDocument(body, "Cookstove model", "Greenway Jumbo");
                                        ReplaceWordInWordDocument(body, "Price paid for the improved cookstove", "INR 300");


                                        // Save modified Word document with a unique name
                                        string outputWordPath = $"{excelRow.Name}_{excelRow.Contact}_{excelRow.Village}_Modified.docx";
                                        wordDoc.SaveAs(outputWordPath);
                                    }
                                }
                                rowIndex++;
                                // Display completion message
                                //ResultsRichTextBox.Text = $"Modified Word Documents saved at the respective paths.";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<Exceltoword.ExcelRowData> ReadExcelUsingEpplus(string filePath)
        {
            List<Exceltoword.ExcelRowData> data = new List<Exceltoword.ExcelRowData>();

            using (var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                    for (int row = 2; row < worksheet.Dimension.Rows; row++) //Assume data starts second row
                    {
                    Exceltoword.ExcelRowData currentrow = new Exceltoword.ExcelRowData();

                    currentrow.Name = worksheet.Cells[row, 3].Text+ worksheet.Cells[row, 4].Text;
                    currentrow.Contact = worksheet.Cells[row, 5].Text;
                    currentrow.District = worksheet.Cells[row, 14].Text;
                    currentrow.Village = worksheet.Cells[row, 13].Text;
                    currentrow.DateofInstallation = worksheet.Cells[row, 2].Text;
                    currentrow.StoveSerialNumber=worksheet.Cells[row, 10].Text;
                    currentrow.TotalHouseHoldMember = worksheet.Cells[row, 7].Text;
                    currentrow.Adults = worksheet.Cells[row, 9].Text;
                    currentrow.NumberofChildren = worksheet.Cells[row, 8].Text;
                    currentrow.Date = worksheet.Cells[row, 1].Text;

                    data.Add(currentrow);
                }
                
            }
            return data;
        }
        private void ReplaceWordInWordDocument(Body body, string placeholder, string newValue)
        {
            foreach (var textElement in body.Descendants<WordprocessingText>())
            {
                string originalText = textElement.Text;

                if (textElement.Text.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase)!=-1)
                {
                    // Replace the placeholder with the specified value
                    textElement.Text = $"{placeholder}: {newValue}";
                }
            }
        }

        private void ReplacePlaceholderInWordDocument(Body body, string placeholder,string value)
        {
            foreach (var textElement in body.Descendants<WordprocessingText>())
            {
                string originalText = textElement.Text;

                if (textElement.Text.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Check if the placeholder is a person detail
                    //bool isPersonDetail = placeholder.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                      //                    placeholder.Equals("Contact", StringComparison.OrdinalIgnoreCase) ||
                        //                  placeholder.Equals("District", StringComparison.OrdinalIgnoreCase) ||
                          //                placeholder.Equals("Village", StringComparison.OrdinalIgnoreCase);
                          //
                    // Replace the placeholder with the concatenated label, value, and semicolon (if not a person detail)
                    textElement.Text = $"{placeholder}:{value}";
                    //if (textElement.Text.EndsWith(":"))
                        // Append a colon only if there is a value for the placeholder
                        //if (!string.IsNullOrWhiteSpace(value) && isPersonDetail)
                        //{
                          //  textElement.Text += ":";
                        //}
                }
            }
            foreach (var table in body.Descendants<Table>())
            {
                foreach (var cell in table.Descendants<TableCell>())
                {
                    foreach (var paragraph in cell.Descendants<Paragraph>())
                    {
                        HandleParagraph(paragraph, placeholder, value);
                    }
                }
            }
            foreach (var paragraph in body.Descendants<Paragraph>().Where(p => !p.Ancestors<Table>().Any()))
            {
                HandleParagraph(paragraph, placeholder, value);
            }

        }
        private void HandleParagraph(Paragraph paragraph, string placeholder, string value)
        {
            foreach (var textElement in paragraph.Descendants<WordprocessingText>())
            {
                string originalText = textElement.Text;

                if (textElement.Text.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Check if the placeholder is a person detail
                    //bool isPersonDetail = placeholder.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                      //                    placeholder.Equals("Contact", StringComparison.OrdinalIgnoreCase) ||
                        //                  placeholder.Equals("District", StringComparison.OrdinalIgnoreCase) ||
                          //                placeholder.Equals("Village", StringComparison.OrdinalIgnoreCase);

                    // Replace the placeholder with the concatenated label, value, and semicolon (if not a person detail)
                    textElement.Text = $"{placeholder}:{value}";
                    //if (textElement.Text.EndsWith(":"))
                    // Append a colon only if there is a value for the placeholder
                    //if (!string.IsNullOrWhiteSpace(value) && isPersonDetail)
                    //{
                    //    textElement.Text += "";
                    //}
                }
            }
        }
     
    }

}

