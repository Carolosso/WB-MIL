using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;

namespace WBMIL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        public static void ChangeTextInCell(string filepath, string txt, int cell_index, int p_index)
        {
            // Use the file name and path passed in as an argument to 
            // open an existing document.        
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(filepath, true))
            {
                // Find the first table in the document.
                Table table =
                    doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                // Find the second row in the table.
                TableRow row = table.Elements<TableRow>().ElementAt(3);

                // Find the third cell in the row.
                TableCell cell = row.Elements<TableCell>().ElementAt(cell_index);

                // Find the first paragraph in the table cell.
                Paragraph p = cell.Elements<Paragraph>().ElementAt(p_index);

                // Find the first run in the paragraph.
                Run r = p.Elements<Run>().First();

                // Set the text for the run.
                Text t = r.Elements<Text>().First();
                t.Text = txt;
            }
        }

        public static String get_file_path()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Wybierz miejsce zapisu";
            string sSelectedPath = "";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                sSelectedPath = fbd.SelectedPath;
            }

            return sSelectedPath;
        }

        public String return_radio_rozszerzenie()
        {
            if (radioButton1.Checked) return radioButton1.Text;
            else return radioButton2.Text;
        }



        public void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        { 
            try
            {


                String dst_path = get_file_path();
                String ext = return_radio_rozszerzenie();
                String filename =@"Resources\szablon1.docx";
                string filetocopy = @"Backup\szablon1.docx";
                File.Copy(filetocopy, filename,true);
                int count = Convert.ToInt32(numericUpDown1.Value);
                String numer = count.ToString();
                ChangeTextInCell(filename, numer, 0, 1);
                ChangeTextInCell(filename, dateTimePicker1.Text, 1, 0);
                ChangeTextInCell(filename, Pole1.Text, 2, 0);
                ChangeTextInCell(filename, Pole2.Text, 3, 1);
                ChangeTextInCell(filename, Pole3.Text, 4, 1);
                ChangeTextInCell(filename, Pole4.Text, 5, 1);
                ChangeTextInCell(filename, Pole5.Text, 6, 1);
                ChangeTextInCell(filename, Pole6.Text, 7, 1);
                ChangeTextInCell(filename, Pole7.Text, 8, 0);
                ChangeTextInCell(filename, Pole8.Text, 9, 0);
                ChangeTextInCell(filename, Pole8.Text, 10, 0);
                ChangeTextInCell(filename, Pole10.Text, 11, 0);
                ChangeTextInCell(filename, Pole11.Text, 12, 0);

              
                    if (return_radio_rozszerzenie().Equals(".pdf"))
                    {
                        //Loads an existing Word document
                        WordDocument wordDocument = new WordDocument(filename, FormatType.Docx);

                        //Creates an instance of the DocToPDFConverter
                        DocToPDFConverter converter = new DocToPDFConverter();

                        //Converts Word document into PDF document
                        PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);

                        //Releases all resources used by DocToPDFConverter
                        converter.Dispose();

                        //Closes the instance of document objects
                        wordDocument.Close();

                        //Saves the PDF file 
                        pdfDocument.Save(dst_path + "\\" + Path.GetFileName(textBox1.Text + ext));

                        //Closes the instance of document objects
                        pdfDocument.Close(true);
                    }
                    else File.Copy(filename, dst_path + "\\" + Path.GetFileName(textBox1.Text + ext));
                

                  MessageBox.Show("Zapisano pomyślnie!");
                    File.Delete(filename);
            }
            catch {
                MessageBox.Show("Błąd zapisu!");
            }
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }
    }
}
