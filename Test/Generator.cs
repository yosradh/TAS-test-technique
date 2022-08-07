using ExcelDataReader;
using IronPdf;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TAS_Test
{

    public class Employee
    {
        public string Name;
        public string Surname;
        public float TotalSalary
        {
            get { return M1 + M2 + M3; }
        }

        public float M1;
        public float M2;
        public float M3;

        public string GetEmployeeText(int index)
        {
            string employeeString = string.Format("{0:000}", index);

            string EmployeeName = this.Name + "" + this.Surname;

            employeeString += string.Format("{0,-30}", EmployeeName);

            employeeString += string.Format("{0:0000000000}", this.TotalSalary * 1000);

            return employeeString;
        }
    }

    public enum AlignmentType : int
    {
        ALIGN_LEFT = 0,
        ALIGN_CENTER = 1,
        ALIGN_RIGHT = 2,
        ALIGN_JUSTIFIED = 3
    }

    public partial class Generator : System.Windows.Forms.Form
    {
        //Text File name
        private string TextFileName = "Cnss.txt";

        //PDF File name
        private string PDFFileName = "GeneratedPDF.pdf";

  
        //Set Write Fonts
        private Font policeTitre = new Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 17f);
        private Font FontTab = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD);
        private Font FontSalairy = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f);
        private Font FontNom = new Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLDITALIC);

        //liste of employees 
        public List<Employee> employees = new List<Employee>();

        public string FilesPath;

        //initialiser form  
        public Generator()
        {
            InitializeComponent();

    }

    #region Employees Region
    /// <summary>
    /// Create Employees From XLS Table and store them in Employees List
    /// </summary>
    /// <param name="table"></param>
    public void CreateEmployees(DataTable table)
        {
            employees.Clear();

            Employee FirstEmployee = new Employee();

            FirstEmployee.Name = table.Columns[0].ColumnName;
            FirstEmployee.Surname = table.Columns[1].ColumnName;

            FirstEmployee.M1 = float.Parse(table.Columns[2].ColumnName);
            FirstEmployee.M2 = float.Parse(table.Columns[3].ColumnName);
            FirstEmployee.M3 = float.Parse(table.Columns[4].ColumnName);


            employees.Add(FirstEmployee);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                Employee employee = new Employee();

                employee.Name = table.Rows[i][table.Columns[0]].ToString();
                employee.Surname = table.Rows[i][table.Columns[1]].ToString();

                employee.M1 = float.Parse((table.Rows[i][table.Columns[2]].ToString()));
                employee.M2 = float.Parse((table.Rows[i][table.Columns[3]].ToString()));
                employee.M3 = float.Parse((table.Rows[i][table.Columns[4]].ToString()));

                employees.Add(employee);
            }
        }
        #endregion

        #region Text file region
        /// <summary>
        /// Get the string format of the employees
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public string GetStringFromEmployees(DataTable table)
        {
            CreateEmployees(table);

            string FileText = string.Empty;

            for (int i = 0; i < employees.Count; i++)
            {
                FileText += employees[i].GetEmployeeText(i + 1) + '\n';
            }
            Console.WriteLine(FileText);

            return FileText;
        }

        #endregion

        #region PDF REGION
        /// <summary>
        /// Add Cell To Table
        /// </summary>
        /// <param name="str">string Data</param>
        /// <param name="f">Font</param>
        /// <param name="table">Table</param>
        /// <param name="alignmentType">AlignmentType : LEFT,RIGHT OR CENTER</param>
        /// <param name="isTitle">IF true,It's a title cell</param>
        /// 
      

        private void AddCelluleToTab(string str, Font f, PdfPTable table,AlignmentType alignmentType = AlignmentType.ALIGN_CENTER, bool isTitle = false)
        {

            PdfPCell cell = new PdfPCell(new Phrase(str, f));
            if (isTitle)
                cell.Padding = 25;
            else
                cell.Padding = 5;

            cell.HorizontalAlignment = (int)alignmentType;
            table.AddCell(cell);
        }

        /// <summary>
        /// Generate the pdf file from xsl file
        /// </summary>
        /// <param name="table"></param>
        private void GenereteFiles(DataTable table)
        {
            // Code for Select the folder to save the file.
            SaveFileDialog SFD = new SaveFileDialog();
            SFD.InitialDirectory = @"C:";
            SFD.Title = "Save Generated Files !";
            SFD.FileName = PDFFileName;
            SFD.RestoreDirectory = true;

            if (SFD.ShowDialog() == DialogResult.OK)
            {
                string path = System.IO.Path.GetDirectoryName(SFD.FileName);

                FilesPath = path;

                CreateEmployees(table);

                CreatePDFFile(Path.Combine(path,PDFFileName));

                File.WriteAllText(Path.Combine(path, TextFileName), GetStringFromEmployees(table));
            }

        }

        public void CreatePDFFile(string path)
        {
            //Create PDF Document 
            //string outfile = Environment.CurrentDirectory + "/"+PDFFileName;


            Console.WriteLine("Path : " + path);

            string outfile = path;

            Document doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream(outfile, FileMode.Create));
            doc.Open();

            //Create Title Paragraph
            Paragraph titre = new Paragraph("Declaration CNSS \n", policeTitre);
            titre.Alignment = Element.ALIGN_CENTER;
            doc.Add(titre);

            //Space between the title and the table
            Paragraph space = new Paragraph(" ", policeTitre);
            space.Alignment = Element.ALIGN_CENTER;
            doc.Add(space);

            AddEmployeesTable(doc);

            //Next page 
            doc.NewPage();

            //Add Total Table
            AddTotalTable(doc);

            //Close document
            doc.Close();
        }
        public void AddEmployeesTable(Document doc)
        {
            var nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberGroupSeparator = " ";

            //Create employees Table
            PdfPTable EmployeesTable = new PdfPTable(5);
            EmployeesTable.WidthPercentage = 100;

            //Add Title Cells
            AddCelluleToTab("Nom", FontTab, EmployeesTable, isTitle: true);
            AddCelluleToTab("Prenom", FontTab, EmployeesTable, isTitle: true);
            AddCelluleToTab("M1", FontTab, EmployeesTable, isTitle: true);
            AddCelluleToTab("M2", FontTab, EmployeesTable, isTitle: true);
            AddCelluleToTab("M3", FontTab, EmployeesTable, isTitle: true);

            //Add Employees Data
            for (int i = 0; i < employees.Count; i++)
            {
                AddCelluleToTab(employees[i].Name, FontNom, EmployeesTable, AlignmentType.ALIGN_LEFT);//Add Name
                AddCelluleToTab(employees[i].Surname, FontNom, EmployeesTable, AlignmentType.ALIGN_LEFT);//Add SurName
                AddCelluleToTab(employees[i].M1.ToString("#,0.000", nfi), FontSalairy, EmployeesTable);//Add M1
                AddCelluleToTab(employees[i].M2.ToString("#,0.000", nfi), FontSalairy, EmployeesTable);//Add M2
                AddCelluleToTab(employees[i].M3.ToString("#,0.000", nfi), FontSalairy, EmployeesTable);//Add M3
            }

            //Add Total Cell
            PdfPCell Total = new PdfPCell(new Phrase("TOTAL", FontTab));
            Total.Colspan = 5;
            Total.HorizontalAlignment = 1;
            EmployeesTable.AddCell(Total);

            //Add Table to the PDF Document
            doc.Add(EmployeesTable);
        }

        public void AddTotalTable(Document doc)
        {
            var nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberGroupSeparator = " ";

            PdfPTable TotalTable = new PdfPTable(1);
            TotalTable.WidthPercentage = 20;
            TotalTable.HorizontalAlignment = 3;

            AddCelluleToTab("TOTAL", FontTab, TotalTable, isTitle: true);

            //Add Cells to total table
            float total = 0f;
            for (int i = 0; i < employees.Count; i++)
            {
                total += employees[i].TotalSalary;
                AddCelluleToTab(employees[i].TotalSalary.ToString("#,0.000", nfi), FontSalairy, TotalTable);
            }

            AddCelluleToTab(total.ToString("#,0.000", nfi), FontTab, TotalTable);
            doc.Add(TotalTable);
        }
        #endregion

        #region File management
        /// <summary>
        /// Open the file with cmd command
        /// </summary>
        /// <param name="path">Path of the file</param>
        /// <param name="FileName">File name with ext</param>
        public void OpenFile(string path, string FileName)
        {
            if (string.IsNullOrEmpty(path)) return;

            var folder = path;
            var fullpath = Path.Combine(folder, FileName);
            Console.WriteLine(fullpath);

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = Path.GetFileName(fullpath);
            psi.WorkingDirectory = Path.GetDirectoryName(fullpath);
            psi.Arguments = "; exit";
            Process.Start(psi);
        }
        #endregion

        #region UI 
        /// <summary>
        /// Get Path to save the file on it
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void parcourir_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Choisir votre fichier ",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "xls files (*.xls)|*.xls",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true

            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pathFichier.Text = openFileDialog.FileName.ToString();
            }
        }
      
        /// <summary>
        /// G
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Generate_Click(object sender, EventArgs e)
        {

            using (var stream = File.Open(pathFichier.Text, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet resultat = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {

                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    
                    GenereteFiles(resultat.Tables[0]);

                    OpenFile(FilesPath, PDFFileName);
                    OpenFile(FilesPath, TextFileName);
                }
            }
        }

        #endregion
    }
}
