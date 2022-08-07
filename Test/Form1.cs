using ExcelDataReader;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Test
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




    public partial class Form1 : Form
    {
        //mon fichier text
        string myPath = @"Cnss.txt";

        //liste des employees 
        public List<Employee> employees = new List<Employee>();

        //initialiser form
        public Form1()
        {
            InitializeComponent();

        }


       

        //*********************************//
        //creer les employees a partir fichier excel 
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
        //*********************************//
        //generation cellule tableau
        private void AddCelluleToTab(string str, Font f, PdfPTable table,AlignmentType alignmentType = AlignmentType.ALIGN_CENTER, bool isTitle = false)
        {

            PdfPCell cell1 = new PdfPCell(new Phrase(str, f));
            if (isTitle)
                cell1.Padding = 25;
            else
                cell1.Padding = 5;

            cell1.HorizontalAlignment = (int)alignmentType;
            table.AddCell(cell1);
        }


        /// <summary>
        /// Generate the pdf file from xsl file
        /// </summary>
        /// <param name="table"></param>
        private void GeneretePDF(DataTable table)
        {

            string outfile = Environment.CurrentDirectory + "/GeneratePDF.pdf";

            var nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberGroupSeparator = " ";


            //creation du document 
            Document doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream(outfile, FileMode.Create));
            doc.Open();

            //Polices d'écriture

            Font policeTitre = new Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 17f);
            Font FontTab = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD);
            Font FontSalairy = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f);
            Font FontNom = new Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLDITALIC);


            //page
            //creation de paragraphe
            Paragraph titre = new Paragraph("Declaration CNSS \n", policeTitre);
            titre.Alignment = Element.ALIGN_CENTER;
            doc.Add(titre);

            Paragraph space = new Paragraph(" ", policeTitre);
            space.Alignment = Element.ALIGN_CENTER;
            doc.Add(space);

            //Creation un tableau 
            PdfPTable tableau1 = new PdfPTable(5);
            tableau1.WidthPercentage = 100;

            //Cellule (les titres de tableau)
            AddCelluleToTab("Nom", FontTab, tableau1, isTitle : true) ;
            AddCelluleToTab("Prenom", FontTab, tableau1, isTitle: true);
            AddCelluleToTab("M1", FontTab, tableau1, isTitle: true);
            AddCelluleToTab("M2", FontTab, tableau1, isTitle: true);
            AddCelluleToTab("M3", FontTab, tableau1, isTitle: true);

            //lister les infos 
            CreateEmployees(table);

            for (int i = 0; i < employees.Count; i++)
            {
                AddCelluleToTab(employees[i].Name, FontNom, tableau1, AlignmentType.ALIGN_LEFT);
                AddCelluleToTab(employees[i].Surname, FontNom, tableau1, AlignmentType.ALIGN_LEFT);
                AddCelluleToTab(employees[i].M1.ToString("#,0.000", nfi), FontSalairy, tableau1);
                AddCelluleToTab(employees[i].M2.ToString("#,0.000", nfi), FontSalairy, tableau1);
                AddCelluleToTab(employees[i].M3.ToString("#,0.000", nfi), FontSalairy, tableau1);

            }
            PdfPCell Total = new PdfPCell(new Phrase("TOTAL", FontTab));
            Total.Colspan = 5;
            Total.HorizontalAlignment = 1;
            tableau1.AddCell(Total);

            //ajouter tableau
            doc.Add(tableau1);

            /******************/
            //2éme page 
            doc.NewPage();

            //2éme tableau
            PdfPTable tableauTotal = new PdfPTable(1);
            tableauTotal.WidthPercentage = 20;
            tableauTotal.HorizontalAlignment = 3;

            AddCelluleToTab("TOTAL", FontTab, tableauTotal,isTitle : true);

            //Add Cells to total table
            float total = 0f;
            for (int i = 0; i < employees.Count; i++)
            {
                total += employees[i].TotalSalary;
                AddCelluleToTab(employees[i].TotalSalary.ToString("#,0.000", nfi), FontSalairy, tableauTotal);
            }

            AddCelluleToTab(total.ToString("#,0.000", nfi), FontTab, tableauTotal);
            doc.Add(tableauTotal);

            //Close document
            doc.Close();

            OpenFile(Environment.CurrentDirectory, "GeneratePDF.pdf");
            OpenFile(Environment.CurrentDirectory, "Cnss.txt");
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

                    File.WriteAllText(myPath, GetStringFromEmployees(resultat.Tables[0]));
                    GeneretePDF(resultat.Tables[0]);

                }
            }
        }

        #endregion
    }
}
