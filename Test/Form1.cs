using ExcelDataReader;
using IronPdf;
using System;
using System.Collections.Generic;
using System.Data;
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

            employeeString += string.Format("{0:0000000000}", this.TotalSalary);

            return employeeString;
        }

        //public string [] GetEmployeePDF()
        //{   
            
        //    string employeeString = string.Format("{0:000}", index);

        //    string EmployeeName = this.Name + "" + this.Surname;

        //    employeeString += string.Format("{0,-30}", EmployeeName);

        //    employeeString += string.Format("{0:0000000000}", this.TotalSalary);

        //    return employeeString;
        //}
    }

    public partial class Form1 : Form
    {

        string myPath = @"Cnss.txt";


        public List<Employee> employees = new List<Employee>();


        public void CreateEmployees(DataTable table)
        {
            employees.Clear();

            Employee FirstEmployee = new Employee();

            FirstEmployee.Name = table.Columns[0].ColumnName;
            FirstEmployee.Surname = table.Columns[1].ColumnName;

            FirstEmployee.M1 = float.Parse(table.Columns[2].ColumnName) * 1000;
            FirstEmployee.M2 = float.Parse(table.Columns[3].ColumnName) * 1000;
            FirstEmployee.M3 = float.Parse(table.Columns[4].ColumnName) * 1000;


            employees.Add(FirstEmployee);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                Employee employee = new Employee();

                employee.Name = table.Rows[i][table.Columns[0]].ToString();
                employee.Surname = table.Rows[i][table.Columns[1]].ToString();

                employee.M1 = float.Parse((table.Rows[i][table.Columns[2]].ToString())) * 1000;
                employee.M2 = float.Parse((table.Rows[i][table.Columns[3]].ToString())) * 1000;
                employee.M3 = float.Parse((table.Rows[i][table.Columns[4]].ToString())) * 1000;



                employees.Add(employee);
            }
        }

        public string Format(DataTable table)
        {

            #region old traitment
            //string FileText = string.Empty;

            //int index = 1;

            //FileText += string.Format("{0:000}", index);
            //FileText += "" + '\t';

            //DataRow row1 = table.NewRow();

            //foreach (DataColumn col in table.Columns)
            //{

            //    FileText += string.Format("{0,-14}", col.ColumnName + "" + '\t');

            //    //Console.Write("{0,-14}", col.ColumnName);
            //}
            ////Console.WriteLine();

            //FileText += "" + '\n';



            //for (int i = 0; i < table.Rows.Count; i++)
            //{
            //    index++;

            //    FileText += string.Format("{0:000}", index);
            //    FileText += "" + '\t';

            //    for (int j = 0; j < table.Columns.Count; j++)
            //    {
            //        if (table.Columns[j].DataType.Equals(typeof(DateTime)))
            //        {
            //            // Console.Write("{0,-14:d}", table.Rows[i][table.Columns[j]]);
            //        }

            //        else if (table.Columns[j].DataType.Equals(typeof(Decimal)))
            //        {
            //            //Console.Write("{0,-14:C}", table.Rows[i][table.Columns[j]]);
            //        }

            //        else
            //            // Console.Write("{0,-14}",table.Rows[i][table.Columns[j]]);


            //            FileText += string.Format("{0,-14}", table.Rows[i][table.Columns[j]] + "" + '\t');
            //    }
            //    // Console.WriteLine();
            //    FileText += "" + '\n';
            //}

            //Console.WriteLine(FileText);

            //return FileText;


            ////foreach (DataRow row in table.Rows)
            ////{
            ////    foreach (DataColumn col in table.Columns)
            ////    {

            ////    }
            ////    Console.WriteLine();
            ////}
            ////Console.WriteLine();
            #endregion

            CreateEmployees(table);

            string FileText = string.Empty;

            for (int i = 0; i < employees.Count; i++)
            {
                FileText += employees[i].GetEmployeeText(i + 1) + '\n';
            }
            Console.WriteLine(FileText);

            return FileText;
        }


        public Form1()
        {
            InitializeComponent();

        }


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


        private void generation_Click(object sender, EventArgs e)
        {

            using (var stream = File.Open(pathFichier.Text, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet resultat = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {

                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    File.WriteAllText(myPath, Format(resultat.Tables[0]));
                }
            }



            var Renderer = new ChromePdfRenderer();
            Renderer.RenderHtmlAsPdf("<h1>This is the Tutorial for C# Generate PDF<h1>").SaveAs("GeneratePDF.pdf");

        }


    }
}
