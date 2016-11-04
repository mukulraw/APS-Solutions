using System;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace APS_Solutions
{



    public partial class Form1 : Form
    {
        private PictureBox p1;
        private int count = 1;

        //Excel.Application excelApp = new Excel.Application();

        //Form1 program = new Form1();


        DataTable clients;





        public Form1()
        {
            InitializeComponent();
            p1 = new PictureBox();
            setUpTable();
        }





        

        private void button1_Click(object sender, EventArgs e)
        {
            

            flowLayoutPanel1.Controls.Remove(p1);

            
            p1.Height = flowLayoutPanel1.Height;
            p1.Width = flowLayoutPanel1.Width;
            p1.SizeMode = PictureBoxSizeMode.AutoSize;
            
            p1.BackColor = Color.Black;
            flowLayoutPanel1.Controls.Add(p1);

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.CheckFileExists = true;
            openDialog.ShowReadOnly = false;
            openDialog.Title = "Browse Image Files";
            openDialog.RestoreDirectory = true;
            openDialog.DefaultExt = "png";
            openDialog.Filter = "Images (*.BMP;*.JPG;*.GIF,*.PNG,*.TIFF)|*.BMP;*.JPG;*.GIF;*.PNG;*.TIFF|" + "All files (*.*)|*.*";
            openDialog.FilterIndex = 2;
            

            


            if(openDialog.ShowDialog() == DialogResult.OK)
            {
                label1.Text = openDialog.FileName;
                
                p1.Image = Image.FromFile(openDialog.FileName);
            }


        }



        private void setUpTable()
        {
            clients = new DataTable("Clients");

            clients.Columns.Add("Name");
            clients.Columns.Add("Email");
            clients.Columns.Add("Code");

        }



        private void ExportDataSetToExcel(DataSet ds)
        {
            //Creae an Excel application instance
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            if(excelApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            excelApp.Visible = true;

            //Create an Excel workbook instance and open it from the predefined location
            //Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(@"E:\"+ textBox15.Text +".xlsx");

            Excel.Workbook workBooh1 = excelApp.Workbooks.Add(Missing.Value);

            //Excel.Worksheet ws = (Excel.Worksheet)workBooh1.Worksheets[1];


            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = (Excel.Worksheet)workBooh1.ActiveSheet;
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            workBooh1.SaveAs(@"C:\"+ textBox15.Text +".xls" , Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            //
            workBooh1.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(workBooh1);


            //workBooh1.Close();

        }





        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            DataSet ds = new DataSet("Reports");
            ds.Tables.Add(clients);

            ExportDataSetToExcel(ds);

            setUpTable();

            textBox15.Clear();






        }

        private void button3_Click(object sender, EventArgs e)
        {
            clients.Rows.Add(textBox1.Text , textBox2.Text , textBox3.Text);
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
        }
    }
}
