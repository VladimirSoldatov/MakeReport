using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Windows.Forms;

namespace MakeReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();



        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            connectionBD();
       

        }
        void dataGridInsert()
        {
            try
            {
                dataGridView1.AllowUserToAddRows = false;
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                ExcelWorkBook = ExcelApp.Workbooks.Add();
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;
                int colCount = dataGridView1.ColumnCount;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }

                }
                saveFileDialog1.FileName = "Отчет ТО.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                saveFileDialog1.Filter = saveFileDialog1.Filter = "Excel files(*.xlsx)|*.xls";
                string filename = saveFileDialog1.FileName;
                ExcelWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ExcelApp.Quit();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        void connectionBD()
        {
            SqlCommand sqlCommand = new SqlCommand();
            DataSet dataSet = new DataSet();
            using (SqlConnection sqlConnection = new SqlConnection())
            {
                try
                {
                    sqlConnection.ConnectionString = $"Server={textBox3.Text};Database={textBox4.Text};User Id={textBox1.Text};Password={textBox2.Text};";
                    sqlConnection.Open();
                    MessageBox.Show("Connection open");
                    sqlCommand.Connection = sqlConnection;
                    sqlCommand.CommandText = "select  client.descr as 'ФИО клиента'," +
                        " FORMAT(last_TO_date, 'dd.MM.yyyy') as 'Дата последнего ТО', " +
                        " cars.vin_id as 'VIN', " +
                        " client.phone_kontakt as 'Телефон клиента', " +
                        " CASE WHEN truster.descr IS NULL THEN 'Нет данных' ELSE truster.descr END as 'ФИО ДЛ'," +
                        " CASE WHEN truster.telefon IS NULL THEN 'Нет данных' ELSE truster.telefon END as 'Телефон ДЛ' " +
                        " from tcavt001 as cars " +
                        " Left JOIN tccom010 as client ON cars.client_id = client.contragent_id " +
                        " Left JOIN tccom004 as truster on cars.client_id = truster.contragent_id " +
                        " where last_TO_date between DATEADD(YEAR, -2,GETDATE()) and DATEADD(YEAR, -1,GETDATE()) " +
                        " GROUP BY cars.vin_id, last_TO_date,client.descr,client.phone_kontakt ,truster.descr  , truster.telefon";
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    while (reader.Read())
                    {
                        dataGridView1.Rows.Add(reader.GetSqlString(0), reader.GetSqlString(1), reader.GetSqlString(2), reader.GetSqlString(3), reader.GetSqlString(4), reader.GetSqlString(5));
                    }

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                finally
                {

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            dataGridInsert();
        }
    }
}
