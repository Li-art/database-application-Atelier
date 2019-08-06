using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace kurs1
{
    public partial class Otch2 : Form
    {

        public int id_mod, id_z;

        public string name_mod, fio_zakr;

        public SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();

        public Otch2()
        {
            InitializeComponent();
        }

        private void Otch2_Load(object sender, EventArgs e)
        {
            dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker2.MaxDate = DateTime.Now;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            con.ConnectionString = Properties.Settings.Default.con;
            com.Connection = con;

            try { con.Open(); }
            catch
            {
                MessageBox.Show("Нет соединения");
                Close(); return;
            }
            com.CommandText = "select  distinct(fio_zakr) from Zakroyshik join Zakaz on Zakaz.id_zakr=Zakroyshik.id_zakr join Model on Model.id_mod=Zakaz.id_mod";
            try
            {
                SqlDataReader rd1 = com.ExecuteReader();
                while (rd1.Read()) comboBox1.Items.Add(rd1[0].ToString());
                rd1.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

            // comboBox1.SelectedIndex = 1;

            con.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count != 0)
            {
                if (radioButton1.Checked == false)
                {
                    comboBox1.Enabled = true;
                   // radioButton1.Enabled = false;
                    dataGridView1.Rows.Clear();
                    com.Parameters.Add("@fio_zakr", SqlDbType.VarChar);
                    if (comboBox1.Text != "")
                    {
                        //com.CommandText = "select Model, Brand.Id_brand from Model join Brand on Brand.Id_brand=Model.Id_brand join Car on Model.Id_model=Car.Id_model where Brand.Brand = @Brand";                   
                        com.Parameters["@fio_zakr"].Value = comboBox1.SelectedItem;
                        fio_zakr = comboBox1.SelectedItem.ToString();
                    }
                    else
                    {
                        
                        com.Parameters["@fio_zakr"].Value = " ";
                        fio_zakr = " ";
                        
                    }
                    
                    try { con.Open(); }
                    catch { MessageBox.Show("нет соединения"); Close(); return; }

                    dataGridView1.Rows.Clear();

                    com.CommandType = CommandType.Text;
                    com.CommandText = "select distinct(Zakaz.id_mod), name_mod, count(name_mod), sum(sum) from Zakaz join Model on Model.id_mod=Zakaz.id_mod join Zakroyshik on Zakroyshik.id_zakr=Zakaz.id_zakr where fio_zakr=@fio_zakr and data_pr>=@Since and data_prim<=@To GROUP BY Zakaz.id_mod, name_mod";
                    // com.CommandText = "select Car.Id_car,Id_tech,Country,Id_model,Colour,Date_of_release,Price, Number_of_cars from Car join Storage on Storage.Id_car=Car.Id_car";
                    com.Parameters.Add("@Since", SqlDbType.Date);
                    com.Parameters["@Since"].Value = dateTimePicker1.Value;
                    com.Parameters.Add("@To", SqlDbType.Date);
                    com.Parameters["@To"].Value = dateTimePicker2.Value;

                    try
                    {
                        SqlDataReader rd2 = com.ExecuteReader();
                        while (rd2.Read())
                        {

                            dataGridView1.Rows.Add(rd2[0].ToString(), rd2[1].ToString(), rd2[2].ToString(), rd2[3].ToString());

                        }
                        rd2.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }

                    com.CommandType = CommandType.Text;
                    com.CommandText = "select count(Zakaz.id_mod), SUM(sum) from Zakaz join Model on Model.id_mod=Zakaz.id_mod join Zakroyshik on Zakroyshik.id_zakr=Zakaz.id_zakr where fio_zakr=@fio_zakr and data_pr>=@Since and data_prim<=@To";
                    try
                    {
                        SqlDataReader rd = com.ExecuteReader();
                        while (rd.Read())
                        {
                            dataGridView1.Rows.Add("", "Итог:", rd[0].ToString(), rd[1].ToString());

                        }
                        rd.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }



                    dataGridView2.Rows.Clear();

                    com.CommandType = CommandType.Text;
                    com.CommandText = "select distinct(Zakaz.id_zakr), fio_zakr, count(name_mod), sum(sum) from Zakaz join Model on Model.id_mod=Zakaz.id_mod join Zakroyshik on Zakroyshik.id_zakr=Zakaz.id_zakr where data_pr>=@Since and data_prim<=@To GROUP BY Zakaz.id_zakr, fio_zakr";
                    /*com.Parameters.Add("@Since", SqlDbType.Date);
                    com.Parameters["@Since"].Value = dateTimePicker1.Value;
                    com.Parameters.Add("@To", SqlDbType.Date);
                    com.Parameters["@To"].Value = dateTimePicker2.Value;*/

                    try
                    {
                        SqlDataReader rd3 = com.ExecuteReader();
                        while (rd3.Read())
                        {

                            dataGridView2.Rows.Add(rd3[0].ToString(), rd3[1].ToString(), rd3[2].ToString(), rd3[3].ToString());

                        }
                        rd3.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }


                    com.Parameters.Clear();

                    con.Close();

                    com.Parameters.Clear();

                    con.Close();
                }
                
                else
                //if (radioButton1.Checked)
                {
                   


                    comboBox1.Enabled = false;

                    dataGridView1.Rows.Clear();
                    try { con.Open(); }
                    catch { MessageBox.Show("нет соединения"); Close(); return; }

                    com.CommandType = CommandType.Text;
                    com.CommandText = "select distinct(Zakaz.id_mod), name_mod, count(name_mod), sum(sum) from Zakaz join Model on Model.id_mod=Zakaz.id_mod join Zakroyshik on Zakroyshik.id_zakr=Zakaz.id_zakr where data_pr>=@Since and data_prim<=@To GROUP BY Zakaz.id_mod, name_mod";
                    // com.CommandText = "select Car.Id_car,Id_tech,Country,Id_model,Colour,Date_of_release,Price, Number_of_cars from Car join Storage on Storage.Id_car=Car.Id_car";
                    com.Parameters.Add("@Since", SqlDbType.Date);
                    com.Parameters["@Since"].Value = dateTimePicker1.Value;
                    com.Parameters.Add("@To", SqlDbType.Date);
                    com.Parameters["@To"].Value = dateTimePicker2.Value;

                    try
                    {
                        SqlDataReader rd2 = com.ExecuteReader();
                        while (rd2.Read())
                        {


                            dataGridView1.Rows.Add(rd2[0].ToString(), rd2[1].ToString(), rd2[2].ToString(), rd2[3].ToString());

                        }
                        rd2.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }

                    com.CommandType = CommandType.Text;
                    com.CommandText = "select count(Zakaz.id_mod), SUM(sum)from Zakaz join Model on Zakaz.id_mod=Model.id_mod where data_pr>=@Since and data_prim<=@To";
                    try
                    {
                        SqlDataReader rd = com.ExecuteReader();
                        while (rd.Read())
                        {

                            dataGridView1.Rows.Add("", "Итог:", rd[0].ToString(), rd[1].ToString());

                        }
                        rd.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }




                    com.Parameters.Clear();

                    con.Close();

                    
                }
            }
            }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            radioButton1.Enabled = true;
            radioButton2.Enabled = true;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            radioButton1.Enabled = true;
            radioButton2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          

                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
            ((Excel.Range)sheet.Columns).ColumnWidth = 30;

            sheet.Cells[2, 1] = "Модель:";
            (sheet.Cells[2, 1] as Excel.Range).Font.Bold = true;
            (sheet.Cells[2, 1] as Excel.Range).Font.Size = 14;

            sheet.Cells[2, 2] = "Количество заказов:";
            (sheet.Cells[2, 2] as Excel.Range).Font.Bold = true;
            (sheet.Cells[2, 2] as Excel.Range).Font.Size = 14;

            sheet.Cells[2, 3] = "Сумма:";
            (sheet.Cells[2, 3] as Excel.Range).Font.Bold = true;
            (sheet.Cells[2, 3] as Excel.Range).Font.Size = 14;

            


           /* sheet.Cells[dataGridView1.RowCount + 3, 5] = "Итогo:";
            (sheet.Cells[dataGridView1.RowCount + 3, 5] as Excel.Range).Font.Bold = true;
            (sheet.Cells[dataGridView1.RowCount + 3, 5] as Excel.Range).Font.Size = 14;*/




            sheet.Cells[1, 1] = "Закройщик:";
            (sheet.Cells[1, 1] as Excel.Range).Font.Bold = true;
            (sheet.Cells[1, 1] as Excel.Range).Font.Size = 14;


            sheet.Cells[1, 2] = comboBox1.Text.ToString();
            //     (sheet.Cells[1, 2] as Excel.Range).Font.Bold = true; 
            (sheet.Cells[1, 2] as Excel.Range).Font.Size = 14;



            sheet.Cells[1, 3] = "Период:";
            (sheet.Cells[1, 3] as Excel.Range).Font.Bold = true;
            (sheet.Cells[1, 3] as Excel.Range).Font.Size = 14;


            sheet.Cells[1, 4] = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            //(sheet.Cells[1, 5] as Excel.Range).Font.Bold = true;
            (sheet.Cells[1, 4] as Excel.Range).Font.Size = 14;

            sheet.Cells[1, 5] = dateTimePicker2.Value.ToString("dd/MM/yyyy");
            // (sheet.Cells[1, 7] as Excel.Range).Font.Bold = true;
            (sheet.Cells[1, 5] as Excel.Range).Font.Size = 14;


            Decimal X = 0;
            int Y = 0;
            for (int i = 1; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {


                    sheet.Cells[j + 3, i] = (dataGridView1[i, j].Value).ToString();
                    (sheet.Cells[j + 3, i] as Excel.Range).Font.Size = 14;
                    (sheet.Cells[j + 3, i] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    (sheet.Cells[2, i] as Excel.Range).Font.Size = 14;
                    (sheet.Cells[2, i] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (sheet.Cells[1, i] as Excel.Range).Font.Size = 14;
                    (sheet.Cells[1, i] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

            }

           /* for (int j = 0; j < dataGridView1.RowCount; j++)
            {
                X = X + Decimal.Parse((dataGridView1[3, j].Value).ToString());
                Y = Y + int.Parse((dataGridView1[2, j].Value).ToString());
            }

            sheet.Cells[dataGridView1.RowCount + 3, 8] = X.ToString();
            sheet.Cells[dataGridView1.RowCount + 3, 6] = Y.ToString();
            (sheet.Cells[dataGridView1.RowCount + 3, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (sheet.Cells[dataGridView1.RowCount + 3, 7] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (sheet.Cells[dataGridView1.RowCount + 3, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (sheet.Cells[dataGridView1.RowCount + 3, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // (sheet.Cells[dataGridView2.RowCount + 3, 8] as Excel.Range).Font.Bold = true;
            (sheet.Cells[dataGridView1.RowCount + 3, 8] as Excel.Range).Font.Size = 14;
            (sheet.Cells[dataGridView1.RowCount + 3, 5] as Excel.Range).Font.Size = 14;
            (sheet.Cells[dataGridView1.RowCount + 3, 6] as Excel.Range).Font.Size = 14;

            (sheet.Cells[dataGridView1.RowCount + 3, 5] as Excel.Range).Font.Bold = true;
            (sheet.Cells[dataGridView1.RowCount + 3, 6] as Excel.Range).Font.Bold = true;
            (sheet.Cells[dataGridView1.RowCount + 3, 8] as Excel.Range).Font.Bold = true;

    */

            ExcelApp.Visible = true;
        
        }
    }
}
