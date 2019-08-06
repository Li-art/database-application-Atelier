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

namespace kurs1
{
    public partial class Model : Form
    {
        public bool v = true;
        int max, id_mod, id_tk;
        public string Modeli, Tkani;
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        EditModel EditModel = new EditModel();
        EditRecom EditRecom = new EditRecom();

        public Model()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EditModel.v = true;
            EditModel.id_mod = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            EditModel.max = max;
            EditModel.con = con;
            EditModel.ShowDialog();
            if (EditModel.ed)
            {
                max = EditModel.max;
                dataGridView1.Rows.Add((max).ToString(), EditModel.name_mod, EditModel.stoim_tr, EditModel.rashod_tk);
                button2.Enabled = true; button3.Enabled = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[3];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EditModel.v = false;
            EditModel.id_mod = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            EditModel.name_mod = (dataGridView1.CurrentRow.Cells[1].Value.ToString());
            EditModel.stoim_tr = Decimal.Parse(dataGridView1.CurrentRow.Cells[2].Value.ToString());
            EditModel.rashod_tk = Decimal.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString());


            EditModel.con = con;
            EditModel.ShowDialog();
            if (EditModel.ed)
            {
                dataGridView1.CurrentRow.Cells[0].Value = EditModel.id_mod;
                dataGridView1.CurrentRow.Cells[1].Value = EditModel.name_mod;
                dataGridView1.CurrentRow.Cells[2].Value = EditModel.stoim_tr;
                dataGridView1.CurrentRow.Cells[3].Value = EditModel.rashod_tk;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            com.Parameters.Clear();
            com.CommandText = "select count(id_mod) from Zakaz where id_mod=@id_mod";
            com.Parameters.Add("@id_mod", SqlDbType.Int);
            com.Parameters["@id_mod"].Value = dataGridView1.CurrentRow.Cells[0].Value.ToString(); con.Open();
            int codn = int.Parse(com.ExecuteScalar().ToString());
            if (codn > 0)
            {
                MessageBox.Show("Удаление не возможно. Есть записи модели в заказах");
                con.Close(); Close(); return;
            }
            com.Parameters.Clear(); con.Close();

            com.Parameters.Clear();
            com.CommandText = "select count(id_mod) from Rekom where id_mod=@id_mod";
            com.Parameters.Add("@id_mod", SqlDbType.Int);
            com.Parameters["@id_mod"].Value = dataGridView1.CurrentRow.Cells[0].Value.ToString(); con.Open();
            int codm = int.Parse(com.ExecuteScalar().ToString());
            if (codm > 0)
            {
                MessageBox.Show("Удаление не возможно. Есть записи модели в рекомендуемых");
                con.Close(); Close(); return;
            }


            com.Parameters.Clear(); con.Close();
            DialogResult res = MessageBox.Show("Удалить?", "Внимание", MessageBoxButtons.YesNo);
            if (DialogResult.Yes == res)
            {
                com.CommandText = "delete from Model where id_mod=@id_mod";
                com.Parameters.Add("@id_mod", SqlDbType.Int);
                com.Parameters["@id_mod"].Value = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                try { con.Open(); }
                catch { MessageBox.Show("Нет соединения"); Close(); return; }
                try { com.ExecuteNonQuery(); }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                con.Close(); dataGridView1.Rows.Remove(dataGridView1.CurrentRow); com.Parameters.Clear();
                if (dataGridView1.Rows.Count == 0) { button2.Enabled = false; button3.Enabled = false; }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            EditRecom.v = true;
            //EditRecom.id_mod = int.Parse(dataGridView2.CurrentRow.Cells[0].Value.ToString());
            EditRecom.max = max;
            EditRecom.con = con;
            EditRecom.ShowDialog();
            /*if (EditModel.ed)
            {
                max = EditRecom.max;
                dataGridView2.Rows.Add((max).ToString(), EditRecom.name_mod);
                //button2.Enabled = true; button3.Enabled = true;
                //dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2];
            }*/
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
            DialogResult res = MessageBox.Show("Удалить?", "Внимание", MessageBoxButtons.YesNo);
            if (DialogResult.Yes == res)
            {
                com.CommandText = "delete from Rekom where id_mod=@id_mod and id_tk=@id_tk";
                /*com.Parameters.Add("@id_mod", SqlDbType.Int);
                com.Parameters["@id_mod"].Value = comboBox1.SelectedItem();*/
                com.Parameters.Add("@id_tk", SqlDbType.Int);
                com.Parameters["@id_tk"].Value = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                try { con.Open(); }
                catch { MessageBox.Show("Нет соединения"); Close(); return; }
                try { com.ExecuteNonQuery(); }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                con.Close(); dataGridView1.Rows.Remove(dataGridView1.CurrentRow); com.Parameters.Clear();
                if (dataGridView1.Rows.Count == 0) { button2.Enabled = false; button3.Enabled = false; }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            

                dataGridView2.Rows.Clear();
                con.ConnectionString = Properties.Settings.Default.con;
                com.Connection = con;

                com.CommandType = CommandType.Text;
                com.CommandText = "select Tkani.id_tk, Tkani.name_tk from Tkani join Rekom on Tkani.id_tk=Rekom.id_tk join Model on Rekom.id_mod=Model.id_mod where  Model.name_mod=@name_mod";
                com.Parameters.Add("@name_mod", SqlDbType.VarChar);
                com.Parameters["@name_mod"].Value = comboBox1.SelectedItem;
                try { con.Open(); }
                catch
                {
                    MessageBox.Show("Нет соединения");
                    Close(); return;
                }
                try
                {
                    SqlDataReader rd = com.ExecuteReader();
                    while (rd.Read())
                    {

                            dataGridView2.Rows.Add(rd[0].ToString(), rd[1].ToString());
                     
                    }
                    rd.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            com.Parameters.Clear();
            con.Close();
           

        }

        private void Model_Load(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView1.Rows.Clear();
            // TODO: данная строка кода позволяет загрузить данные в таблицу "atelieDataSet.Model". При необходимости она может быть перемещена или удалена.
            this.modelTableAdapter.Fill(this.atelieDataSet.Model);
            con.ConnectionString = Properties.Settings.Default.con;
            com.Connection = con;
            com.CommandType = CommandType.Text;
            
            com.CommandText = "select * from Model";
            
            try { con.Open(); }
            catch
            {
                MessageBox.Show("Нет соединения");
                Close(); return;
            }
            try
            {
                SqlDataReader rd = com.ExecuteReader();
                while (rd.Read()) dataGridView1.Rows.Add(rd[0].ToString(), rd[1].ToString(), rd[2].ToString(), rd[3].ToString());
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            con.Close();
           
        }
    }
}
