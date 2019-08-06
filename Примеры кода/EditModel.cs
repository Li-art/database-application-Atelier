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
    public partial class EditModel : Form
    {
        public bool v, ed;
        public int id_mod, max;
        public string name_mod;
        public decimal stoim_tr, rashod_tk;

        public SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();

        public EditModel()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "" || textBox3.Text.Trim() == "")
            {
                MessageBox.Show("Не все поля заполненны!"); textBox1.Focus(); return;
            }
            decimal tel1;
            if (!Decimal.TryParse(textBox2.Text, out tel1))
            { MessageBox.Show("Стоимость введена не правильно"); textBox2.Focus(); return; }

            decimal tel2;
            if (!Decimal.TryParse(textBox3.Text, out tel2))
            { MessageBox.Show("Расход введен не правильно"); textBox3.Focus(); return; }


            stoim_tr = Decimal.Parse(textBox2.Text);
            rashod_tk = Decimal.Parse(textBox3.Text);

            com.Parameters.Clear();
            com.Connection = con;
            if (v)
            {
                com.Parameters.Clear(); com.CommandText = "insert into Model values (@name_mod,@stoim_tr,@rashod_tk)";

                //com.Parameters.Add("@id_tk", SqlDbType.Int);
                com.Parameters.Add("@name_mod", SqlDbType.VarChar);
                com.Parameters.Add("@stoim_tr", SqlDbType.Decimal);
                com.Parameters.Add("@rashod_tk", SqlDbType.Decimal);

                //com.Parameters["@id_tk"].Value = max + 1;

                com.Parameters["@name_mod"].Value = textBox1.Text;
                com.Parameters["@stoim_tr"].Value = textBox2.Text;
                com.Parameters["@rashod_tk"].Value = textBox3.Text;

                try { con.Open(); }
                catch { MessageBox.Show("Нет соединения"); Close(); return; }
                try { com.ExecuteNonQuery(); ed = true; }
                catch (Exception ex)
                { MessageBox.Show(ex.Message); ed = false; Close(); con.Close(); return; }
                if (max == -1) max = 1;
                else max += 1;
                name_mod = textBox1.Text;
                //width = textBox2.Text;
                //cena = textBox3.Text;
                stoim_tr = Decimal.Parse(textBox2.Text);
                rashod_tk = Decimal.Parse(textBox3.Text);


                ed = true;
                com.Parameters.Clear();
                con.Close();
                Close();
            }
            else
            {
                com.Parameters.Clear();
                com.CommandText = "update Model set name_mod=@name_mod,stoim_tr=@stoim_tr,rashod_tk=@rashod_tk  where  id_mod=@id_mod";

                com.Parameters.Add("@id_mod", SqlDbType.Int);

                com.Parameters.Add("@name_mod", SqlDbType.VarChar);
                com.Parameters.Add("@stoim_tr", SqlDbType.Decimal);
                com.Parameters.Add("@rashod_tk", SqlDbType.Decimal);


                com.Parameters["@id_mod"].Value = id_mod;

                com.Parameters["@name_mod"].Value = textBox1.Text;
                com.Parameters["@stoim_tr"].Value = Decimal.Parse(textBox2.Text);
                com.Parameters["@rashod_tk"].Value = Decimal.Parse(textBox3.Text);

                con.Open();
                try { com.ExecuteNonQuery(); }
                catch (Exception ex) { MessageBox.Show(ex.Message); }

                name_mod = textBox1.Text;
                stoim_tr = Decimal.Parse(textBox2.Text);
                rashod_tk = Decimal.Parse(textBox3.Text);

                ed = true;
                com.Parameters.Clear();
                con.Close();
                Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ed = false; Close();
        }

        private void EditModel_Load(object sender, EventArgs e)
        {
            con.ConnectionString = Properties.Settings.Default.con;
            com.Connection = con;
            com.CommandType = CommandType.Text;
            if (v) { textBox1.Text = ""; textBox2.Text = ""; textBox3.Text = ""; }
            else
            {

                textBox1.Text = name_mod.ToString(); textBox2.Text = stoim_tr.ToString(); textBox3.Text = rashod_tk.ToString();

            }
        }
    }
}
