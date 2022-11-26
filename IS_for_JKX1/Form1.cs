using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace IS_for_JKX1
{
    public partial class Form1 : Form
    {
        public static SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
            pictureBox4.Visible = false;
        }
        bool proverka = true;
        bool proverka2 = true;//textbox1
        bool proverka3 = true;//textbox2
        public int i = -1;
        private async void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "Логин";
            textBox1.ForeColor = Color.Gray;
            textBox2.Text = "Пароль";
            textBox2.ForeColor = Color.Gray;
            string connectionString = @"Data Source = 31.31.198.141; Initial Catalog = u1666130_JKH34; User ID = u1666130_Yuliya; Password = csb#G254";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;
            SqlCommand command1 = new SqlCommand("SELECT * FROM [Сотрудники_предприятия]", sqlConnection);

            try
            {
                sqlReader = await command1.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {

                    i = i + 1;
                    comboBox1.Items.Add(Convert.ToString(sqlReader["Логин"]));
                    comboBox2.Items.Add(Convert.ToString(sqlReader["Пароль"]));
                    comboBox3.Items.Add(Convert.ToString(sqlReader["Код_сотрудника"]));
                    comboBox4.Items.Add(Convert.ToString(sqlReader["Статус"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

        }
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            if (proverka == false)
            {
                if (proverka3 == false)
                {
                    textBox2.UseSystemPasswordChar = true;
                }
                pictureBox3.Visible = true;
                pictureBox4.Visible = false;
                proverka = true;
            }

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (proverka == true)
            {
                if (proverka3 == false)
                {
                    textBox2.UseSystemPasswordChar = false;
                }
                pictureBox3.Visible = false;
                pictureBox4.Visible = true;
                proverka = false;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int ch = i; ch > -1; ch--)
            {
                string comb = Convert.ToString(comboBox1.Items[ch]);
                string text = Convert.ToString(textBox1.Text);
                if (comb == text)
                {
                    string comb2 = Convert.ToString(comboBox2.Items[ch]);
                    string text2 = Convert.ToString(textBox2.Text);
                    if (comb2 == text2)
                    {
                        string comb4 = Convert.ToString(comboBox4.Items[ch]);
                        if (comb4 != "Уволен")
                        {
                            string kod = Convert.ToString(comboBox3.Items[ch]);

                            MessageBox.Show("Добро пожаловать!");

                            Form2 form = new Form2(kod);
                            this.Hide();
                            form.textBox1.Text = "" + Convert.ToString(kod);
                            form.ShowDialog();
                            ch = -1;
                        }
                        else 
                        { 
                            MessageBox.Show("У вас больше нет доступа к системе!");
                            ch = -1;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Неправильный пароль!");
                        ch = -1;
                    }
                }
                if (ch == 0 && comb != text)
                {
                    MessageBox.Show("Неправильный логин!");
                }

            }
        }


        private void textBox1_Click(object sender, EventArgs e)
        {
            if (proverka2 == true)
            {
                textBox1.Clear();
                proverka2 = false;
            }
            textBox1.ForeColor = Color.Black;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (proverka3 == true)
            {
                textBox2.Clear();
                proverka3 = false;
            }

            textBox2.ForeColor = Color.Black;
            if (proverka == true)
                textBox2.UseSystemPasswordChar = true;

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (proverka3 == true)
            {
                textBox2.Clear();
                proverka3 = false;
            }
            if (proverka == true)
            {
                textBox2.UseSystemPasswordChar = true;
                textBox2.ForeColor = Color.Black;
            }

            if (e.KeyCode == Keys.Enter)
            {
                for (int ch = i; ch > -1; ch--)
                {
                    string comb = Convert.ToString(comboBox1.Items[ch]);
                    string text = Convert.ToString(textBox1.Text);
                    if (comb == text)
                    {
                        string comb2 = Convert.ToString(comboBox2.Items[ch]);
                        string text2 = Convert.ToString(textBox2.Text);
                        if (comb2 == text2)
                        {
                            string comb4 = Convert.ToString(comboBox4.Items[ch]);
                            if (comb4 != "Уволен")
                            {
                                string kod = Convert.ToString(comboBox3.Items[ch]);

                                MessageBox.Show("Добро пожаловать!");

                                Form2 form = new Form2(kod);
                                this.Hide();
                                form.textBox1.Text = "" + Convert.ToString(kod);
                                form.ShowDialog();
                                ch = -1;
                            }
                            else
                            {
                                MessageBox.Show("У вас больше нет доступа к системе!");
                                ch = -1;
                            }
                        }
                        if (comb2 != text2)
                        {
                            MessageBox.Show("Неправильный пароль!");
                            ch = -1;
                        }
                    }
                    if (ch == 0 && comb != text)
                    {
                        MessageBox.Show("Неправильный логин!");
                    }

                }
            }

        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (proverka2 == true)
            {
                textBox1.Clear();
                proverka2 = false;
            }
            textBox1.ForeColor = Color.Black;
                if (e.KeyCode == Keys.Enter)
            {

                for (int ch = i; ch > -1; ch--)
                {
                    string comb = Convert.ToString(comboBox1.Items[ch]);
                    string text = Convert.ToString(textBox1.Text);
                    if (comb == text)
                    {
                        string comb2 = Convert.ToString(comboBox2.Items[ch]);
                        string text2 = Convert.ToString(textBox2.Text);
                        if (comb2 == text2)
                        {
                            string comb4 = Convert.ToString(comboBox4.Items[ch]);
                            if (comb4!="Уволен")
                            {                             
                            string kod = Convert.ToString(comboBox3.Items[ch]);

                            MessageBox.Show("Добро пожаловать!");

                            Form2 form = new Form2(kod);
                            this.Hide();
                            form.textBox1.Text = "" + Convert.ToString(kod);
                            form.ShowDialog();
                            ch = -1;
                            }
                            else
                            {
                                MessageBox.Show("У вас больше нет доступа к системе!");
                                ch = -1;
                            }
                        }
                        if (comb2 != text2)
                        {
                            MessageBox.Show("Неправильный пароль!");
                            ch = -1;
                        }
                    }
                    if (ch == 0 && comb != text)
                    {
                        MessageBox.Show("Неправильный логин!");
                    }

                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
