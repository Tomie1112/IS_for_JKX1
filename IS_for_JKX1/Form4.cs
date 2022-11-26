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

namespace IS_for_JKX1
{
    public partial class Form4 : Form
    {
        public static SqlConnection sqlConnection = null;
        public Form4(string kod_zayavki)
        {
            InitializeComponent();
        }

        private async void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) &&
            !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(comboBox1.Text) &&
            !string.IsNullOrEmpty(comboBox2.Text) && !string.IsNullOrEmpty(richTextBox1.Text))
            {
                //update
                SqlCommand command = new SqlCommand("UPDATE [Заявки] SET [Дата_оформления]=@Дата_оформления,[Заявитель]=@Заявитель," +
                "[Контактный_телефон]=@Контактный_телефон,[Электронная_почта]=@Электронная_почта,[Адрес]=@Адрес,[Код_организации]=@Код_организации," +
                "[Тип_заявки]=@Тип_заявки,[Жалоба]=@Жалоба, [Результат_заявки]=@Результат_заявки " +
                "WHERE [Номер_заявки]=" + Convert.ToInt32(textBox5.Text), sqlConnection);
                command.Parameters.AddWithValue("Дата_оформления", dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("Заявитель", Convert.ToString(textBox1.Text));
                command.Parameters.AddWithValue("Контактный_телефон", Convert.ToString(textBox3.Text));
                command.Parameters.AddWithValue("Электронная_почта", Convert.ToString(textBox4.Text));
                command.Parameters.AddWithValue("Адрес", Convert.ToString(textBox2.Text));
                command.Parameters.AddWithValue("Код_организации", Convert.ToInt32(comboBox3.Items[comboBox1.SelectedIndex]));
                command.Parameters.AddWithValue("Тип_заявки", Convert.ToString(comboBox2.Text));
                command.Parameters.AddWithValue("Жалоба", Convert.ToString(richTextBox1.Text));

                command.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox2.Text));
                await command.ExecuteNonQueryAsync();
                MessageBox.Show("Успешно изменено!");
            }
            else
                MessageBox.Show("Заполните все значения!");
        }

        private async void Form4_Load(object sender, EventArgs e)
        {
            string connectionString = @"Data Source = 31.31.198.141; Initial Catalog = u1666130_JKH34; User ID = u1666130_Yuliya; Password = csb#G254";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            SqlDataReader sqlReader2 = null;
            SqlCommand command2 = new SqlCommand("SELECT * FROM [Области_заявок]", sqlConnection);

            try
            {
                sqlReader2 = await command2.ExecuteReaderAsync();

                while (await sqlReader2.ReadAsync())
                {
                    comboBox2.Items.Add(Convert.ToString(sqlReader2["Тип_заявки"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader2 != null)
                    sqlReader2.Close();
            }



            SqlDataReader sqlReader5 = null;
            SqlCommand command25 = new SqlCommand("SELECT * FROM [Организации]", sqlConnection);

            try
            {
                sqlReader5 = await command25.ExecuteReaderAsync();

                while (await sqlReader5.ReadAsync())
                {
                    string proverka_org= Convert.ToString(sqlReader5["Статус"]);
                    if(proverka_org=="Действующий")
                    { 
                    comboBox1.Items.Add(Convert.ToString(sqlReader5["Название"]));
                    comboBox3.Items.Add(Convert.ToString(sqlReader5["Код_организации"]));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader5 != null)
                    sqlReader5.Close();
            }




            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Заявки] WHERE Номер_заявки=" + Convert.ToInt32(textBox5.Text), sqlConnection);
            string proverka = "";
            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = (Convert.ToString(sqlReader["Заявитель"]));

                    proverka = Convert.ToString(sqlReader["Контактный_телефон"]);
                    if (proverka != "") textBox3.Text = (Convert.ToString(sqlReader["Контактный_телефон"]));
                    //else textBox3.Text = "123";
                    proverka = "";

                    proverka = Convert.ToString(sqlReader["Электронная_почта"]);
                    if (proverka != "") textBox4.Text = (Convert.ToString(sqlReader["Электронная_почта"]));
                    proverka = "";

                    dateTimePicker1.Value = Convert.ToDateTime(sqlReader["Дата_оформления"]);
                    textBox2.Text = Convert.ToString(sqlReader["Адрес"]);
                    comboBox3.Text = (Convert.ToString(sqlReader["Код_организации"]));
                    int index = comboBox3.SelectedIndex;
                    comboBox1.Text= comboBox1.Items[index].ToString();
                    comboBox2.Text = (Convert.ToString(sqlReader["Тип_заявки"]));
                    richTextBox1.Text = (Convert.ToString(sqlReader["Жалоба"]));
                    proverka = (Convert.ToString(sqlReader["Результат_заявки"]));
                    if (proverka != "")
                        richTextBox2.Text = (Convert.ToString(sqlReader["Результат_заявки"]));
                    proverka = "";
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
    }
}
