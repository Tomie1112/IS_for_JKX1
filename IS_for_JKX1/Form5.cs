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
    public partial class Form5 : Form
    {
        public static SqlConnection sqlConnection = null;
        public Form5(string kod)
        {
            InitializeComponent();
        }

        private async void Form5_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet4.Повторное_открытие_заявок". При необходимости она может быть перемещена или удалена.
            this.повторное_открытие_заявокTableAdapter.Fill(this.u1666130_JKH34DataSet4.Повторное_открытие_заявок);
            string connectionString = @"Data Source = 31.31.198.141; Initial Catalog = u1666130_JKH34; User ID = u1666130_Yuliya; Password = csb#G254";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
        }

        private async void button7_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text != "")
            {
                string kod_z = Convert.ToString(textBox1.Text);
                string new_status = "В процессе";
                SqlCommand command = new SqlCommand("UPDATE [Заявки] SET [Статус]=@Статус WHERE [Номер_заявки]=@Номер_заявки", sqlConnection);
                command.Parameters.AddWithValue("Статус", new_status);
                command.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(kod_z));
                await command.ExecuteNonQueryAsync();

                int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
                int kod_otkr = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
                SqlCommand command2 = new SqlCommand("INSERT INTO [dbo].[Повторное_открытие_заявок] (Номер_повторного_открытия_заявки, Номер_заявки," +
                            " Причина_повторного_открытия_заявки)VALUES(@Номер_повторного_открытия_заявки,@Номер_заявки,@Причина_повторного_открытия_заявки)", sqlConnection);
                command2.Parameters.AddWithValue("Номер_повторного_открытия_заявки", Convert.ToInt32(kod_otkr) + 1);
                command2.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(kod_z));
                command2.Parameters.AddWithValue("Причина_повторного_открытия_заявки", Convert.ToString(richTextBox1.Text));
                await command2.ExecuteNonQueryAsync();
                MessageBox.Show("Заявка успешно открыта!");
                this.Close();
            }
            else MessageBox.Show("Сначала заполните причину открытия!");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
