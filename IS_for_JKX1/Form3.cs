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
using Word = Microsoft.Office.Interop.Word;
using System.Net;
using System.Net.Mail;

namespace IS_for_JKX1
{
   
    public partial class Form3 : Form
    { 
        public static SqlConnection sqlConnection = null;
        public Form3(string kod_zayavki)
        {
            InitializeComponent();
        }
        private void ReplaceWordSrub(string stubToReplace, string text, Word.Document wordDocument)//заполнение документов
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);

        }

        private readonly string TemplateFileName = @"C:\otvet.doc";
        public string mail;
        public string data_ofr;
        public string name;
        public string adress;
        public string kodorg;
        public bool b;
        public bool proverka = false;
        public async void update()
        {
            richTextBox4.Clear();
            string number = "" + textBox1.Text;
            SqlDataReader sqlReader = null;
            SqlCommand command3 = new SqlCommand("SELECT * FROM [Заявки] WHERE Номер_заявки=" + Convert.ToInt32(number), sqlConnection);

            try
            {
                sqlReader = await command3.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string date12 = Convert.ToString(sqlReader["Дата_оформления"]);
                    string[] date1 = date12.Split(new char[] { ' ' });
                    string date = date1[0];
                    data_ofr= date1[0];
                    richTextBox4.Text += "Дата_оформления:                  " + (date + "\n");
                    richTextBox4.Text += "Заявитель:                                 " + (Convert.ToString(sqlReader["Заявитель"] + "\n"));
                    name = Convert.ToString(sqlReader["Заявитель"]);
                    richTextBox4.Text += "Контактный телефон:             " + (Convert.ToString(sqlReader["Контактный_телефон"] + "\n"));
                    richTextBox4.Text += "Электронная почта:                " + (Convert.ToString(sqlReader["Электронная_почта"] + "\n"));
                    mail = Convert.ToString(sqlReader["Электронная_почта"]);
                    richTextBox4.Text += "Адрес:                                       " + (Convert.ToString(sqlReader["Адрес"] + "\n"));
                    adress = Convert.ToString(sqlReader["Адрес"]);

                    kodorg = (Convert.ToString(sqlReader["Код_организации"]));
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
            SqlDataReader sqlReader2 = null;
            SqlCommand command2 = new SqlCommand("SELECT * FROM [Организации] WHERE Код_организации=" + Convert.ToInt32(kodorg), sqlConnection);
            try
            {
                sqlReader2 = await command2.ExecuteReaderAsync();
                while (await sqlReader2.ReadAsync())
                {

                    richTextBox4.Text += "Организация:                           " + (Convert.ToString(sqlReader2["Название"] + "\n"));
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader2 != null)
                {
                    sqlReader2.Close();
                    b = false;
                }

            }
            SqlDataReader sqlReader3 = null;
            SqlCommand command4 = new SqlCommand("SELECT * FROM [Заявки] WHERE Номер_заявки=" + Convert.ToInt32(number), sqlConnection);

            try
            {
                sqlReader3 = await command4.ExecuteReaderAsync();

                while (await sqlReader3.ReadAsync())
                {
                    string date12 = Convert.ToString(sqlReader3["Дата_завершения"]);
                    string[] date1 = date12.Split(new char[] { ' ' });
                    string date = date1[0];
                    richTextBox4.Text += "Область заявки:                        " + (Convert.ToString(sqlReader3["Тип_заявки"] + "\n"));
                    richTextBox4.Text += "Жалоба:                                     " + (Convert.ToString(sqlReader3["Жалоба"] + "\n"));
                    richTextBox4.Text += "Дата_завершения:                    " + (date + "\n");
                    richTextBox4.Text += "Статус:" + "                                       " + (Convert.ToString(sqlReader3["Статус"] + "\n"));
                    richTextBox4.Text += "Результат заявки:                      " + (Convert.ToString(sqlReader3["Результат_заявки"] + "\n"));
                    textBox2.Text = Convert.ToString(sqlReader3["Результат_заявки"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader3 != null)
                    sqlReader3.Close();
            }
        }
        private async void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns[4].Width = 205;
            dataGridView1.Columns[3].Width = 129;
            string number = "" + textBox1.Text;
            label6.Text = "Работа  с завкой №" + number;
            label1.Text += number;
            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet5.Работы_по_заявкам". При необходимости она может быть перемещена или удалена.
            this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
            работыпозаявкамBindingSource.Filter = "Номер_заявки=" + number;
            string connectionString = @"Data Source = 31.31.198.141; Initial Catalog = u1666130_JKH34; User ID = u1666130_Yuliya; Password = csb#G254; MultipleActiveResultSets=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            update();
            proverka= true;
        }


        private  void Form3_Activated(object sender, EventArgs e)
        {
            if (proverka==true)
            {
                update();
            }
        }
      
        private void button6_Click(object sender, EventArgs e)
        {
            string kod_zayavki = textBox1.Text;
            Form4 form = new Form4(kod_zayavki);
            form.textBox5.Text = "" + Convert.ToString(kod_zayavki);
            form.ShowDialog();
            b = false;

        }

        private async void button5_Click(object sender, EventArgs e)//сохранить как документ
        {
            string nomer=Convert.ToString(textBox1.Text);
            string date12 = Convert.ToString(dateTimePicker1.Value.Date);
            string[] date1 = date12.Split(new char[] { ' ' });
            string date_otp = date1[0];
            string[] name1 = name.Split(new char[] { ' ' });
            string nameIO=name1[1] +" " +name1[2];
            string text;
            string[] words;
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            if (!string.IsNullOrEmpty(richTextBox3.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE [Заявки] SET [Результат_заявки]=@Результат_заявки WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                command.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox3.Text));
                await command.ExecuteNonQueryAsync();
                text = Convert.ToString(richTextBox3.Text);
                words = text.Split(new char[] { '.' });
            }
            else
            {

                text = Convert.ToString(textBox2.Text);

                words = text.Split(new char[] { '.' });
            }
            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWordSrub("{name}", name, wordDocument);
                ReplaceWordSrub("{adress}", adress, wordDocument);
                ReplaceWordSrub("{nomer}", nomer, wordDocument);
                ReplaceWordSrub("{date_otp}", date_otp, wordDocument);
                ReplaceWordSrub("{nameIO}", nameIO, wordDocument);
                ReplaceWordSrub("{date_ofr}", data_ofr, wordDocument);
                ReplaceWordSrub("{text}", words[0] + ".", wordDocument);
                int index = 1;
                string zamena = "";
                for (int i = 1; i < words.Length - 1; i++)
                {

                    ReplaceWordSrub("{text" + index + "}", words[i] + ".", wordDocument);
                    index++;
                }
                if (index != 40)
                {
                    for (int i = index; i < 41; i++)
                    {
                        ReplaceWordSrub("{text" + i + "}", zamena, wordDocument);
                    }
                }
                //wordDocument.SaveAs(@"C:\Users\Lilith\Documents\sotrudnik1.docx");
                wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка!");
            }

        }

        private async void button4_Click(object sender, EventArgs e)//отправить на электронную почту
        {
            if (mail != "")
            {
                if (richTextBox3.Text != "")
                {
                    string [] name1 = name.Split(new char[] { ' ' });
                    string nameIO = name1[1] + " " + name1[2];
                    var fromAddress = new MailAddress("gkupk111@gmail.com", "ГКУ ПК Гражданская защита");
                    var toAddress = new MailAddress(mail, name);
                    const string fromPassword = "Orhedeya0852";
                    string subject = "Ответ по заявке №" + Convert.ToString(textBox1.Text);
                    string body = "Уважаемый(ая) "+ nameIO +"!\n"+ "В ответ на ваше обращение от "+ data_ofr + " сообщаем следующее:\n"+Convert.ToString(richTextBox3.Text);

                    var smtp = new SmtpClient
                    {
                        Host = "smtp.gmail.com",
                        Port = 587,
                        EnableSsl = true,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                    };
                    using (var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = subject,
                        Body = body
                    })
                    {
                        smtp.Send(message);
                    };


                    SqlCommand command = new SqlCommand("UPDATE [Заявки] SET [Результат_заявки]=@Результат_заявки WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                    command.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox3.Text));
                    await command.ExecuteNonQueryAsync();
                    MessageBox.Show("Успешно отправлено!");
                }
                else
                {
                    var fromAddress = new MailAddress("gkupk111@gmail.com", "ГКУ ПК Гражданская защита");
                    var toAddress = new MailAddress(mail, name);
                    const string fromPassword = "Orhedeya0852";
                    string subject = "Ответ по заявке №" + Convert.ToString(textBox1.Text);
                    string body = Convert.ToString(textBox2.Text);

                    var smtp = new SmtpClient
                    {
                        Host = "smtp.gmail.com",
                        Port = 587,
                        EnableSsl = true,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                    };
                    using (var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = subject,
                        Body = body
                    })
                    {
                        smtp.Send(message);
                    };


                    SqlCommand command = new SqlCommand("UPDATE [Заявки] SET [Результат_заявки]=@Результат_заявки WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                    command.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox3.Text));
                    await command.ExecuteNonQueryAsync();
                    MessageBox.Show("Успешно отправлено!");
                }
            }
            else MessageBox.Show("У данного заявителя не указан адрес электронной почты!");

        }

        private void richTextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            string proverka_zayavki = Convert.ToString(richTextBox3.Text);
            string[] words = proverka_zayavki.Split(new char[] { '.' });
            if (proverka_zayavki.Length < 2000 && words.Length < 42)
            {
                richTextBox3.ReadOnly = false;
            }
            else
            {
                if (e.KeyCode != Keys.Back)
                {
                    richTextBox3.ReadOnly = true;
                    MessageBox.Show("Длина заявки превышает допустимую!");
                }
                else
                {
                    richTextBox3.ReadOnly = false;
                }
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            работыпозаявкамBindingSource.Filter = "";
            int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
            int kod_org = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
            работыпозаявкамBindingSource.Filter = "Номер_заявки=" + Convert.ToString(textBox1.Text);
            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
               SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Работы_по_заявкам] (Номер_работы, Номер_заявки," +
                    " Дата_заполнения,Действие_в_отношении_заявки)VALUES(@Номер_работы,@Номер_заявки,@Дата_заполнения,@Действие_в_отношении_заявки)", sqlConnection);
                command.Parameters.AddWithValue("Номер_работы", Convert.ToInt32(kod_org) + 1);
                command.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(textBox1.Text));
                command.Parameters.AddWithValue("Дата_заполнения", dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("Действие_в_отношении_заявки", Convert.ToString(comboBox1.Text));
                await command.ExecuteNonQueryAsync();
                this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
                MessageBox.Show("Успешно добавлено!");
            }
            else MessageBox.Show("Заполните все значения!");
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 1)
            {
                string proverka = Convert.ToString(dataGridView1[4, dataGridView1.RowCount - 2].Value);
                if (proverka != "")
                {
                    работыпозаявкамBindingSource.Filter = "";
                    int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
                    int kod_org = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
                    работыпозаявкамBindingSource.Filter = "Номер_заявки=" + Convert.ToString(textBox1.Text);

                    if (!string.IsNullOrEmpty(richTextBox1.Text))
                    {
                        SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Работы_по_заявкам] (Номер_работы, Номер_заявки," +
                             " Дата_заполнения,Действие_в_отношении_заявки,Ход_работы)VALUES(@Номер_работы,@Номер_заявки,@Дата_заполнения,@Действие_в_отношении_заявки," +
                             "@Ход_работы)", sqlConnection);
                        command.Parameters.AddWithValue("Номер_работы", Convert.ToInt32(kod_org) + 1);
                        command.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(textBox1.Text));
                        command.Parameters.AddWithValue("Дата_заполнения", dateTimePicker1.Value.Date);
                        command.Parameters.AddWithValue("Действие_в_отношении_заявки", Convert.ToString(dataGridView1[3, dataGridView1.RowCount - 2].Value));
                        command.Parameters.AddWithValue("Ход_работы", Convert.ToString(richTextBox1.Text));
                        await command.ExecuteNonQueryAsync();
                        richTextBox1.Clear();
                        this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
                        MessageBox.Show("Успешно добавлено!");
                    }
                    else MessageBox.Show("Заполните все значения!");
                }
                else
                {
                    if (richTextBox1.Text != "")
                    {
                        int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
                        int kod_org = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
                        SqlCommand command = new SqlCommand("UPDATE [Работы_по_заявкам] SET [Ход_работы]=@Ход_работы WHERE [Номер_работы]=" + kod_org, sqlConnection);
                        command.Parameters.AddWithValue("Ход_работы", Convert.ToString(richTextBox1.Text));
                        await command.ExecuteNonQueryAsync();
                        richTextBox1.Clear();
                        this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
                        MessageBox.Show("Успешно добавлено!");
                    }
                    else
                    {
                        MessageBox.Show(
                          "Вы не заполнили ответ!");
                    }

                }
            }
            else MessageBox.Show("Вы еще не вынесли решение по заявке!");
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count!=1)
            { 
            string status_zavershen = "Завершен";
            string proverka = Convert.ToString(dataGridView1[4, dataGridView1.RowCount - 2].Value);
            if (proverka != "")
            {
                работыпозаявкамBindingSource.Filter = "";
                int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
                int kod_org = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
                работыпозаявкамBindingSource.Filter = "Номер_заявки=" + Convert.ToString(textBox1.Text);

                if (!string.IsNullOrEmpty(richTextBox1.Text))
                {
                    SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Работы_по_заявкам] (Номер_работы, Номер_заявки," +
                         " Дата_заполнения,Действие_в_отношении_заявки,Ход_работы)VALUES(@Номер_работы,@Номер_заявки,@Дата_заполнения,@Действие_в_отношении_заявки," +
                         "@Ход_работы)", sqlConnection);
                    command.Parameters.AddWithValue("Номер_работы", Convert.ToInt32(kod_org) + 1);
                    command.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(textBox1.Text));
                    command.Parameters.AddWithValue("Дата_заполнения", dateTimePicker1.Value.Date);
                    command.Parameters.AddWithValue("Действие_в_отношении_заявки", Convert.ToString(dataGridView1[3, dataGridView1.RowCount - 2].Value));
                    command.Parameters.AddWithValue("Ход_работы", Convert.ToString(richTextBox1.Text));
                    await command.ExecuteNonQueryAsync();
                    SqlCommand command1 = new SqlCommand("UPDATE [Заявки] SET [Результат_заявки]=@Результат_заявки,[Статус]=@Статус,[Дата_завершения]=@Дата_завершения WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                    command1.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox1.Text));
                    command1.Parameters.AddWithValue("Статус", status_zavershen);
                    command1.Parameters.AddWithValue("Дата_завершения", dateTimePicker1.Value.Date);
                    await command1.ExecuteNonQueryAsync();
                    update();
                    richTextBox1.Clear();
                    this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
                    MessageBox.Show("Заявка успешно закрыта!");
                }
                else MessageBox.Show("Заполните все значения!");
            }
            else
            {
                if (richTextBox1.Text != "")
                {
                    int lastrow = Convert.ToInt32(dataGridView1.RowCount.ToString()) - 2;
                    int kod_org = Convert.ToInt32(dataGridView1.Rows[lastrow].Cells[0].Value.ToString());
                    SqlCommand command = new SqlCommand("UPDATE [Работы_по_заявкам] SET [Ход_работы]=@Ход_работы WHERE [Номер_работы]=" + kod_org, sqlConnection);
                    command.Parameters.AddWithValue("Ход_работы", Convert.ToString(richTextBox1.Text));
                    await command.ExecuteNonQueryAsync();
                    SqlCommand command1 = new SqlCommand("UPDATE [Заявки] SET [Результат_заявки]=@Результат_заявки,[Статус]=@Статус,[Дата_завершения]=@Дата_завершения WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                    command1.Parameters.AddWithValue("Результат_заявки", Convert.ToString(richTextBox1.Text));
                    command1.Parameters.AddWithValue("Статус", status_zavershen);
                    command1.Parameters.AddWithValue("Дата_завершения", dateTimePicker1.Value.Date);
                    await command1.ExecuteNonQueryAsync();
                    update();
                    richTextBox1.Clear();
                    this.работы_по_заявкамTableAdapter.Fill(this.u1666130_JKH34DataSet5.Работы_по_заявкам);
                    MessageBox.Show("Заявка успешно закрыта!");
                }
                else
                {
                    MessageBox.Show(
                      "Вы не заполнили ответ!");
                }

            }
            }
            else MessageBox.Show("Вы еще не вынесли решение по заявке!");
        }

        private async void button7_Click(object sender, EventArgs e)
        {
            string status_zavershen = "Завершен";
            if (textBox2.Text != "")
            {
                SqlCommand command1 = new SqlCommand("UPDATE [Заявки] SET [Статус]=@Статус,[Дата_завершения]=@Дата_завершения WHERE [Номер_заявки]=" + Convert.ToString(textBox1.Text), sqlConnection);
                command1.Parameters.AddWithValue("Статус", status_zavershen);
                command1.Parameters.AddWithValue("Дата_завершения", dateTimePicker1.Value.Date);
                await command1.ExecuteNonQueryAsync();
                update();
                MessageBox.Show("Заявка закрыта!");
            }
            else MessageBox.Show("Вы не можете закрыть заявку пока не\nотправили ответ заявителю или не\nуказали результат в ходе работы! ");

        }
    }
}
