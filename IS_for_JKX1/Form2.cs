using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace IS_for_JKX1
{
    public partial class Form2 : Form
    {
        public static SqlConnection sqlConnection = null;
        public void redup()
        {
            for (int i = 0; Convert.ToInt32(dataGridView1.Rows.Count - 1) > i; i++)
            {
                DateTime date1 = dateTimePicker6.Value.Date;
                DateTime date2 = (DateTime)dataGridView1[1, i].Value;
                TimeSpan timeSpan = (date1.Subtract(date2));
                string date = Convert.ToString(timeSpan);
                string[] date3 = date.Split(new char[] { '.' });
                if (date3[0] != "00:00:00")
                {
                    int proverka_red = Convert.ToInt32(date3[0]);
                    if (proverka_red > 29)
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Tomato;
                    }
                }
            }
        }
        public Form2(string kod)
        {
            InitializeComponent();
            заявкиBindingSource.Filter = "Статус like '" + status_v_processe + "'";
            организацииBindingSource.Filter = "Статус like '" + status_deyst + "'";
        }
        private readonly string TemplateFileName = @"C:\sotrudniki.docx";
        private readonly string TemplateFileName1 = @"C:\zayavka.doc";
        private void ReplaceWordSrub(string stubToReplace, string text, Word.Document wordDocument)//заполнение документов
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);

        }
        public bool proverka_perenosa = false;
        public string kod_zayavki;
        public string status_zavershen = "Завершен";
        public string status_v_processe = "В процессе";
        string status_deyst = "Действующий";
        string status_nedeyst = "Прекратил работу";
        string status_rab = "Работает";
        string status_yvolt = "Уволен";
        public bool pr = false;//проверка добавления организаций
        public bool pr1 = false;//проверка добавления сотрудников
        public bool pr_obl = false;//проверка добавления областей
        public string obl;//добавленная область-если удалять
        public string org;
        public bool PROVERKA;//проверка загразки формы
        public void combUPorg()
        {
            comboBox1.Items.Clear();
            comboBox4.Items.Clear();
            comboBox6.Items.Clear();
            for (int i = 0; i < dataGridView4.RowCount; i++)
            {
                int j = 2;
                if (dataGridView4.Rows[i].Cells[j].Value != null)
                {
                    comboBox1.Items.Add(Convert.ToString(dataGridView4.Rows[i].Cells[j].Value));
                    comboBox4.Items.Add(Convert.ToString(dataGridView4.Rows[i].Cells[j].Value));
                }
                j = 0;
                if (dataGridView4.Rows[i].Cells[j].Value != null)
                {
                    comboBox6.Items.Add(Convert.ToString(dataGridView4.Rows[i].Cells[j].Value));
                }
            }
        }
        public void combUPobl()
        {
            comboBox2.Items.Clear();
            comboBox10.Items.Clear();
            comboBox3.Items.Clear();
            for (int i = 0; i < dataGridView8.RowCount; i++)
            {
                int j = 0;
                if (dataGridView8.Rows[i].Cells[j].Value != null)
                {
                    comboBox2.Items.Add(Convert.ToString(dataGridView8.Rows[i].Cells[j].Value));
                    comboBox10.Items.Add(Convert.ToString(dataGridView8.Rows[i].Cells[j].Value));
                    comboBox3.Items.Add(Convert.ToString(dataGridView8.Rows[i].Cells[j].Value));
                }
            }
        }
        private async void Form2_Load(object sender, EventArgs e)
        {

            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet3.Области_заявок". При необходимости она может быть перемещена или удалена.
            this.области_заявокTableAdapter.Fill(this.u1666130_JKH34DataSet3.Области_заявок);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet2.Сотрудники_предприятия". При необходимости она может быть перемещена или удалена.
            this.сотрудники_предприятияTableAdapter.Fill(this.u1666130_JKH34DataSet2.Сотрудники_предприятия);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet1.Организации". При необходимости она может быть перемещена или удалена.
            this.организацииTableAdapter.Fill(this.u1666130_JKH34DataSet1.Организации);
            comboBox5.ItemHeight = 20;
            // TODO: данная строка кода позволяет загрузить данные в таблицу "u1666130_JKH34DataSet.Заявки". При необходимости она может быть перемещена или удалена.
            this.заявкиTableAdapter.Fill(this.u1666130_JKH34DataSet.Заявки);
            string connectionString = @"Data Source = 31.31.198.141; Initial Catalog = u1666130_JKH34; User ID = u1666130_Yuliya; Password = csb#G254";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            dataGridView4.Columns[2].Width = 127;
            dataGridView8.Columns[0].Width = 435;
            dataGridView5.Columns[4].Width = 150;

            combUPobl();
            combUPorg();
            redup();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void label23_Click(object sender, EventArgs e)//помощь1
        {
            MessageBox.Show(
                 "Уважаемый пользователь!\nДля сортировки текущих заявок выберите\nинтересующие вас параметры поиска и\nнажмите конпку найти. Для более подробной\nинформации о заявке дважды кликните\nпо ней.",
                 "Сообщение", MessageBoxButtons.OK,
                 MessageBoxIcon.Information,
                 MessageBoxDefaultButton.Button1,
                 MessageBoxOptions.DefaultDesktopOnly);
        }

        private void label22_Click(object sender, EventArgs e)//помощь2
        {
            MessageBox.Show(
                "Уважаемый пользователь!\nДля сортировки завершенных заявок выберите\nинтересующие вас параметры поиска и нажмите\nкнопку найти. Для перемещения заявки в\nтекущие дважды кликните по ней.",
                 "Сообщение", MessageBoxButtons.OK,
                 MessageBoxIcon.Information,
                 MessageBoxDefaultButton.Button1,
                 MessageBoxOptions.DefaultDesktopOnly);
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)//сортировка заявок по умолчанию
        {
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])
            {
                заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
            }
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage7"])
            {
                заявкиBindingSource.Filter = "";
            }
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage3"])
            {
                заявкиBindingSource.Filter = "Статус like '" + status_v_processe + "'";
                redup();
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)//поиск по номеру в текщих заявках
        {
            if (textBox2.Text != "")
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].Selected = false;
                    int j = 0;
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox2.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                        }
                    }

                }
            }
            else dataGridView1.ClearSelection();

        }

        private void textBox3_TextChanged(object sender, EventArgs e)//поиск по заявителю в текущих заявках
        {
            if (textBox3.Text != "")
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].Selected = false;
                    int j = 2;
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox3.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                        }
                    }
                }
            }
            else dataGridView1.ClearSelection();

        }

        private void textBox5_TextChanged(object sender, EventArgs e)//поиск по номеру в завершенных заявках
        {
            if (textBox5.Text != "")
            {
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    dataGridView2.Rows[i].Selected = false;
                    int j = 0;
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox5.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                        }
                    }

                }
            }
            else dataGridView2.ClearSelection();

        }

        private void textBox4_TextChanged(object sender, EventArgs e)//поиск по заявителю в завершенныхх щаявках
        {
            if (textBox4.Text != "")
            {
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    dataGridView2.Rows[i].Selected = false;
                    int j = 2;
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox4.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                        }
                    }
                }
            }
            else dataGridView2.ClearSelection();

        }

        private void button1_Click(object sender, EventArgs e)//сортировка текущих заявок
        {
            int kodorg = 0;
            if (comboBox1.SelectedIndex != -1)
            {
                int index = comboBox1.SelectedIndex;
                kodorg = Convert.ToInt32(comboBox6.Items[index].ToString());
            }
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            if (checkBox1.Checked) заявкиBindingSource.Filter = "Дата_оформления=" + "'" + dateTimePicker1.Value.Date + "'" + "AND " + "Статус like '" + status_v_processe + "'";
            if (checkBox2.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Статус like '" + status_v_processe + "'";
            if (checkBox3.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox2.Text + "'" + "AND " + "Статус like '" + status_v_processe + "'";
            if (checkBox1.Checked && checkBox2.Checked) заявкиBindingSource.Filter = "Дата_оформления=" + "'" + dateTimePicker1.Value.Date + "'" + "AND " + "Статус like '" + status_v_processe + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'";
            if (checkBox2.Checked && checkBox3.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Статус like '" + status_v_processe + "'" + "AND " + "Тип_заявки=" + "'" + comboBox2.Text + "'";
            if (checkBox1.Checked && checkBox3.Checked) заявкиBindingSource.Filter = "Дата_оформления=" + "'" + dateTimePicker1.Value.Date + "'" + "AND " + "Статус like '" + status_v_processe + "'" + "AND " + "Тип_заявки=" + "'" + comboBox2.Text + "'";
            if (checkBox1.Checked && checkBox3.Checked && checkBox2.Checked) заявкиBindingSource.Filter = "Дата_оформления=" + "'" + dateTimePicker1.Value.Date + "'" + "AND " + "Статус like '" + status_v_processe + "'" + "AND " + "Тип_заявки=" + "'" + comboBox2.Text + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'";

        }

        private void button2_Click(object sender, EventArgs e)//отмена сортировка текущих заявок
        {
            заявкиBindingSource.Filter = "Статус like '" + status_v_processe + "'";
        }

        private void button4_Click(object sender, EventArgs e)//сортировка завершенных заявок
        {
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            dateTimePicker3.CustomFormat = "dd.MM.yyyy";
            dateTimePicker4.CustomFormat = "dd.MM.yyyy";
            if (checkBox6.Checked) заявкиBindingSource.Filter = "Дата_завершения=" + "'" + dateTimePicker2.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            //период
            if (checkBox5.Checked) заявкиBindingSource.Filter = "Дата_завершения>=" + "'" + dateTimePicker3.Value.Date + "'" + "AND " + "Дата_завершения<=" + "'" + dateTimePicker4.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
        }

        private void button3_Click(object sender, EventArgs e)//отмена сортировка завершенных заявок
        {
            заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
        }

        private void textBox16_TextChanged(object sender, EventArgs e)//поиск по организациям
        {
            if (textBox16.Text != "")
            {
                for (int i = 0; i < dataGridView4.RowCount; i++)
                {
                    dataGridView4.Rows[i].Selected = false;
                    int j = 2;
                    if (dataGridView4.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView4.Rows[i].Cells[j].Value.ToString().Contains(textBox16.Text))
                        {
                            dataGridView4.Rows[i].Selected = true;
                        }
                    }
                }
            }
            else dataGridView4.ClearSelection();

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)//фильтр архива
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])
                {
                    заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
                }
                if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage7"])
                {
                    заявкиBindingSource.Filter = "";
                }
                if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage3"])
                {
                    заявкиBindingSource.Filter = "Статус like '" + status_v_processe + "'";
                    redup();
                }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                организацииBindingSource.Filter = "Статус like '" + status_deyst + "'";
                сотрудникипредприятияBindingSource.Filter = "Статус like '" + status_rab + "'";
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                организацииBindingSource.Filter = "Статус like '" + status_nedeyst + "'";
                сотрудникипредприятияBindingSource.Filter = "Статус like '" + status_yvolt + "'";
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage6"])
            {
                заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
                otch();
                label44.Text = "Количество завершенных заявок: "+Convert.ToString(dataGridView9.Rows.Count);
            }

        }

        public async void otch()
        {
            comboBox11.Items.Clear();
            comboBox7.Items.Clear();
            comboBox13.Items.Clear();
            comboBox12.Items.Clear();
            организацииBindingSource.Filter = "";
            сотрудникипредприятияBindingSource.Filter = "";
            SqlDataReader sqlReader2 = null;
            SqlCommand command2 = new SqlCommand("SELECT * FROM [Организации]", sqlConnection);
            try
            {
                sqlReader2 = await command2.ExecuteReaderAsync();
                while (await sqlReader2.ReadAsync())
                {

                    comboBox11.Items.Add(Convert.ToString(sqlReader2["Название"]));
                    comboBox7.Items.Add(Convert.ToString(sqlReader2["Код_организации"]));
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
                }

            }
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Сотрудники_предприятия]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {

                    comboBox12.Items.Add(Convert.ToString(sqlReader["ФИО_сотрудника"]));
                    comboBox13.Items.Add(Convert.ToString(sqlReader["Код_сотрудника"]));
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                {
                    sqlReader.Close();
                }

            }
        }

        private void button12_Click(object sender, EventArgs e)//генерация логина сотрудника
        {
            string iPass = "";
            string[] arr = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Z", "b", "c", "d", "f", "g", "h", "j", "k", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "z", "A", "E", "U", "Y", "a", "e", "i", "o", "u", "y" };
            Random rnd = new Random();
            for (int i = 0; i < 6; i = i + 1)
            {
                iPass = iPass + arr[rnd.Next(0, 57)];
            }
            textBox19.Text = "user" + iPass;

        }

        private void button11_Click(object sender, EventArgs e)//генерация пароля сотрудника
        {
            string iPass = "";
            string[] arr = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Z", "b", "c", "d", "f", "g", "h", "j", "k", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "z", "A", "E", "U", "Y", "a", "e", "i", "o", "u", "y" };
            Random rnd = new Random();
            for (int i = 0; i < 10; i = i + 1)
            {
                iPass = iPass + arr[rnd.Next(0, 57)];
            }
            textBox18.Text = iPass;

        }

        private void textBox19_TextChanged(object sender, EventArgs e)//предложение генерации пароля
        {
            button12.Visible = true;
        }

        private void textBox18_TextChanged(object sender, EventArgs e)//предложение генерации логина
        {
            button11.Visible = true;
        }

        private async void button10_Click(object sender, EventArgs e)//добавление сотрдуников
        {
            if (!string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrEmpty(textBox18.Text) &&
                !string.IsNullOrEmpty(comboBox9.Text) && !string.IsNullOrEmpty(textBox20.Text) &&
                !string.IsNullOrEmpty(textBox19.Text))
            {
                сотрудникипредприятияBindingSource.Filter = "";
                int lastrow = Convert.ToInt32(dataGridView7.RowCount.ToString()) - 2;
                int kod_org = Convert.ToInt32(dataGridView7.Rows[lastrow].Cells[0].Value.ToString());
                сотрудникипредприятияBindingSource.Filter = "Статус like '" + status_rab + "'";
                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Сотрудники_предприятия] (Код_сотрудника, ФИО_сотрудника," +
                    " Роль, Контактный_телефон, Логин, Пароль, Статус, Дата_приема_в_работу)VALUES(@Код_сотрудника,@ФИО_сотрудника,@Роль,@Контактный_телефон," +
                    "@Логин, @Пароль, @Статус, @Дата_приема_в_работу)", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", Convert.ToInt32(kod_org) + 1);
                command.Parameters.AddWithValue("ФИО_сотрудника", Convert.ToString(textBox21.Text));
                command.Parameters.AddWithValue("Роль", Convert.ToString(comboBox9.Text));
                command.Parameters.AddWithValue("Контактный_телефон", Convert.ToString(textBox20.Text));
                command.Parameters.AddWithValue("Логин", Convert.ToString(textBox19.Text));
                command.Parameters.AddWithValue("Пароль", Convert.ToString(textBox18.Text));
                command.Parameters.AddWithValue("Статус", Convert.ToString(status_rab));
                command.Parameters.AddWithValue("Дата_приема_в_работу", dateTimePicker6.Value.Date);
                await command.ExecuteNonQueryAsync();
                this.сотрудники_предприятияTableAdapter.Fill(this.u1666130_JKH34DataSet2.Сотрудники_предприятия);
                textBox21.Clear();
                textBox20.Clear();
                textBox19.Clear();
                textBox18.Clear();
                pr1 = true;
                MessageBox.Show("Успешно добавлено!");
            }
            else MessageBox.Show("Заполните все значения!");
            button12.Visible = false;
            button11.Visible = false;
        }

        private async void button9_Click(object sender, EventArgs e)//удаление сотрдуников
        {
            if (pr1 == true)
            {
                int lastrow = Convert.ToInt32(dataGridView7.RowCount.ToString()) - 2;
                int kod_s = Convert.ToInt32(dataGridView7.Rows[lastrow].Cells[0].Value.ToString());
                SqlCommand command = new SqlCommand("DELETE FROM [dbo].[Сотрудники_предприятия] WHERE Код_сотрудника=@Код_сотрудника", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", Convert.ToInt32(kod_s));
                await command.ExecuteNonQueryAsync();
                pr1 = false;
                this.сотрудники_предприятияTableAdapter.Fill(this.u1666130_JKH34DataSet2.Сотрудники_предприятия);
                MessageBox.Show("Успешно удалено!");
            }
            else
            {
                MessageBox.Show("Вы еще не добавили ни одной строки!");
            }

        }

        private async void dataGridView7_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//перенос сотрудников в архив
        {
            if (e.RowIndex < dataGridView7.Rows.Count - 1)
            {
                string kod = dataGridView7.Rows[dataGridView7.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult result = MessageBox.Show(
                       "Отменить данного сотрудника как прекратившего\n работу в предприятии и перенести в архив?",
                       "Сообщение",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Information,
                       MessageBoxDefaultButton.Button1,
                       MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    SqlCommand command = new SqlCommand("UPDATE [dbo].[Сотрудники_предприятия] SET [Статус]=@Статус,[Дата_увольнения]=@Дата_увольнения WHERE [Код_сотрудника]=@Код_сотрудника", sqlConnection);
                    command.Parameters.AddWithValue("Статус", status_yvolt);
                    command.Parameters.AddWithValue("Код_сотрудника", Convert.ToInt32(kod));
                    command.Parameters.AddWithValue("Дата_увольнения", dateTimePicker6.Value.Date);
                    await command.ExecuteNonQueryAsync();
                    this.сотрудники_предприятияTableAdapter.Fill(this.u1666130_JKH34DataSet2.Сотрудники_предприятия);
                    MessageBox.Show("Успешно перемещено!");
                }
            }
        }

        private async void tabControl3_SelectedIndexChanged(object sender, EventArgs e)//контроль администратора
        {
            if (tabControl3.SelectedTab == tabControl3.TabPages["tabPage9"])
            {
                string proverka_adm;
                string proverka_proverka;
                SqlDataReader sqlReader = null;
                SqlCommand command1 = new SqlCommand("SELECT * FROM [Сотрудники_предприятия]", sqlConnection);

                try
                {
                    sqlReader = await command1.ExecuteReaderAsync();

                    while (await sqlReader.ReadAsync())
                    {
                        proverka_proverka = Convert.ToString(sqlReader["Код_сотрудника"]);
                        proverka_adm = Convert.ToString(sqlReader["Роль"]);
                        if (textBox1.Text == proverka_proverka && proverka_adm != "Администратор")
                        {
                            tabControl3.SelectedTab = tabControl3.TabPages["tabPage8"];
                            MessageBox.Show("У вас недосаточно прав для доступа к этому разделу!");
                        }
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

        private async void button14_Click(object sender, EventArgs e)//добавление области
        {
            if (textBox17.Text != "")
            {
                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Области_заявок] (Тип_заявки)VALUES(@Тип_заявки)", sqlConnection);
                command.Parameters.AddWithValue("Тип_заявки", Convert.ToString(textBox17.Text));
                await command.ExecuteNonQueryAsync();
                pr_obl = true;
                obl = Convert.ToString(textBox17.Text);
                textBox17.Clear();
                this.области_заявокTableAdapter.Fill(this.u1666130_JKH34DataSet3.Области_заявок);
                MessageBox.Show("Успешно добавлено!");
                combUPobl();
            }
            else MessageBox.Show("Заполните наименование области!");
        }

        private async void button15_Click(object sender, EventArgs e)//удаление области
        {
            if (pr_obl == true)
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Области_заявок] WHERE Тип_заявки=@Тип_заявки", sqlConnection);
                command.Parameters.AddWithValue("Тип_заявки", Convert.ToString(obl));
                await command.ExecuteNonQueryAsync();
                this.области_заявокTableAdapter.Fill(this.u1666130_JKH34DataSet3.Области_заявок);
                MessageBox.Show("Успешно удалено!");
                pr_obl = false;
                combUPobl();
            }
            else MessageBox.Show("Вы еще не добавили область\n чтобы ее удалить!");

        }

        private async void button7_Click(object sender, EventArgs e)//добавление огранизация
        {
            организацииBindingSource.Filter = "";
            int lastrow = Convert.ToInt32(dataGridView4.RowCount.ToString()) - 2;
            int kod_org = Convert.ToInt32(dataGridView4.Rows[lastrow].Cells[0].Value.ToString());
            организацииBindingSource.Filter = "Статус like '" + status_deyst + "'";
            if (!string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrEmpty(textBox13.Text) &&
                !string.IsNullOrEmpty(textBox14.Text) && !string.IsNullOrEmpty(comboBox8.Text) &&
                !string.IsNullOrEmpty(textBox15.Text))
            {
                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Организации] (Код_организации, Тип_организации," +
                    " Название, Электронная_почта, Адрес, Номер_телефона, Статус)VALUES(@Код_организации,@Тип_организации,@Название,@Электронная_почта," +
                    "@Адрес, @Номер_телефона, @Статус)", sqlConnection);
                command.Parameters.AddWithValue("Код_организации", Convert.ToInt32(kod_org) + 1);
                command.Parameters.AddWithValue("Тип_организации", Convert.ToString(comboBox8.Text));
                command.Parameters.AddWithValue("Название", Convert.ToString(textBox12.Text));
                command.Parameters.AddWithValue("Электронная_почта", Convert.ToString(textBox13.Text));
                command.Parameters.AddWithValue("Адрес", Convert.ToString(textBox14.Text));
                command.Parameters.AddWithValue("Номер_телефона", Convert.ToString(textBox15.Text));
                command.Parameters.AddWithValue("Статус", Convert.ToString(status_deyst));
                await command.ExecuteNonQueryAsync();
                this.организацииTableAdapter.Fill(this.u1666130_JKH34DataSet1.Организации);
                textBox12.Clear();
                textBox13.Clear();
                textBox14.Clear();
                textBox15.Clear();
                org = Convert.ToString(kod_org + 1);
                pr = true;
                MessageBox.Show("Успешно добавлено!");
                combUPorg();
            }

            else MessageBox.Show("Заполните все значения!");


        }

        private async void button8_Click(object sender, EventArgs e)//удаление огранизация
        {
            if (pr == true)
            {
                int lastrow = Convert.ToInt32(dataGridView4.RowCount.ToString()) - 2;
                int kod_org = Convert.ToInt32(dataGridView4.Rows[lastrow].Cells[0].Value.ToString());
                SqlCommand command = new SqlCommand("DELETE FROM [Организации] WHERE Код_организации=@Код_организации", sqlConnection);
                command.Parameters.AddWithValue("Код_организации", Convert.ToInt32(kod_org));
                await command.ExecuteNonQueryAsync();
                pr = false;
                this.организацииTableAdapter.Fill(this.u1666130_JKH34DataSet1.Организации);
                MessageBox.Show("Успешно удалено!");
                combUPorg();
            }
            else
            {
                MessageBox.Show("Вы еще не добавили ни одной строки!");
            }
        }

        private async void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < dataGridView4.Rows.Count - 1)
            {
                string kod = dataGridView4.Rows[dataGridView4.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult result = MessageBox.Show(
                       "Перенести данную организацю в архив\n и отметить как не действительную?",
                       "Сообщение",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Information,
                       MessageBoxDefaultButton.Button1,
                       MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    SqlCommand command = new SqlCommand("UPDATE [Организации] SET [Статус]=@Статус WHERE [Код_организации]=@Код_организации", sqlConnection);
                    command.Parameters.AddWithValue("Статус", status_nedeyst);
                    command.Parameters.AddWithValue("Код_организации", Convert.ToInt32(kod));
                    await command.ExecuteNonQueryAsync();
                    this.организацииTableAdapter.Fill(this.u1666130_JKH34DataSet1.Организации);
                    организацииBindingSource.Filter = "Статус like '" + status_deyst + "'";
                    combUPorg();
                    MessageBox.Show("Успешно перемещено!");
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)//сохранение документа для сотрудников
        {
            var name="";
            var role = "";
            var login = "";
            var password = "";
            if(!string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrEmpty(textBox19.Text) &&
                !string.IsNullOrEmpty(comboBox9.Text) && !string.IsNullOrEmpty(textBox18.Text)) { 
            name = Convert.ToString(textBox21.Text);
            role = Convert.ToString(comboBox9.Text);
            login = Convert.ToString(textBox19.Text);
            password = Convert.ToString(textBox18.Text);
            }
            else
            {
                name = Convert.ToString(dataGridView7[1, dataGridView7.RowCount - 2].Value);
                role = Convert.ToString(dataGridView7[2, dataGridView7.RowCount - 2].Value);
                login = Convert.ToString(dataGridView7[4, dataGridView7.RowCount - 2].Value);
                password = Convert.ToString(dataGridView7[5, dataGridView7.RowCount - 2].Value);
            }
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWordSrub("{name}", name, wordDocument);
                ReplaceWordSrub("{login}", login, wordDocument);
                ReplaceWordSrub("{role}", role, wordDocument);
                ReplaceWordSrub("{password}", password, wordDocument);
                //wordDocument.SaveAs(@"C:\Users\Lilith\Documents\sotrudnik1.docx");
                wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка!");
            }

        }

        private void button6_Click(object sender, EventArgs e)//сохранение документа для заявок
        {
            string nomer;
            string name2;
            string name;
            string adress;
            string jkx = "1";
            string text;
            string date;
            string[] words;
            int lastrow = Convert.ToInt32(dataGridView3.RowCount.ToString()) - 2;
            int kod_zayavki2 = Convert.ToInt32(dataGridView3.Rows[lastrow].Cells[0].Value.ToString());
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            if (!string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrEmpty(textBox7.Text) &&
                !string.IsNullOrEmpty(comboBox4.Text) && !string.IsNullOrEmpty(richTextBox1.Text))
            {
                сотрудникипредприятияBindingSource.Filter = "Код_сотрудника=" + textBox1.Text;
                name2 = Convert.ToString(dataGridView7[1, 0].Value);
                nomer = Convert.ToString(kod_zayavki2 + 1);
                name = Convert.ToString(textBox6.Text);
                adress = Convert.ToString(textBox7.Text);
                jkx = Convert.ToString(comboBox4.Text);
                text = Convert.ToString(richTextBox1.Text);
                string date12 = Convert.ToString(dateTimePicker5.Value.Date);
                string[] date1 = date12.Split(new char[] { ' ' });
                date = date1[0];

                words = text.Split(new char[] { '.' });
            }
            else
            {
                сотрудникипредприятияBindingSource.Filter = "Код_сотрудника=" + textBox1.Text;
                name2 = name2 = Convert.ToString(dataGridView7[1, 0].Value);
                nomer = Convert.ToString(dataGridView3[0, dataGridView3.RowCount - 2].Value);
                name = Convert.ToString(dataGridView3[2, dataGridView3.RowCount - 2].Value);
                adress = Convert.ToString(dataGridView3[6, dataGridView3.RowCount - 2].Value);
                string jkx1 = Convert.ToString(dataGridView3[7, dataGridView3.RowCount - 2].Value);
                jkx = comboBox4.Items[comboBox6.Items.IndexOf(jkx1)].ToString();
                text = Convert.ToString(dataGridView3[9, dataGridView3.RowCount - 2].Value);
                string date12 = Convert.ToString(dataGridView3[1, dataGridView3.RowCount - 2].Value);
                string[] date1 = date12.Split(new char[] { ' ' });
                date = date1[0];

                words = text.Split(new char[] { '.' });
            }
            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName1);
                ReplaceWordSrub("{name}", name, wordDocument);
                ReplaceWordSrub("{adress}", adress, wordDocument);
                ReplaceWordSrub("{jkx}", jkx, wordDocument);
                ReplaceWordSrub("{date}", date, wordDocument);
                ReplaceWordSrub("{name1}", name, wordDocument);
                ReplaceWordSrub("{name2}", name2, wordDocument);
                ReplaceWordSrub("{nomer)", nomer, wordDocument);
                ReplaceWordSrub("{text}", words[0] + ".", wordDocument);
                int index = 1;
                string zamena = "";
                for (int i = 1; i < words.Length - 1; i++)
                {

                    ReplaceWordSrub("{ text" + index + "}", words[i] + ".", wordDocument);
                    index++;
                }
                if (index != 30)
                {
                    for (int i = index; i < 31; i++)
                    {
                        ReplaceWordSrub("{ text" + i + "}", zamena, wordDocument);
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

        private void Form2_Activated(object sender, EventArgs e)
        {
            this.заявкиTableAdapter.Fill(this.u1666130_JKH34DataSet.Заявки);
            if (proverka_perenosa == true)
            {
                proverka_perenosa = false;
                заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
            }
            redup();
        }

        private async void button5_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrEmpty(textBox7.Text) &&
                !string.IsNullOrEmpty(comboBox10.Text) && !string.IsNullOrEmpty(richTextBox1.Text)
                && !string.IsNullOrEmpty(comboBox4.Text))
            {
                int lastrow = Convert.ToInt32(dataGridView3.RowCount.ToString()) - 2;
                int kod_zayavki2 = Convert.ToInt32(dataGridView3.Rows[lastrow].Cells[0].Value.ToString());
                int index = comboBox4.SelectedIndex;
                int kodorg = Convert.ToInt32(comboBox6.Items[index].ToString());////код оганизации присвоить через комбобокс comboBox1.Items[3].ToString()
                string mail = Convert.ToString(textBox11.Text) + Convert.ToString(comboBox5.Text);
                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Заявки] (Номер_заявки, Дата_оформления," +
                    " Заявитель, Код_сотрудника, Контактный_телефон, Электронная_почта, Адрес, Код_организации," +
                    "Тип_заявки, Жалоба, Статус)VALUES(@Номер_заявки,@Дата_оформления,@Заявитель,@Код_сотрудника," +
                    "@Контактный_телефон, @Электронная_почта, @Адрес,@Код_организации,@Тип_заявки,@Жалоба,@Статус)", sqlConnection);
                command.Parameters.AddWithValue("Номер_заявки", Convert.ToInt32(kod_zayavki2) + 1);
                command.Parameters.AddWithValue("Дата_оформления", dateTimePicker5.Value.Date);
                command.Parameters.AddWithValue("Заявитель", Convert.ToString(textBox6.Text));
                command.Parameters.AddWithValue("Код_сотрудника", Convert.ToString(textBox1.Text));
                command.Parameters.AddWithValue("Контактный_телефон", Convert.ToString(textBox10.Text));
                command.Parameters.AddWithValue("Электронная_почта", Convert.ToString(mail));
                command.Parameters.AddWithValue("Адрес", Convert.ToString(textBox7.Text));
                command.Parameters.AddWithValue("Код_организации", Convert.ToInt32(kodorg));
                command.Parameters.AddWithValue("Тип_заявки", Convert.ToString(comboBox10.Text));
                command.Parameters.AddWithValue("Жалоба", Convert.ToString(richTextBox1.Text));
                command.Parameters.AddWithValue("Статус", Convert.ToString(status_v_processe));
                await command.ExecuteNonQueryAsync();
                if (mail != "")
                {
                    string date12 = Convert.ToString(dateTimePicker5.Value.Date);
                    string[] date1 = date12.Split(new char[] { ' ' });
                    string date_otp = date1[0];
                    сотрудникипредприятияBindingSource.Filter = "Код_сотрудника=" + textBox1.Text;
                    string name2 = Convert.ToString(dataGridView7[1, 0].Value);
                    var fromAddress = new MailAddress("gkupk111@gmail.com", "ГКУ ПК Гражданская защита");
                    var toAddress = new MailAddress(mail, Convert.ToString(textBox6.Text));
                    const string fromPassword = "Orhedeya0852";
                    string subject = "Уведомлении о принятии заявки ГКУ ПК Гражданская защита";
                    string body = "Ваша заявка №" + kod_zayavki2 + " принята в работу.\nТекст заявки: " + Convert.ToString(richTextBox1.Text) + "\nОформитель заявки: " + name2 + ".\nДата оформления: " + date_otp;

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
                }
                richTextBox1.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox2.Clear();
                textBox11.Clear();
                textBox10.Clear();
                this.заявкиTableAdapter.Fill(this.u1666130_JKH34DataSet.Заявки);
                MessageBox.Show("Успешно добавлено!");
            }
            else
                MessageBox.Show("Заполните все значения!");

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex < dataGridView1.Rows.Count - 1)
            //{
                kod_zayavki = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
                Form3 form = new Form3(kod_zayavki);
                form.textBox1.Text = "" + Convert.ToString(kod_zayavki);
                form.ShowDialog();
            //}
        }

        private void button17_Click(object sender, EventArgs e)
        {
            заявкиBindingSource.Filter = "Статус like '" + status_zavershen + "'";
            label44.Text = "Количество выполненных заявок: " + Convert.ToString(dataGridView9.Rows.Count);
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            string proverka_zayavki = Convert.ToString(richTextBox1.Text);
            string[] words = proverka_zayavki.Split(new char[] { '.' });
            if (proverka_zayavki.Length < 2000 && words.Length < 32)
            {
                richTextBox1.ReadOnly = false;
            }
            else
            {
                if (e.KeyCode != Keys.Back)
                {
                    richTextBox1.ReadOnly = true;
                    MessageBox.Show("Длина заявки превышает допустимую!");
                }
                else
                {
                    richTextBox1.ReadOnly = false;
                }
            }
        }

        private void dataGridView2_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)//изменение статуса заявки на текующую
        {
            if (e.RowIndex < dataGridView2.Rows.Count - 1)
            {
                string kod = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult result = MessageBox.Show(
                       "Изменить статус данной\nзаявки на текущую?",
                       "Сообщение",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Information,
                       MessageBoxDefaultButton.Button1,
                       MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    proverka_perenosa = true;
                    Form5 form = new Form5(kod);
                    form.textBox1.Text = "" + Convert.ToString(kod);
                    form.ShowDialog();
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            int kodorg = 0;
            int kodsotr = 0;
            if (comboBox11.SelectedIndex != -1)
            {
                int index = comboBox11.SelectedIndex;
                kodorg = Convert.ToInt32(comboBox7.Items[index].ToString());
            }
            if (comboBox12.SelectedIndex != -1)
            {
                int index = comboBox12.SelectedIndex;
                kodsotr = Convert.ToInt32(comboBox13.Items[index].ToString());
            }
            dateTimePicker7.CustomFormat = "dd.MM.yyyy";
            dateTimePicker8.CustomFormat = "dd.MM.yyyy";
            if (checkBox7.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox8.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox9.Checked) заявкиBindingSource.Filter = "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox10.Checked) заявкиBindingSource.Filter = "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox8.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox9.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox8.Checked && checkBox9.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox8.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox9.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox8.Checked && checkBox9.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox8.Checked && checkBox9.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Код_организации=" + "'" + kodorg + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox9.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox8.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            if (checkBox7.Checked && checkBox8.Checked && checkBox9.Checked && checkBox10.Checked) заявкиBindingSource.Filter = "Тип_заявки=" + "'" + comboBox3.Text + "'" + "AND " + "Код_организации=" + "'" + kodorg + "'" + "AND " + "Дата_завершения >= " + "'" + dateTimePicker8.Value.Date + "'" + "AND " + "Дата_завершения <= " + "'" + dateTimePicker7.Value.Date + "'" + "AND " + "Код_сотрудника=" + "'" + kodsotr + "'" + "AND " + "Статус like '" + status_zavershen + "'";
            label44.Text ="Количество выполненных заявок: "+ Convert.ToString(dataGridView9.Rows.Count);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1250, BaseFont.EMBEDDED);
            PdfPTable pdftable = new PdfPTable(dataGridView9.Columns.Count);;
            pdftable.DefaultCell.Padding = 3;
            pdftable.WidthPercentage = 100;
            pdftable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdftable.DefaultCell.BorderWidth = 1;
            string FONT_LOCATION = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "times.ttf");
            BaseFont baseFont = BaseFont.CreateFont(FONT_LOCATION, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font text = new iTextSharp.text.Font(baseFont, 8, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font text1 = new iTextSharp.text.Font(baseFont, 11, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font text2 = new iTextSharp.text.Font(baseFont, 3, iTextSharp.text.Font.NORMAL);
            foreach (DataGridViewColumn column in dataGridView9.Columns)
            {  
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText, text));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);             
                pdftable.AddCell(cell);
            }
            
            foreach (DataGridViewRow row in dataGridView9.Rows)
            {
                int i = 0;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    
                    if(i==1 || i==9)
                    {
                        string date = cell.Value.ToString();
                        string [] date1=date.Split(' ');
                        pdftable.AddCell(new Phrase(date1[0], text));                    
                    }
                   
                    else
                    pdftable.AddCell(new Phrase(cell.Value.ToString(), text));
                    i += 1;
                }

            }
            var savefiledialoge = new SaveFileDialog();
            savefiledialoge.FileName = "Отчетность";
            savefiledialoge.DefaultExt = ".pdf";
            savefiledialoge.Filter = "PDF (*.pdf)|*.pdf";
            if (savefiledialoge.ShowDialog() == DialogResult.OK)
            {
                using (FileStream stream = new FileStream(savefiledialoge.FileName, FileMode.Create))
                {
                    Document pdfdoc = new Document(PageSize.A4.Rotate(), 10f, 10f, 10f, 0f);
                    PdfWriter.GetInstance(pdfdoc, stream);
                    pdfdoc.Open();
                    iTextSharp.text.Font titleFont = FontFactory.GetFont("Arial", 12);
                    Paragraph title;
                    Paragraph probel;
                    probel = new Paragraph("          ",text2);
                    title = new Paragraph("Отчетность по завершенным заявкам", text1);
                    title.Alignment = Element.ALIGN_CENTER;
                    pdfdoc.Add(title);
                    pdfdoc.Add(probel);
                    pdfdoc.Add(pdftable);
                    pdfdoc.Close();
                    stream.Close();
                    System.Diagnostics.Process.Start(Convert.ToString(savefiledialoge.FileName));
                }
            }
        }

        private async void button8_Click_1(object sender, EventArgs e)
        {
            if (pr == true)
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Организации] WHERE Код_организации=@Код_организации", sqlConnection);
                command.Parameters.AddWithValue("Код_организации", Convert.ToString(org));
                await command.ExecuteNonQueryAsync();
                this.организацииTableAdapter.Fill(this.u1666130_JKH34DataSet1.Организации);
                MessageBox.Show("Успешно удалено!");
                pr = false;
                combUPorg();
            }
            else MessageBox.Show("Вы еще не добавили организацию\n чтобы ее удалить!");
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            сотрудникипредприятияBindingSource.Filter = "Статус like '" + status_yvolt + "'";
            string text = Convert.ToString(comboBox12.Text);
            int chet = dataGridView6.RowCount - 1;
            for (int ch = chet; ch > -1; ch--)
            {
                string arhiv = Convert.ToString(dataGridView6[1, ch].Value);

                if (arhiv == text)
                {
                    string s = "Сотрудник уже не работает";
                    toolTip1.SetToolTip(comboBox12,s);
                    break;
                }
                else
                {
                    toolTip1.RemoveAll();
                    toolTip1.Hide(comboBox12);
                }
            }
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            организацииBindingSource.Filter = "Статус like '" + status_nedeyst + "'";
            int chet = dataGridView5.RowCount - 1;
            for (int ch = chet; ch > -1; ch--)
            {
                string arhiv = Convert.ToString(dataGridView5[2, ch].Value);
                string text = Convert.ToString(comboBox11.Text);
                if (arhiv == text)
                {
                    toolTip1.SetToolTip(comboBox11, "Организация уже не действительна");
                    break;
                }
                else
                {
                    toolTip1.RemoveAll();
                    toolTip1.Hide(comboBox11);
                }

            }
        }
    }
}
