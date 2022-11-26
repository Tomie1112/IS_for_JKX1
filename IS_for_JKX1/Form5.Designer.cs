namespace IS_for_JKX1
{
    partial class Form5
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form5));
            this.label24 = new System.Windows.Forms.Label();
            this.повторноеоткрытиезаявокBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.u1666130_JKH34DataSet4 = new IS_for_JKX1.u1666130_JKH34DataSet4();
            this.повторное_открытие_заявокTableAdapter = new IS_for_JKX1.u1666130_JKH34DataSet4TableAdapters.Повторное_открытие_заявокTableAdapter();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.button7 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.номерповторногооткрытиязаявкиDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.номерзаявкиDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.повторноеоткрытиезаявокBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.u1666130_JKH34DataSet4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Book Antiqua", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label24.Location = new System.Drawing.Point(24, 9);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(398, 21);
            this.label24.TabIndex = 7;
            this.label24.Text = "Укажите причину повторного открытия заявки:";
            // 
            // повторноеоткрытиезаявокBindingSource
            // 
            this.повторноеоткрытиезаявокBindingSource.DataMember = "Повторное_открытие_заявок";
            this.повторноеоткрытиезаявокBindingSource.DataSource = this.u1666130_JKH34DataSet4;
            // 
            // u1666130_JKH34DataSet4
            // 
            this.u1666130_JKH34DataSet4.DataSetName = "u1666130_JKH34DataSet4";
            this.u1666130_JKH34DataSet4.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // повторное_открытие_заявокTableAdapter
            // 
            this.повторное_открытие_заявокTableAdapter.ClearBeforeFill = true;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Book Antiqua", 10F, System.Drawing.FontStyle.Bold);
            this.richTextBox1.Location = new System.Drawing.Point(17, 34);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(405, 114);
            this.richTextBox1.TabIndex = 9;
            this.richTextBox1.Text = "";
            // 
            // button7
            // 
            this.button7.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button7.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button7.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button7.Location = new System.Drawing.Point(17, 154);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(191, 29);
            this.button7.TabIndex = 22;
            this.button7.Text = "ОК";
            this.button7.UseVisualStyleBackColor = false;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(231, 154);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(191, 29);
            this.button1.TabIndex = 23;
            this.button1.Text = "ОТМЕНА";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(8, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(10, 20);
            this.textBox1.TabIndex = 24;
            this.textBox1.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.номерповторногооткрытиязаявкиDataGridViewTextBoxColumn,
            this.номерзаявкиDataGridViewTextBoxColumn,
            this.причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.повторноеоткрытиезаявокBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(8, 38);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(10, 33);
            this.dataGridView1.TabIndex = 25;
            this.dataGridView1.Visible = false;
            // 
            // номерповторногооткрытиязаявкиDataGridViewTextBoxColumn
            // 
            this.номерповторногооткрытиязаявкиDataGridViewTextBoxColumn.DataPropertyName = "Номер_повторного_открытия_заявки";
            this.номерповторногооткрытиязаявкиDataGridViewTextBoxColumn.HeaderText = "Номер_повторного_открытия_заявки";
            this.номерповторногооткрытиязаявкиDataGridViewTextBoxColumn.Name = "номерповторногооткрытиязаявкиDataGridViewTextBoxColumn";
            // 
            // номерзаявкиDataGridViewTextBoxColumn
            // 
            this.номерзаявкиDataGridViewTextBoxColumn.DataPropertyName = "Номер_заявки";
            this.номерзаявкиDataGridViewTextBoxColumn.HeaderText = "Номер_заявки";
            this.номерзаявкиDataGridViewTextBoxColumn.Name = "номерзаявкиDataGridViewTextBoxColumn";
            // 
            // причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn
            // 
            this.причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn.DataPropertyName = "Причина_повторного_открытия_заявки";
            this.причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn.HeaderText = "Причина_повторного_открытия_заявки";
            this.причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn.Name = "причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn";
            // 
            // Form5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 197);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.label24);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form5";
            this.Load += new System.EventHandler(this.Form5_Load);
            ((System.ComponentModel.ISupportInitialize)(this.повторноеоткрытиезаявокBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.u1666130_JKH34DataSet4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label24;
        private u1666130_JKH34DataSet4 u1666130_JKH34DataSet4;
        private System.Windows.Forms.BindingSource повторноеоткрытиезаявокBindingSource;
        private u1666130_JKH34DataSet4TableAdapters.Повторное_открытие_заявокTableAdapter повторное_открытие_заявокTableAdapter;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn номерповторногооткрытиязаявкиDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn номерзаявкиDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn причинаповторногооткрытиязаявкиDataGridViewTextBoxColumn;
    }
}