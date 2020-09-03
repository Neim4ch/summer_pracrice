namespace Практика1
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.selectFile = new System.Windows.Forms.Button();
            this.generateXmlButton = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
            this.showAuthorsButton = new System.Windows.Forms.Button();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxShuffle = new System.Windows.Forms.CheckBox();
            this.comboBoxNumeration = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = " ";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 38);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(301, 20);
            this.textBox1.TabIndex = 0;
            // 
            // selectFile
            // 
            this.selectFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.selectFile.FlatAppearance.BorderColor = System.Drawing.Color.Cyan;
            this.selectFile.FlatAppearance.BorderSize = 2;
            this.selectFile.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.selectFile.Location = new System.Drawing.Point(376, 38);
            this.selectFile.Name = "selectFile";
            this.selectFile.Size = new System.Drawing.Size(143, 23);
            this.selectFile.TabIndex = 1;
            this.selectFile.Text = "Выбрать файл";
            this.toolTip1.SetToolTip(this.selectFile, "Файл должен быть формата *.xlsx");
            this.selectFile.UseVisualStyleBackColor = false;
            this.selectFile.Click += new System.EventHandler(this.selectFileButton_Click);
            // 
            // generateXmlButton
            // 
            this.generateXmlButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.generateXmlButton.Location = new System.Drawing.Point(376, 67);
            this.generateXmlButton.Name = "generateXmlButton";
            this.generateXmlButton.Size = new System.Drawing.Size(143, 23);
            this.generateXmlButton.TabIndex = 3;
            this.generateXmlButton.Text = "Сгенерировать xml";
            this.toolTip2.SetToolTip(this.generateXmlButton, "файл *.xml появится в папке рядом с файлом *.xlsx ,названия будут одинаковы");
            this.generateXmlButton.UseVisualStyleBackColor = false;
            this.generateXmlButton.Click += new System.EventHandler(this.generateXmlButton_Click);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog1";
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.White;
            this.progressBar1.ForeColor = System.Drawing.Color.Lime;
            this.progressBar1.Location = new System.Drawing.Point(525, 67);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // progressBar2
            // 
            this.progressBar2.BackColor = System.Drawing.Color.White;
            this.progressBar2.ForeColor = System.Drawing.Color.Lime;
            this.progressBar2.Location = new System.Drawing.Point(525, 38);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(100, 23);
            this.progressBar2.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(255, 15);
            this.label1.TabIndex = 6;
            this.label1.Text = "При наведении на кнопки появляются посказки";
            // 
            // toolTip1
            // 
            this.toolTip1.ShowAlways = true;
            // 
            // toolTip2
            // 
            this.toolTip2.ShowAlways = true;
            // 
            // showAuthorsButton
            // 
            this.showAuthorsButton.Location = new System.Drawing.Point(525, 138);
            this.showAuthorsButton.Name = "showAuthorsButton";
            this.showAuthorsButton.Size = new System.Drawing.Size(100, 21);
            this.showAuthorsButton.TabIndex = 12;
            this.showAuthorsButton.Text = "Об авторах";
            this.showAuthorsButton.UseVisualStyleBackColor = true;
            this.showAuthorsButton.Click += new System.EventHandler(this.showAuthorsButton_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(134, 99);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(115, 60);
            this.button4.TabIndex = 13;
            this.button4.Text = "Ввод ответа, сопоставление, верно \\ неверно";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(12, 99);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(116, 60);
            this.button5.TabIndex = 14;
            this.button5.Text = "Bыбор одного \\ нескольких верных";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Location = new System.Drawing.Point(12, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(142, 15);
            this.label2.TabIndex = 15;
            this.label2.Text = "примеры вопросов в *.xlsx";
            // 
            // checkBoxShuffle
            // 
            this.checkBoxShuffle.AutoSize = true;
            this.checkBoxShuffle.Location = new System.Drawing.Point(376, 99);
            this.checkBoxShuffle.Name = "checkBoxShuffle";
            this.checkBoxShuffle.Size = new System.Drawing.Size(138, 17);
            this.checkBoxShuffle.TabIndex = 16;
            this.checkBoxShuffle.Text = "Перемешать вопросы";
            this.checkBoxShuffle.UseVisualStyleBackColor = true;
            this.checkBoxShuffle.CheckedChanged += new System.EventHandler(this.checkBoxShuffle_CheckedChanged);
            // 
            // comboBoxNumeration
            // 
            this.comboBoxNumeration.AutoCompleteCustomSource.AddRange(new string[] {
            "a, b, c, ...",
            "A, B, C, ...",
            "1, 2, 3, ...",
            "I, II, III, ...",
            "i, ii, iii, ...",
            "Не нумеровать"});
            this.comboBoxNumeration.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxNumeration.FormattingEnabled = true;
            this.comboBoxNumeration.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.comboBoxNumeration.Items.AddRange(new object[] {
            "a, b, c, ...",
            "A, B, C, ...",
            "1, 2, 3, ...",
            "I, II, III, ...",
            "i, ii, iii, ...",
            "Не нумеровать"});
            this.comboBoxNumeration.Location = new System.Drawing.Point(376, 138);
            this.comboBoxNumeration.Name = "comboBoxNumeration";
            this.comboBoxNumeration.Size = new System.Drawing.Size(138, 21);
            this.comboBoxNumeration.TabIndex = 17;
            this.comboBoxNumeration.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(376, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Выбор нумерации: ";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(255, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 60);
            this.button1.TabIndex = 19;
            this.button1.Text = "Выпадающие меню, числовой ";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 168);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxNumeration);
            this.Controls.Add(this.checkBoxShuffle);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.showAuthorsButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.generateXmlButton);
            this.Controls.Add(this.selectFile);
            this.Controls.Add(this.textBox1);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Генератор xml для е-курсов ";
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolTip toolTip2;
        private System.Windows.Forms.Button selectFile;
        private System.Windows.Forms.Button generateXmlButton;
        private System.Windows.Forms.Button showAuthorsButton;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox comboBoxNumeration;
        private System.Windows.Forms.CheckBox checkBoxShuffle;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
    }
}

