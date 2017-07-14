namespace IrbisMoto
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
            this.rtbMiniText = new System.Windows.Forms.RichTextBox();
            this.rtbFullText = new System.Windows.Forms.RichTextBox();
            this.tbTitle = new System.Windows.Forms.TextBox();
            this.tbDescription = new System.Windows.Forms.TextBox();
            this.tbKeywords = new System.Windows.Forms.TextBox();
            this.btnSaveTemplates = new System.Windows.Forms.Button();
            this.btnActual = new System.Windows.Forms.Button();
            this.btnUpdateImage = new System.Windows.Forms.Button();
            this.tbLogin = new System.Windows.Forms.TextBox();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.lblLogin = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.cbMiniText = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // rtbMiniText
            // 
            this.rtbMiniText.Location = new System.Drawing.Point(12, 12);
            this.rtbMiniText.Name = "rtbMiniText";
            this.rtbMiniText.Size = new System.Drawing.Size(596, 159);
            this.rtbMiniText.TabIndex = 0;
            this.rtbMiniText.Text = "";
            // 
            // rtbFullText
            // 
            this.rtbFullText.Location = new System.Drawing.Point(12, 191);
            this.rtbFullText.Name = "rtbFullText";
            this.rtbFullText.Size = new System.Drawing.Size(596, 146);
            this.rtbFullText.TabIndex = 1;
            this.rtbFullText.Text = "";
            // 
            // tbTitle
            // 
            this.tbTitle.Location = new System.Drawing.Point(12, 353);
            this.tbTitle.Name = "tbTitle";
            this.tbTitle.Size = new System.Drawing.Size(596, 20);
            this.tbTitle.TabIndex = 2;
            // 
            // tbDescription
            // 
            this.tbDescription.Location = new System.Drawing.Point(12, 388);
            this.tbDescription.Name = "tbDescription";
            this.tbDescription.Size = new System.Drawing.Size(596, 20);
            this.tbDescription.TabIndex = 3;
            // 
            // tbKeywords
            // 
            this.tbKeywords.Location = new System.Drawing.Point(12, 424);
            this.tbKeywords.Name = "tbKeywords";
            this.tbKeywords.Size = new System.Drawing.Size(596, 20);
            this.tbKeywords.TabIndex = 4;
            // 
            // btnSaveTemplates
            // 
            this.btnSaveTemplates.Location = new System.Drawing.Point(614, 151);
            this.btnSaveTemplates.Name = "btnSaveTemplates";
            this.btnSaveTemplates.Size = new System.Drawing.Size(161, 28);
            this.btnSaveTemplates.TabIndex = 5;
            this.btnSaveTemplates.Text = "Сохранить шаблон";
            this.btnSaveTemplates.UseVisualStyleBackColor = true;
            this.btnSaveTemplates.Click += new System.EventHandler(this.btnSaveTemplates_Click);
            // 
            // btnActual
            // 
            this.btnActual.Location = new System.Drawing.Point(613, 12);
            this.btnActual.Name = "btnActual";
            this.btnActual.Size = new System.Drawing.Size(162, 27);
            this.btnActual.TabIndex = 6;
            this.btnActual.Text = "Запчасти";
            this.btnActual.UseVisualStyleBackColor = true;
            this.btnActual.Click += new System.EventHandler(this.btnActual_Click);
            // 
            // btnUpdateImage
            // 
            this.btnUpdateImage.Location = new System.Drawing.Point(614, 78);
            this.btnUpdateImage.Name = "btnUpdateImage";
            this.btnUpdateImage.Size = new System.Drawing.Size(161, 28);
            this.btnUpdateImage.TabIndex = 7;
            this.btnUpdateImage.Text = "Обновить картинки";
            this.btnUpdateImage.UseVisualStyleBackColor = true;
            this.btnUpdateImage.Click += new System.EventHandler(this.btnUpdateImage_Click);
            // 
            // tbLogin
            // 
            this.tbLogin.Location = new System.Drawing.Point(614, 125);
            this.tbLogin.Name = "tbLogin";
            this.tbLogin.Size = new System.Drawing.Size(75, 20);
            this.tbLogin.TabIndex = 8;
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(700, 125);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.Size = new System.Drawing.Size(75, 20);
            this.tbPassword.TabIndex = 9;
            this.tbPassword.UseSystemPasswordChar = true;
            // 
            // lblLogin
            // 
            this.lblLogin.AutoSize = true;
            this.lblLogin.Location = new System.Drawing.Point(614, 109);
            this.lblLogin.Name = "lblLogin";
            this.lblLogin.Size = new System.Drawing.Size(38, 13);
            this.lblLogin.TabIndex = 10;
            this.lblLogin.Text = "Логин";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(697, 109);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(45, 13);
            this.lblPassword.TabIndex = 11;
            this.lblPassword.Text = "Пароль";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(613, 45);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(162, 27);
            this.button1.TabIndex = 12;
            this.button1.Text = "Запчасти для снегоходов";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // cbMiniText
            // 
            this.cbMiniText.AutoSize = true;
            this.cbMiniText.Location = new System.Drawing.Point(614, 185);
            this.cbMiniText.Name = "cbMiniText";
            this.cbMiniText.Size = new System.Drawing.Size(170, 17);
            this.cbMiniText.TabIndex = 13;
            this.cbMiniText.Text = "Обновить краткое описание";
            this.cbMiniText.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 450);
            this.Controls.Add(this.cbMiniText);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblLogin);
            this.Controls.Add(this.tbPassword);
            this.Controls.Add(this.tbLogin);
            this.Controls.Add(this.btnUpdateImage);
            this.Controls.Add(this.btnActual);
            this.Controls.Add(this.btnSaveTemplates);
            this.Controls.Add(this.tbKeywords);
            this.Controls.Add(this.tbDescription);
            this.Controls.Add(this.tbTitle);
            this.Controls.Add(this.rtbFullText);
            this.Controls.Add(this.rtbMiniText);
            this.Name = "Form1";
            this.Text = "Irbis Moto";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbMiniText;
        private System.Windows.Forms.RichTextBox rtbFullText;
        private System.Windows.Forms.TextBox tbTitle;
        private System.Windows.Forms.TextBox tbDescription;
        private System.Windows.Forms.TextBox tbKeywords;
        private System.Windows.Forms.Button btnSaveTemplates;
        private System.Windows.Forms.Button btnActual;
        private System.Windows.Forms.Button btnUpdateImage;
        private System.Windows.Forms.TextBox tbLogin;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.Label lblLogin;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox cbMiniText;
    }
}

