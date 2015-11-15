namespace GBConverter {
    partial class MainForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.ExitButton = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.ConvertButton = new System.Windows.Forms.Button();
            this.RollbackButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(12, 110);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(98, 23);
            this.OpenFileButton.TabIndex = 0;
            this.OpenFileButton.Text = "Выбрать файл";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // ExitButton
            // 
            this.ExitButton.Location = new System.Drawing.Point(298, 110);
            this.ExitButton.Name = "ExitButton";
            this.ExitButton.Size = new System.Drawing.Size(72, 23);
            this.ExitButton.TabIndex = 2;
            this.ExitButton.Text = "Выход";
            this.ExitButton.UseVisualStyleBackColor = true;
            this.ExitButton.Click += new System.EventHandler(this.ExitButton_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 52);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(358, 23);
            this.progressBar.TabIndex = 3;
            // 
            // ConvertButton
            // 
            this.ConvertButton.Enabled = false;
            this.ConvertButton.Location = new System.Drawing.Point(116, 110);
            this.ConvertButton.Name = "ConvertButton";
            this.ConvertButton.Size = new System.Drawing.Size(98, 23);
            this.ConvertButton.TabIndex = 4;
            this.ConvertButton.Text = "Конвертировать";
            this.ConvertButton.UseVisualStyleBackColor = true;
            this.ConvertButton.Click += new System.EventHandler(this.ConvertButton_Click);
            // 
            // RollbackButton
            // 
            this.RollbackButton.Location = new System.Drawing.Point(220, 110);
            this.RollbackButton.Name = "RollbackButton";
            this.RollbackButton.Size = new System.Drawing.Size(72, 23);
            this.RollbackButton.TabIndex = 5;
            this.RollbackButton.Text = "Откатить";
            this.RollbackButton.UseVisualStyleBackColor = true;
            this.RollbackButton.Click += new System.EventHandler(this.RollbackButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 145);
            this.Controls.Add(this.RollbackButton);
            this.Controls.Add(this.ConvertButton);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.ExitButton);
            this.Controls.Add(this.OpenFileButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "Конвертер \"Зелёной книги\"";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.Button ExitButton;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button ConvertButton;
        private System.Windows.Forms.Button RollbackButton;
    }
}

