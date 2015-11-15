using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GBConverter {
    public partial class MainForm : Form {
        private Converter converter;

        public MainForm() {
            InitializeComponent();
            CenterToScreen();
        }

        private void ExitButton_Click(object sender, EventArgs e) {
            Close();
        }

        private void OpenFileButton_Click(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                this.converter = new Converter();
                if (this.converter.CheckFile(openFileDialog.FileNames[0], progressBar)) {
                    MessageBox.Show("Несоответствий не выявлено.", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.ConvertButton.Enabled = true;
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
            DB.CloseConnection();
        }

        private void ConvertButton_Click(object sender, EventArgs e) {
            try {
                if (this.converter.Convert(progressBar)) {
                    MessageBox.Show("Конвертация завершена.", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("Во время конвертации возникла ошибка.", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            } catch (Exception ex) {
                string ExtraInfo = "";
                if( ex.Data.Contains("UserMessage") ) {
                    ExtraInfo = "\n" + ex.Data["UserMessage"].ToString();
                }
                MessageBox.Show("Во время конвертации возникла ошибка." + ExtraInfo + "\n\n" + ex.Message, "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void RollbackButton_Click(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Журналы конвертации|*.log";
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                this.converter = new Converter();
                try {
                    long DeletedAppeals = this.converter.Rollback(openFileDialog.FileNames[0], progressBar);
                    if (DeletedAppeals >= 0) {
                        MessageBox.Show("Откат завершён.\nУдалено заявок: " + DeletedAppeals.ToString() + ".", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } else {
                        MessageBox.Show("Во время отката возникла ошибка.", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                } catch (Exception ex) {
                    MessageBox.Show("Во время отката возникла ошибка." + ex.Message, "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }
    }
}
