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
        public MainForm() {
            InitializeComponent();
            CenterToScreen();
        }

        private void exitButton_Click(object sender, EventArgs e) {
            Close();
        }

        private void openFileButton_Click(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                //MessageBox.Show(openFileDialog.FileNames[0]);
                Converter converter = new Converter();
                converter.Convert(openFileDialog.FileNames[0], progressBar);
            }
        }
    }
}
