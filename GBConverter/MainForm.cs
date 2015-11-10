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
            Console.Write("Super!");
            /*
            string[] args = Environment.GetCommandLineArgs();
            foreach (string s in args) {
                Console.WriteLine(s);
            }
             * */
        }

        private void exitButton_Click(object sender, EventArgs e) {
            Close();
        }

        private void openFileButton_Click(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                Converter converter = new Converter();
                converter.Convert(openFileDialog.FileNames[0], progressBar);
            }
        }
    }
}
