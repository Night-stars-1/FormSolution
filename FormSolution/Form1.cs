using System;
using System.Windows.Forms;

namespace FormSolution
{

    public partial class Form1 : Form
    {

        public string InputText { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.InputText = this.textBox1.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 将所有\n替换为\r\n
            textBox1.Text = textBox1.Text.Replace("\r\n", "\n").Replace("\n", "\r\n");
        }

    }
}
