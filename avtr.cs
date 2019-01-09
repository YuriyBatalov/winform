using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace version04
{
    public partial class avtr : Form
    {
        public avtr()
        {
            InitializeComponent();
            textBox1.Text = "Admin";
            textBox2.Text = "123";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //SqlConnection cn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename='C: \\Users\\134217\\Google Диск\\УЧЕБА\\Семестр №8\\ДИПЛОМ\\version04\\Molotdel.mdf';Initial Catalog=;Integrated Security=True;Connect Timeout=30");
            if (textBox1.Text == "Admin" & textBox2.Text == "123")
            {
                Hide();
                Form1 newForm = new Form1(); newForm.Show();
            }
            else { MessageBox.Show("Данные введены не верно!"); }

        }
    }
}
