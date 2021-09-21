using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace EduCenter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            DB.db = new cSQL();
            DB.db.cSQL_init(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=" + AppDomain.CurrentDomain.BaseDirectory + "MyBase.mdf" + ";Integrated Security=True;Connect Timeout=30");
            DB.db.Connect();
            
            
            textBox2.PasswordChar = '*';
        }
        public void Ini()
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(textBox2.Text);
            string Login = textBox1.Text;
            string Password = textBox2.Text;
            try
            {
                StreamReader sr = new StreamReader("parol.txt");
                var log = sr.ReadLine();
                var pass = sr.ReadLine();
                sr.Close();
                if (Login == log && Password == pass)
                {
                    Registratsiya reg = new Registratsiya();
                    this.Hide();
                    reg.ShowDialog();
                    this.Show();
                }
                else
                {
                    MessageBox.Show("Login yoki parol xato. Qayta urinib ko'ring!");
                }
            }
            catch (Exception)
            {
                StreamWriter sw = new StreamWriter("parol.txt");
                sw.WriteLine(Login);
                sw.WriteLine(Password);
                sw.Close();
                Registratsiya reg = new Registratsiya();
                this.Hide();
                reg.ShowDialog();
                this.Show();
            }
            
            

        }

        
    }
}
