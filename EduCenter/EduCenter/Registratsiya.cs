using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace EduCenter
{
    public partial class Registratsiya : Form
    {
        public Registratsiya()
        {
            InitializeComponent();
        }

        private void Registratsiya_Load(object sender, EventArgs e)
        {
            

            comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
            DataTable dt1 = DB.db.Query("SELECT * FROM Fan");
            //fanlarni olish
            comboBox1.DataSource = dt1;
            comboBox1.DisplayMember = "Nomi";
            comboBox1.ValueMember = "Id";
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            DataTable dt2 = DB.db.Query("SELECT * FROM Fan");
            comboBox3.DataSource = dt2;
            comboBox3.DisplayMember = "Nomi";
            comboBox3.ValueMember = "Id";
            

            DataTable dt3 = DB.db.Query("SELECT * FROM Kun");
            //Daraja olish
            comboBox2.DataSource = dt3;
            comboBox2.DisplayMember = "Kuni";
            comboBox2.ValueMember = "Id";
            DataTable dt4 = DB.db.Query("SELECT * FROM Kun");
            comboBox4.DataSource = dt4;
            comboBox4.DisplayMember = "Kuni";
            comboBox4.ValueMember = "Id";
            refresh();
            loadFan();
        }
        public void refresh()
        {
            dataGridView1.DataSource = DB.db.Query("select Talaba.Familya, Talaba.Ism, Talaba.TelNomer, Fan.Nomi as Fani,Kun.Kuni as Kuni from Talaba, Fan, Kun where Talaba.KunId=Kun.Id and Talaba.FanId=Fan.Id");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string Familya = textBox1.Text;
            string Ism = textBox2.Text;
            string Tel = textBox3.Text;
            long number1 = 0;
            bool canConvert = long.TryParse(Tel, out number1);
            if (Tel.IndexOf('+').ToString() == "-1" && canConvert)
            {
                if (Familya.Length > 0 && Ism.Length > 0 && Tel.Length > 0 && comboBox1.SelectedValue.ToString().Length > 0 && comboBox2.SelectedValue.ToString().Length > 0)
                {
                    if ((int)Familya[0] > 125 &&  (int)Ism[0] > 125)
                    {
                        Familya = Converter.ConvertToLatin(Familya);
                        Ism = Converter.ConvertToLatin(Ism);
                    }
                    if (Familya != "0" && Ism != "0")
                    {
                        Familya = Familya.Replace("'", "`");
                        Ism = Ism.Replace("'", "`");

                        string zapros = "";
                        zapros += " insert into Talaba (Familya,Ism,TelNomer,FanId,KunId) values('" + Familya + "','" + Ism + "','" + Tel + "','" + comboBox1.SelectedValue + "','" + comboBox2.SelectedValue + "' )";
                        // zapros += "insert into AllUsers (Fullname,UsTypeId,Email,Phone,Ish_joyi) values('" + text_fish.Text + "','" + 2 + "','" + text_email.Text + "','" + text_tel.Text + "','" + text_ish_joyi.Text + "')";
                        if (DB.db.SetCommand(zapros) == 1)
                        {
                            MessageBox.Show("Saqlandi");
                            refresh();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Qayta urinib ko'ring!");
                    }
                }
                else
                {
                    MessageBox.Show("F.I.Sh, Telefon, Kurs yoki Jadval bandiga ma'lumot kiritilmadi. Ma'lumotni tekshirib qayta urinib ko'ring!");

                }
            }
            else
            {
                MessageBox.Show("Telefon nomerni quydagi ko'rinishlardan birida kiriting! Namuna: 998941234567 yoki 941234567");

            }


        }
        public void loadFan()
        {
            dataGridView3.DataSource = DB.db.Query("select Nomi from Fan");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(textBox4.Text.Length>0)
            {
                string fan = textBox4.Text.ToString().Replace("'", "`");
                string zapros = "";
                zapros += " insert into Fan (Nomi) values('" + fan + "' )";
                // zapros += "insert into AllUsers (Fullname,UsTypeId,Email,Phone,Ish_joyi) values('" + text_fish.Text + "','" + 2 + "','" + text_email.Text + "','" + text_tel.Text + "','" + text_ish_joyi.Text + "')";
                if (DB.db.SetCommand(zapros) == 1)
                {
                    MessageBox.Show("Saqlandi");
                    textBox4.Text = "";
                    loadFan();
                }

            }
            else
            {
                MessageBox.Show("Fan nomi kiritilmadi");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
                string id_del = dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells["Nomi"].Value.ToString();
                string zapros = "";    
                zapros += "delete from Fan where Nomi= '" + id_del + "'";
                if (DB.db.SetCommand(zapros) == 1)
                {
                    MessageBox.Show("O'chirildi");
                }
                loadFan();
        }

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

           // MessageBox.Show(dataGridView3.CurrentRow.Index.ToString());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string del_Fam = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Familya"].Value.ToString();
            string del_Ism = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Ism"].Value.ToString();
            string zapros = "";
            zapros += "delete from Talaba where Familya= '" + del_Fam + "' and Ism ='"+del_Ism+"'";
            if (DB.db.SetCommand(zapros) == 1)
            {
                MessageBox.Show("O'chirildi");
            }
            refresh();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string fam = "Familya like '%" + textBox5.Text + "%'";
            string ism = "Ism like '%" + textBox5.Text + "%'";

            dataGridView1.DataSource = DB.db.Query("select Talaba.Familya, Talaba.Ism, Talaba.TelNomer, Fan.Nomi as Fani,Kun.Kuni as Kuni from Talaba, Fan, Kun where Talaba.KunId=Kun.Id and Talaba.FanId=Fan.Id and(" + fam + " or " + ism + ")");
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string fanId = comboBox3.SelectedValue.ToString();
            dataGridView1.DataSource = DB.db.Query("select Talaba.Familya, Talaba.Ism, Talaba.TelNomer, Fan.Nomi as Fani,Kun.Kuni as Kuni from Talaba, Fan, Kun where Talaba.KunId=Kun.Id and Talaba.FanId=Fan.Id and Fan.Id=" + comboBox3.SelectedValue + "");

        }

        private void comboBox4_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dataGridView1.DataSource = DB.db.Query("select Talaba.Familya, Talaba.Ism, Talaba.TelNomer, Fan.Nomi as Fani,Kun.Kuni as Kuni from Talaba, Fan, Kun where Talaba.KunId=Kun.Id and Talaba.FanId=Fan.Id and Kun.Id=" + comboBox4.SelectedValue + "");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;

                for (i = 0; i <= dataGridView1.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView1[j, i];

                        xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                    }
                }
                DateTime time = DateTime.Now;
                string soat = time.Hour.ToString();
                string minu = time.Minute.ToString();
                string secu = time.Second.ToString();

                xlWorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\base-"+soat+minu+secu+ ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                MessageBox.Show("Malumotlar yuklandi");
            }
            catch
            {
                
            }
       
        }
      
    }
    public static class Converter
    {
        private static readonly Dictionary<char, string> ConvertedLetters = new Dictionary<char, string>
    {
        {'Ҳ',"H"},
        {'Ғ',"G‘"},
        {'Қ',"Q"},
        {'Ў',"O‘"},
        {'ғ',"g‘"},
        {'қ',"q"},
        {'ў',"o‘"},
        {'ҳ',"h"},
        {'А',"A"},
        {'Б',"B"},
        {'В',"V"},
        {'Г',"G"},
        {'Д',"D"},
        {'Е',"E"},
        {'Ё',"YO"},
        {'Ж',"J"},
        {'З',"Z"},
        {'И',"I"},
        {'Й',"Y"},
        {'К',"K"},
        {'Л',"L"},
        {'М',"M"},
        {'Н',"N"},
        {'О',"O"},
        {'П',"P"},
        {'Р',"R"},
        {'С',"S"},
        {'Т',"T"},
        {'У',"U"},
        {'Ф',"F"},
        {'Х',"X"},
        {'Ц',"S"},
        {'Ч',"CH"},
        {'Ш',"SH"},
        {'Ъ',"’"},
        {'Э',"E"},
        {'Ю',"YU"},
        {'Я',"YA"},
        {'а',"a"},
        {'б',"b"},
        {'в',"v"},
        {'г',"g"},
        {'д',"d"},
        {'е',"e"},
        {'ё',"yo"},
        {'±',"yo"},
        {'ж',"j"},
        {'з',"z"},
        {'и',"i"},
        {'й',"y"},
        {'к',"k"},
        {'л',"l"},
        {'м',"m"},
        {'н',"n"},
        {'о',"o"},
        {'п',"p"},
        {'р',"r"},
        {'с',"s"},
        {'т',"t"},
        {'у',"u"},
        {'ф',"f"},
        {'х',"x"},
        {'ц',"ts"},
        {'ч',"ch"},
        {'ш',"sh"},
        {'ъ',"'"},
        {'ь',"'"},
        {'э',"e"},
        {'ю',"yu"},
        {'я',"ya"}
    };

        public static string ConvertToLatin(string source)
        {
           
                var result = new StringBuilder();

                try
                {
                    foreach (var letter in source)
                    {
                        result.Append(ConvertedLetters[letter]);
                    }
                    return result.ToString();

                }
                catch
                {
                    MessageBox.Show("Ismi yoki Familyada hatolik bor!");
                }
                return "0";
            
        }
    }
}
