using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sayaboc
{
    public partial class Dashboard: Form
    {
       
        Form2 f2 = new Form2();
        public Dashboard()
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(DisplayIt.ProfilePath) && File.Exists(DisplayIt.ProfilePath))
            {
                pictureBox1.Image = Image.FromFile(DisplayIt.ProfilePath);
                pictureBox1.Image = Image.FromFile(DisplayIt.ProfilePath);
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage; // Set the SizeMode to StretchImage
            }
            lblActiveStudentCount.Text = showCount(12, "1").ToString();
            lblInactiveStudentCount.Text= showCount(12, "0").ToString();
            lblMaleCount.Text= showCount(2, "MALE").ToString();
            lblFemaleCount.Text = showCount(2, "FEMALE").ToString();
            lblChessCount.Text = showCount(3, "CHESS").ToString();
            lblGamingCount.Text = showCount(3, "GAMES").ToString();
            lblCyclingCount.Text = showCount(3, "CYCLING").ToString();
            lblRedCount.Text = showCount(4, "Red").ToString();
            lblGreenCount.Text = showCount(4, "Green").ToString();
            lblBlueCount.Text = showCount(4, "Blue").ToString();
            lblYellowCount.Text = showCount(4, "Yellow").ToString();
            lblBSITCount.Text = showCount(13, "BSIT").ToString();
            lblBSTMCount.Text = showCount(13, "BSTM").ToString();

            lblNickname.Text = DisplayIt.DisplayName;
        }
        public int showCount(int c, string field)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sh = book.Worksheets[0];
            int counter = 0;
            int row = sh.Rows.Length;
            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, c].Value == field)
                {
                    counter++;
                }
            }
            return counter;
        }


            private void button3_Click(object sender, EventArgs e)
        {
            f2.loadLogs();
            f2.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            this.Hide();
            f1.Show();
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {

        }

        private void btnActiveStatus_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.showStudent("1");
            f2.Show();
            this.Hide();
        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.showStudent("0");
            f2.Show();
            this.Hide();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();
            logs.insertLogs(DisplayIt.CurrentUser, "Logged Out.");
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Hide();
        }

        private void lblNickname_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.txtName.Text = lblNickname.Text;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
