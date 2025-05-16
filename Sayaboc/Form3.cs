using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sayaboc
{
    public partial class Form3: Form
    {
        MyLogs log = new MyLogs();
        Dashboard d = new Dashboard();
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            log.insertLogs(txtUsername.Text, "Success");
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = sheet.Rows.Length;
            bool logs = false;
            if (string.IsNullOrEmpty(txtUsername.Text) || string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("Required fields!", "Notice!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                for (int i = 2; i <= row; i++)
                {

                    if (sheet.Range[i, 6].Value == txtUsername.Text && sheet.Range[i, 7].Value == txtPassword.Text)
                    {
                        DisplayIt.CurrentUser = txtUsername.Text;
                        DisplayIt.DisplayName = sheet.Range[i, 1].Value;
                        DisplayIt.ProfilePath = sheet.Range[i, 14].Value;
                        log.insertLogs(txtUsername.Text, txtUsername.Text + " logged in");
                        logs = true;
                        break;

                    }
                    else
                    {
                        logs = false;
                    }
                }
            }
            if (logs == true)
            {

                MessageBox.Show("Success", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Dashboard f1 = new Dashboard();
                f1.Show();
                this.Hide();

            }

            else
            {
                MessageBox.Show("Invalid Information", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }
        
        

        private void Form3_Load(object sender, EventArgs e)
        {

        }
    }
}
