using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sayaboc
{
    
    public partial class Form2 : Form
    {
       
       
        public void insert(string name,string gender,string hobbies,string color,string saying)
        {
              int i = dataGridView1.Rows.Add();
            dataGridView1.Rows[i].Cells[0].Value = name;
            dataGridView1.Rows[i].Cells[1].Value = gender;
            dataGridView1.Rows[i].Cells[2].Value = hobbies;
            dataGridView1.Rows[i].Cells[3].Value = color;
            dataGridView1.Rows[i].Cells[4].Value = saying;


        }
        public void update(int id,string name, string gender, string hobbies, string color, string saying)
        {
            
            dataGridView1.Rows[id].Cells[0].Value = name;
            dataGridView1.Rows[id].Cells[1].Value = gender;
            dataGridView1.Rows[id].Cells[2].Value = hobbies;
            dataGridView1.Rows[id].Cells[3].Value = color;
            dataGridView1.Rows[id].Cells[4].Value = saying;


        }

        public Form2()
        {
            InitializeComponent();
            LoadExcelFile();
        }
        private int GetSelectedRow()
        {

            if (dataGridView1.SelectedCells.Count > 0)
            {

                int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                return selectedRowIndex;
            }
            return -1;
        }
        public void loadLogs()
        {
            MyLogs logs = new MyLogs();
            logs.showLogs(dataGridView1);
        }
        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dataGridView1.DataSource = dt; 

        }
         public void showStudent(string status)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sh = book.Worksheets[0];
            DataTable dt = sh.ExportDataTable();
            DataTable filteredTable = dt.Clone();
            DataRow[] row = dt.Select("Status = " + status);

            foreach (DataRow r in row)
            {
                filteredTable.ImportRow(r);

                
            }
            dataGridView1.DataSource = filteredTable;
        }
        
        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
           
            DialogResult result = MessageBox.Show("Are you sure you want to delete", "Notice",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
                Worksheet sh = book.Worksheets[0];
                int row = dataGridView1.CurrentCell.RowIndex + 2;

                sh.Range[row, 13].Value = "0";
                this.showStudent("1");
                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx", ExcelVersion.Version2016);
            }
            
            
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            Dashboard frm4 = new Dashboard();
            Form1 frm1 = new Form1();

            int r = dataGridView1.CurrentCell.RowIndex;

            frm1.lblInfo.Text = r.ToString();
            string name = dataGridView1.Rows[r].Cells[0].Value.ToString();
            string gender = dataGridView1.Rows[r].Cells[1].Value.ToString();
            string address = dataGridView1.Rows[r].Cells[2].Value.ToString();
            string email = dataGridView1.Rows[r].Cells[3].Value.ToString();
            string birthday = dataGridView1.Rows[r].Cells[4].Value.ToString();
            string age = dataGridView1.Rows[r].Cells[5].Value.ToString();
            string user = dataGridView1.Rows[r].Cells[6].Value.ToString();
            string pass = dataGridView1.Rows[r].Cells[7].Value.ToString();
            string hobbies = dataGridView1.Rows[r].Cells[8].Value.ToString();
            string favColor = dataGridView1.Rows[r].Cells[9].Value.ToString();
            string saying = dataGridView1.Rows[r].Cells[10].Value.ToString();
            string course = dataGridView1.Rows[r].Cells[11].Value.ToString();
            string status = dataGridView1.Rows[r].Cells[12].Value.ToString();
            string profile = dataGridView1.Rows[r].Cells[13].Value.ToString();

            profile = frm4.lblProfPathHolder.Text;


            frm1.UpdateTextFields(r, name, gender, hobbies, address, email, birthday, age, favColor, user, pass, saying, course, status, profile);
            frm1.btnAdd.Visible = false;
            frm1.btnBrowse.Visible = false;
            frm1.lblProfile.Visible = false;
            frm1.txtBrowse.Visible = false;
            frm1.btnUpdate.Visible = true;
            frm1.Show();
            this.Hide();

            //Form1 f1 = new Form1();
            //f1= (Form1)Application.OpenForms["Form1"];
            //int r=dataGridView1.CurrentCell.RowIndex;
            //f1.lbl.Text=r.ToString();
            //f1.txtName.Text = dataGridView1.Rows[r].Cells[0].Value.ToString();
            //string gender =dataGridView1.Rows[r].Cells[1].Value.ToString();
            //if(gender =="MALE")
            //{
            //    f1.rdoMale.Checked = true;
            //}

            //if (gender == "FEMALE")
            //{
            //    f1.rdoFemale.Checked = true;
            //}
            //string hobby=dataGridView1.Rows[r].Cells[2].Value.ToString();
            //if(hobby=="CHESS")
            //{
            //    f1.chkChess.Checked = true;
            //}
            //if(hobby=="GAMES")
            //{
            //    f1.chkMobile.Checked = true;
            //}
            //if(hobby=="CYCLING")
            //{
            //    f1.chkCycling.Checked = true;
            //}



            //f1.cmbColor.Text = dataGridView1.Rows[r].Cells[3].Value.ToString();
            //f1.txtSaying.Text = dataGridView1.Rows[r].Cells[4].Value.ToString();
            //f1.txtUsername.Text = dataGridView1.Rows[r].Cells[5].Value.ToString();
            //f1.txtPassword.Text = dataGridView1.Rows[r].Cells[6].Value.ToString();
            //f1.txtEmail.Text = dataGridView1.Rows[r].Cells[7].Value.ToString();
            //f1.dateTimePicker1.Text = dataGridView1.Rows[r].Cells[8].Value.ToString();
            //f1.txtAge.Text = dataGridView1.Rows[r].Cells[9].Value.ToString();
            //f1.txtAddress.Text = dataGridView1.Rows[r].Cells[10].Value.ToString();
            //f1.txtStatus.Text = dataGridView1.Rows[r].Cells[11].Value.ToString();

            //f1.btnUpdate.Visible = true ;
            //f1.btnAdd.Visible = false ;

            //f1.Show();



        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();
            logs.insertLogs(DisplayIt.CurrentUser, "Searched in the active list.");
            string searchText = txtSearch.Text.ToLower();
            bool foundMatch = false;

            if (string.IsNullOrEmpty(txtSearch.Text))
            {
                MessageBox.Show("Please enter the cell you want to search.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }



            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Split(' ').Contains(searchText))
                    {
                        cell.Style.BackColor = Color.Yellow;
                        foundMatch = true;
                    }
                    else
                    {
                        cell.Style.BackColor = dataGridView1.DefaultCellStyle.BackColor;
                    }
                }
            }

            if (foundMatch)
            {
                MessageBox.Show("Matching cells have been highlighted.", "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("No matching cells found.", "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.ToLower();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = dataGridView1.DefaultCellStyle.BackColor;
                }
            }
        }
        public void UpdateToExcel(int ID, string name, string gender, string hobbies, string address, string email, string birthday, string age, string favColor, string user, string pass, string saying, string course, string status, string profile)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int id = ID + 2;
            sheet.Range[id, 1].Value = name;
            sheet.Range[id, 2].Value = gender;
            sheet.Range[id, 11].Value = address;
            sheet.Range[id, 8].Value = email;
            sheet.Range[id, 9].Value = birthday;
            sheet.Range[id, 10].Value = age;
            sheet.Range[id, 6].Value = user;
            sheet.Range[id, 7].Value = pass;
            sheet.Range[id, 3].Value = hobbies;
            sheet.Range[id, 4].Value = favColor;
            sheet.Range[id, 5].Value = saying;
            sheet.Range[id, 13].Value = course;
            sheet.Range[id, 12].Value = status;
            sheet.Range[id, 14].Value = profile;

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");

            int dgvIndex = ID;
            dataGridView1.Rows[dgvIndex].Cells[0].Value = name;
            dataGridView1.Rows[dgvIndex].Cells[1].Value = gender;
            dataGridView1.Rows[dgvIndex].Cells[10].Value = address;
            dataGridView1.Rows[dgvIndex].Cells[7].Value = email;
            dataGridView1.Rows[dgvIndex].Cells[8].Value = birthday;
            dataGridView1.Rows[dgvIndex].Cells[9].Value = age;
            dataGridView1.Rows[dgvIndex].Cells[5].Value = user;
            dataGridView1.Rows[dgvIndex].Cells[6].Value = pass;
            dataGridView1.Rows[dgvIndex].Cells[2].Value = hobbies;
            dataGridView1.Rows[dgvIndex].Cells[3].Value = favColor;
            dataGridView1.Rows[dgvIndex].Cells[4].Value = saying;
            dataGridView1.Rows[dgvIndex].Cells[12].Value = course;
            dataGridView1.Rows[dgvIndex].Cells[11].Value = status;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
        }

        private void btnDeleteLogs_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete the selected info?", "Notice", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {

                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
                Worksheet sh = book.Worksheets[0];

                int row = dataGridView1.CurrentCell.RowIndex + 2;

                sh.DeleteRow(row);

                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx", ExcelVersion.Version2013);
            }
        }
    }
}
