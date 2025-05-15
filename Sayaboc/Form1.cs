using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Sayaboc
{
    public partial class Form1 : Form
    {
        string[] Person = new string[5];
        int i = 0;
        Form2 f2=new Form2();
        public Form1()
        {
            InitializeComponent();
        }
        public string checkEmpty()
        {
            string errors = "Empty fields";
            foreach (Control c in Controls)
            {
                if (c is TextBox)
                {
                    if (c.Text == "")
                    {
                        errors += c.Name + " is empty";
                    }
                }
                if (c is RadioButton)
                {
                    if (c.Text == "")
                    {
                        errors += c.Name + " is empty";
                    }
                }
                if (c is ComboBox)
                {
                    if (c.Text == "")
                    {
                        errors += c.Name + " is empty";
                    }
                }
                
            }
            return errors;
        }
        public bool IsValidEmail(string email)
        {

            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            Regex regex = new Regex(pattern);
            return regex.IsMatch(email);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblInfo.BackColor = Color.FromArgb(30, 50, 50, 100);
            txtAge.Enabled = false;
        }
        public void UpdateTextFields(int ID, string name, string gender, string hobbies, string address, string email, string birthday, string age, string favColor, string user, string pass, string saying, string course, string status, string profile)
        {
            txtName.Text = name;

            ID = Convert.ToInt32(ID);

            if (gender == "MALE")
            {
                rdoMale.Checked = true;
            }
            else if (gender == "FEMALE")
            {
                rdoFemale.Checked = true;
            }

            chkChess.Checked = hobbies.Contains("CHESS");
            chkMobile.Checked = hobbies.Contains("GAMES");
            chkCycling.Checked = hobbies.Contains("CYCLING");


            txtAddress.Text = address;
            txtEmail.Text = email;
            dateTimePicker1.Text = birthday;
            txtAge.Text = age;
            cmbColor.Text = favColor;
            txtUsername.Text = user;
            txtPassword.Text = pass;
            txtSaying.Text = saying;
            cmbCourse.Text = course;
            txtStatus.Text = status;
            lblProfile.Text = profile;
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {

            Workbook book = new Workbook();
            lblerror.Visible = true;

            lblerror.Text = "";
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(txtUsername.Text)) errors.AppendLine("• Username is required.");
            if (string.IsNullOrWhiteSpace(txtPassword.Text)) errors.AppendLine("• Password is required.");
            if (string.IsNullOrWhiteSpace(txtName.Text)) errors.AppendLine("• NickName is required.");
            if (!rdoMale.Checked && !rdoFemale.Checked) errors.AppendLine("• Gender is required.");
            if (string.IsNullOrWhiteSpace(txtAddress.Text)) errors.AppendLine("• Address is required.");
            if (string.IsNullOrWhiteSpace(txtEmail.Text)) errors.AppendLine("• Email is required.");
            if (dateTimePicker1.Checked) errors.AppendLine("• Birthday is required.");
            if (!chkChess.Checked && !chkMobile.Checked && !chkCycling.Checked) errors.AppendLine("• At least one sport must be selected.");
            if (cmbColor.SelectedIndex == -1) errors.AppendLine("• Favorite color must be selected.");
            if (cmbCourse.SelectedIndex == -1) errors.AppendLine("• Course must be selected.");
            if (string.IsNullOrWhiteSpace(txtSaying.Text)) errors.AppendLine("• Saying is required.");
            if (string.IsNullOrWhiteSpace(txtStatus.Text)) errors.AppendLine("• Status is required.");
            if (string.IsNullOrWhiteSpace(txtBrowse.Text)) errors.AppendLine("• Profile is required.");

            DateTime birthDate = dateTimePicker1.Value;
            int calculatedAge = CalculateAge(birthDate);
            txtAge.Text = calculatedAge.ToString();


            if (errors.Length > 0)
            {
                lblerror.Text = errors.ToString();
                lblerror.Visible = true;
                MessageBox.Show("Please fill in all required fields!", "MISSING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //txtName.Clear();
                //txtAddress.Clear();
                //txtEmail.Clear();
                //dtpBirthday.Checked = false;
                //txtUsername.Clear();
                //txtPassword.Clear();
                //cmbFavColor.SelectedIndex = -1;
                //cmbCourse.SelectedIndex = -1;
                //radMale.Checked = false;
                //radFemale.Checked = false;
                //chkBasketball.Checked = false;
                //chkVolleyball.Checked = false;
                //chkSoccer.Checked = false;
                //txtAge.Clear();
                //txtSaying.Clear();
                //txtBrowse.Clear();
                return;

            }

            try
            {

                string name = txtName.Text;
                string gender = "";
                if (rdoMale.Checked)
                {
                    gender = "Male";
                }
                if (rdoFemale.Checked)
                {
                    gender = "Female";
                }

                string hobbies = "";
                if (chkChess.Checked) hobbies += "CHESS";
                if (chkMobile.Checked) hobbies += "GAMES";
                if (chkCycling.Checked) hobbies += "CYCLING ";

                string address = txtAddress.Text;
                string email = txtEmail.Text;
                string birthday = dateTimePicker1.Text;
                string age = txtAge.Text;
                string favColor = cmbColor.Text;
                string user = txtUsername.Text;
                string pass = txtPassword.Text;
                string saying = txtSaying.Text;
                string course = cmbCourse.Text;
                string profile = txtBrowse.Text;
                string status = txtStatus.Text;

                book.LoadFromFile(@"C:\Users\Computer\Desktop\EVENTDRIVEN\sint\EVENTDRIVE_ALEGADO\BOOKDB.xlsx");
                Worksheet sheet = book.Worksheets[0];

                for (int row = 2; row <= sheet.LastRow; row++)//ERROR FOR EXISTING USER AND PASS
                {
                    string existingUsername = sheet.Range[row, 6].Value;
                    string existingPassword = sheet.Range[row, 7].Value;

                    if (existingUsername == txtUsername.Text)
                    {
                        MessageBox.Show("Username already exists. Please choose a different one.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (existingPassword == txtPassword.Text)
                    {
                        MessageBox.Show("Password already exists. Please choose a different one.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                if (!IsValidEmail(email))
                {
                    MessageBox.Show("Invalid email format. Please enter a valid email.");
                    return;
                }


                int i = sheet.Rows.Length + 1;
                sheet.Range[i, 1].Value = name;
                sheet.Range[i, 2].Value = gender;
                sheet.Range[i, 3].Value = hobbies;
                sheet.Range[i, 4].Value = favColor;
                sheet.Range[i, 5].Value = saying;
                sheet.Range[i, 6].Value = user;
                sheet.Range[i, 7].Value = pass;
                sheet.Range[i, 8].Value = email;
                sheet.Range[i, 9].Value = birthday;
                sheet.Range[i, 10].Value = age;
                sheet.Range[i, 11].Value = address;
                sheet.Range[i, 12].Value = status;
                sheet.Range[i, 13].Value = course;
                sheet.Range[i, 14].Value = profile;

                book.SaveToFile(@"C:\Users\Computer\Desktop\EVENTDRIVEN\sint\EVENTDRIVE_ALEGADO\BOOKDB.xlsx", ExcelVersion.Version2016);

                DialogResult result = MessageBox.Show("Student successfully added!", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (result == DialogResult.OK)
                {
                    Dashboard frm4 = new Dashboard();
                    MyLogs logs = new MyLogs();
                    logs.insertLogs(DisplayIt.CurrentUser, "Added a new Student to the list.");
                    frm4.Show();

                }
                txtUsername.Clear();
                txtPassword.Clear();
                txtName.Clear();
                rdoMale.Checked = false;
                rdoFemale.Checked = false;
                txtAddress.Clear();
                txtEmail.Clear();
                dateTimePicker1.Checked = false;
                chkChess.Checked = false;
                chkMobile.Checked = false;
                chkCycling.Checked = false;
                cmbColor.SelectedIndex = -1;
                cmbCourse.SelectedIndex = -1;
                txtSaying.Clear();
                txtStatus.Clear();
                txtBrowse.Clear();



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //string data = "";
            //string gender = "";
            //string hobbies = "";

            //data = txtName.Text;
            //if (radMale.Checked == true)
            //{
            //    gender = radMale.Text;
            //}
            //if (radFemale.Checked == true)
            //{
            //    gender = radFemale.Text;
            //}
            //if (chkBasketball.Checked)
            //{
            //    hobbies += " " + chkBasketball.Text;
            //}
            //if (chkVolleyball.Checked)
            //{
            //    hobbies += " " + chkVolleyball.Text;
            //}
            //if (chkSoccer.Checked)
            //{
            //    hobbies += " " + chkSoccer.Text;
            //}
            //data += cmbFavColor.Text;
            //data += txtSaying.Text;

            //info[i] = data;
            //i++;

            ////f2.insert(txtName.Text, gender, hobbies, cmbFavColor.Text, txtSaying.Text);
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"C:\Users\Computer\Desktop\EVENTDRIVEN\sint\EVENTDRIVE_ALEGADO\BOOKDB.xlsx");
            //Worksheet sh = book.Worksheets[0];
            //int row = sh.Rows.Length + 1;
            //sh.Range[row, 1].Value = txtName.Text;
            //sh.Range[row, 2].Value = gender;
            //sh.Range[row, 3].Value = txtAddress.Text;
            //sh.Range[row, 4].Value = txtEmail.Text;
            //sh.Range[row, 5].Value = dtpBirthday.Text;
            //sh.Range[row, 6].Value = txtAge.Text;
            //sh.Range[row, 7].Value = txtUsername.Text;
            //sh.Range[row, 8].Value = txtPassword.Text;
            //sh.Range[row, 9].Value = hobbies;
            //sh.Range[row, 10].Value = cmbFavColor.Text;
            //sh.Range[row, 11].Value = txtSaying.Text;
            //sh.Range[row, 12].Value = cmbCourse.Text;
            //sh.Range[row, 13].Value = txtStatus.Text;



            //if (lblMessage.Text == "")
            //{
            //    book.SaveToFile(@"C:\Users\Computer\Desktop\EVENTDRIVEN\sint\EVENTDRIVE_ALEGADO\BOOKDB.xlsx", ExcelVersion.Version2016);
            //    DataTable dt = sh.ExportDataTable();
            //    f2.dtgInfo.DataSource = dt;
            //}
            //else return;

            txtName.Clear();
            txtSaying.Clear();
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            //string val = "";
            //for(int x=0;x<Person.Length;x++) 
            //{
            //    val += "[" + x + "]=" + Person[x] + "\n";
            //}
            //MessageBox.Show(val);
            
            f2.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnAdd.Visible = false;

            lblerror.Text = "";
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(txtUsername.Text)) errors.AppendLine("• Username is required.");
            if (string.IsNullOrWhiteSpace(txtPassword.Text)) errors.AppendLine("• Password is required.");
            if (string.IsNullOrWhiteSpace(txtName.Text)) errors.AppendLine("• Name is required.");
            if (!rdoMale.Checked && !rdoFemale.Checked) errors.AppendLine("• Gender is required.");
            if (string.IsNullOrWhiteSpace(txtAddress.Text)) errors.AppendLine("• Address is required.");
            if (string.IsNullOrWhiteSpace(txtEmail.Text)) errors.AppendLine("• Email is required.");
            if (!dateTimePicker1.Checked) errors.AppendLine("• Birthday is required.");
            if (!chkChess.Checked && !chkMobile.Checked && !chkCycling.Checked) errors.AppendLine("• At least one sport must be selected.");
            if (cmbColor.SelectedIndex == -1) errors.AppendLine("• Favorite color must be selected.");
            if (cmbCourse.SelectedIndex == -1) errors.AppendLine("• Course must be selected.");
            if (string.IsNullOrWhiteSpace(txtSaying.Text)) errors.AppendLine("• Saying is required.");
            if (string.IsNullOrWhiteSpace(txtStatus.Text)) errors.AppendLine("• Status is required.");
            if (string.IsNullOrWhiteSpace(txtBrowse.Text)) errors.AppendLine("• Profile is required.");

            if (errors.Length > 0)
            {
                lblerror.Text = errors.ToString();
                lblerror.Visible = true;
                MessageBox.Show("Please fill in all required fields!", "MISSING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //txtName.Clear();
                //txtAddress.Clear();
                //txtEmail.Clear();
                //dtpBirthday.Checked = false;
                //txtUsername.Clear();
                //txtPassword.Clear();
                //cmbFavColor.SelectedIndex = -1;
                //cmbCourse.SelectedIndex = -1;
                //radMale.Checked = false;
                //radFemale.Checked = false;
                //chkBasketball.Checked = false;
                //chkVolleyball.Checked = false;
                //chkSoccer.Checked = false;
                //txtAge.Clear();
                //txtSaying.Clear();
                //txtBrowse.Clear();
                return;

            }
            try
            {
                Dashboard fmr4 = new Dashboard();
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\Computer\Desktop\EVENTDRIVEN\sint\EVENTDRIVE_ALEGADO\BOOKDB.xlsx");
                Worksheet sheet = book.Worksheets[0];
                string name = txtName.Text;
                string gender = "";
                if (rdoMale.Checked)
                {
                    gender = "Male";
                }
                if (rdoFemale.Checked)
                {
                    gender = "Female";
                }

                string hobbies = "";
                if (chkChess.Checked) hobbies += "CHESS";
                if (chkMobile.Checked) hobbies += "GAMES";
                if (chkCycling.Checked) hobbies += "CYCLING ";

                string address = txtAddress.Text;
                string email = txtEmail.Text;
                string birthday = dateTimePicker1.Text;
                string age = txtAge.Text;
                string favColor = cmbColor.Text;
                string user = txtUsername.Text;
                string pass = txtPassword.Text;
                string saying = txtSaying.Text;
                string course = cmbCourse.Text;
                string profile = txtBrowse.Text;
                string status = txtStatus.Text;

                if (!IsValidEmail(email))
                {
                    MessageBox.Show("Invalid email format. Please enter a valid email.");
                    return;
                }

                int ID = Convert.ToInt32(lblInfo.Text);
                Form2 frm2 = new Form2();
                frm2.UpdateToExcel(ID, name, gender, hobbies, address, email, birthday, age, favColor, user, pass, saying, course, status, profile);

                DialogResult result = MessageBox.Show("Student details updated successfully!", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result == DialogResult.OK)
                {
                    frm2.Show();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            // Get the selected date from the DateTimePicker
            DateTime selectedDate = dateTimePicker1.Value;

            // Calculate the age
            int age = CalculateAge(selectedDate);

            // Update the label with the calculated age
            txtAge.Text = age.ToString();
        }
        private int CalculateAge(DateTime birthDate)
        {
            DateTime today = DateTime.Today;
            int age = today.Year - birthDate.Year;

            // Adjust age if the birthday hasn't occurred yet this year
            if (birthDate > today.AddYears(-age)) age--;

            return age;
        }

        private void lblerror_Click(object sender, EventArgs e)
        {
            lblerror.Visible = false;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            if (d.ShowDialog() == DialogResult.OK)
            {
                txtBrowse.Text = d.FileName;
            }
        }

        private void txtBrowse_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblInfo_Click(object sender, EventArgs e)
        {
            

        }
    }
}
