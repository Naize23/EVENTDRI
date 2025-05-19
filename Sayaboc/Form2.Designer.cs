namespace Sayaboc
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnDeleteLogs = new System.Windows.Forms.Button();
            this.lblExit = new System.Windows.Forms.Label();
            this.lblActive = new System.Windows.Forms.Label();
            this.lblInactive = new System.Windows.Forms.Label();
            this.lblLogs = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(54, 129);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(780, 303);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseDoubleClick);
            // 
            // txtSearch
            // 
            this.txtSearch.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSearch.Location = new System.Drawing.Point(369, 77);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(309, 21);
            this.txtSearch.TabIndex = 1;
            // 
            // btnSearch
            // 
            this.btnSearch.BackColor = System.Drawing.Color.Yellow;
            this.btnSearch.FlatAppearance.BorderSize = 0;
            this.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSearch.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(697, 74);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 2;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnNew
            // 
            this.btnNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnNew.FlatAppearance.BorderSize = 0;
            this.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNew.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNew.ForeColor = System.Drawing.Color.White;
            this.btnNew.Location = new System.Drawing.Point(554, 449);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(138, 49);
            this.btnNew.TabIndex = 4;
            this.btnNew.Text = "New";
            this.btnNew.UseVisualStyleBackColor = false;
            this.btnNew.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnDeleteLogs
            // 
            this.btnDeleteLogs.BackColor = System.Drawing.Color.Red;
            this.btnDeleteLogs.FlatAppearance.BorderSize = 0;
            this.btnDeleteLogs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDeleteLogs.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteLogs.ForeColor = System.Drawing.Color.White;
            this.btnDeleteLogs.Location = new System.Drawing.Point(200, 449);
            this.btnDeleteLogs.Name = "btnDeleteLogs";
            this.btnDeleteLogs.Size = new System.Drawing.Size(138, 49);
            this.btnDeleteLogs.TabIndex = 5;
            this.btnDeleteLogs.Text = "Delete ";
            this.btnDeleteLogs.UseVisualStyleBackColor = false;
            this.btnDeleteLogs.Click += new System.EventHandler(this.btnDeleteLogs_Click);
            // 
            // lblExit
            // 
            this.lblExit.AutoSize = true;
            this.lblExit.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExit.ForeColor = System.Drawing.Color.Red;
            this.lblExit.Location = new System.Drawing.Point(858, 9);
            this.lblExit.Name = "lblExit";
            this.lblExit.Size = new System.Drawing.Size(20, 19);
            this.lblExit.TabIndex = 6;
            this.lblExit.Text = "X";
            this.lblExit.Click += new System.EventHandler(this.lblExit_Click_1);
            // 
            // lblActive
            // 
            this.lblActive.AutoSize = true;
            this.lblActive.BackColor = System.Drawing.Color.Transparent;
            this.lblActive.Font = new System.Drawing.Font("Century", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblActive.ForeColor = System.Drawing.Color.White;
            this.lblActive.Location = new System.Drawing.Point(31, 31);
            this.lblActive.Name = "lblActive";
            this.lblActive.Size = new System.Drawing.Size(284, 33);
            this.lblActive.TabIndex = 7;
            this.lblActive.Text = "ACTIVE STUDENT";
            // 
            // lblInactive
            // 
            this.lblInactive.AutoSize = true;
            this.lblInactive.BackColor = System.Drawing.Color.Transparent;
            this.lblInactive.Font = new System.Drawing.Font("Century", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInactive.ForeColor = System.Drawing.Color.White;
            this.lblInactive.Location = new System.Drawing.Point(31, 31);
            this.lblInactive.Name = "lblInactive";
            this.lblInactive.Size = new System.Drawing.Size(319, 33);
            this.lblInactive.TabIndex = 8;
            this.lblInactive.Text = "INACTIVE STUDENT";
            // 
            // lblLogs
            // 
            this.lblLogs.AutoSize = true;
            this.lblLogs.BackColor = System.Drawing.Color.Transparent;
            this.lblLogs.Font = new System.Drawing.Font("Century", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLogs.ForeColor = System.Drawing.Color.White;
            this.lblLogs.Location = new System.Drawing.Point(117, 31);
            this.lblLogs.Name = "lblLogs";
            this.lblLogs.Size = new System.Drawing.Size(96, 33);
            this.lblLogs.TabIndex = 8;
            this.lblLogs.Text = "LOGS";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(890, 561);
            this.Controls.Add(this.lblLogs);
            this.Controls.Add(this.lblInactive);
            this.Controls.Add(this.lblActive);
            this.Controls.Add(this.lblExit);
            this.Controls.Add(this.btnDeleteLogs);
            this.Controls.Add(this.btnNew);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnSearch;
        public System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnDeleteLogs;
        private System.Windows.Forms.Label lblExit;
        public System.Windows.Forms.Label lblActive;
        public System.Windows.Forms.Label lblInactive;
        public System.Windows.Forms.Label lblLogs;
    }
}