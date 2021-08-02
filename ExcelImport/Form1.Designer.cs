
namespace ExcelImport
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnGetJson = new System.Windows.Forms.Button();
            this.btnNewJson = new System.Windows.Forms.Button();
            this.btnURLReplace = new System.Windows.Forms.Button();
            this.btnJSON = new System.Windows.Forms.Button();
            this.btnDownloadFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(10, 76);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(106, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Generate Report";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(119, 35);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(569, 23);
            this.txtPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(10, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Excel File Path :";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(325, 106);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(405, 23);
            this.progressBar1.TabIndex = 3;
            this.progressBar1.Visible = false;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lblStatus.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.lblStatus.Location = new System.Drawing.Point(10, 108);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(40, 15);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "label2";
            this.lblStatus.Visible = false;
            // 
            // btnGetJson
            // 
            this.btnGetJson.Location = new System.Drawing.Point(121, 76);
            this.btnGetJson.Name = "btnGetJson";
            this.btnGetJson.Size = new System.Drawing.Size(96, 23);
            this.btnGetJson.TabIndex = 5;
            this.btnGetJson.Text = "Generate JSON";
            this.btnGetJson.UseVisualStyleBackColor = true;
            this.btnGetJson.Click += new System.EventHandler(this.btnGetJson_Click);
            // 
            // btnNewJson
            // 
            this.btnNewJson.Location = new System.Drawing.Point(223, 76);
            this.btnNewJson.Name = "btnNewJson";
            this.btnNewJson.Size = new System.Drawing.Size(96, 23);
            this.btnNewJson.TabIndex = 6;
            this.btnNewJson.Text = "New JSON";
            this.btnNewJson.UseVisualStyleBackColor = true;
            this.btnNewJson.Click += new System.EventHandler(this.btnNewJson_Click);
            // 
            // btnURLReplace
            // 
            this.btnURLReplace.Location = new System.Drawing.Point(325, 76);
            this.btnURLReplace.Name = "btnURLReplace";
            this.btnURLReplace.Size = new System.Drawing.Size(96, 23);
            this.btnURLReplace.TabIndex = 7;
            this.btnURLReplace.Text = "Find - Replace";
            this.btnURLReplace.UseVisualStyleBackColor = true;
            this.btnURLReplace.Click += new System.EventHandler(this.btnURLReplace_Click);
            // 
            // btnJSON
            // 
            this.btnJSON.Location = new System.Drawing.Point(427, 76);
            this.btnJSON.Name = "btnJSON";
            this.btnJSON.Size = new System.Drawing.Size(156, 23);
            this.btnJSON.TabIndex = 8;
            this.btnJSON.Text = "Generate JSON from Excel";
            this.btnJSON.UseVisualStyleBackColor = true;
            this.btnJSON.Click += new System.EventHandler(this.btnJSON_Click);
            // 
            // btnDownloadFile
            // 
            this.btnDownloadFile.Location = new System.Drawing.Point(589, 77);
            this.btnDownloadFile.Name = "btnDownloadFile";
            this.btnDownloadFile.Size = new System.Drawing.Size(156, 23);
            this.btnDownloadFile.TabIndex = 9;
            this.btnDownloadFile.Text = "Download FWS Pdf File";
            this.btnDownloadFile.UseVisualStyleBackColor = true;
            this.btnDownloadFile.Click += new System.EventHandler(this.btnDownloadFile_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(965, 132);
            this.Controls.Add(this.btnDownloadFile);
            this.Controls.Add(this.btnJSON);
            this.Controls.Add(this.btnURLReplace);
            this.Controls.Add(this.btnNewJson);
            this.Controls.Add(this.btnGetJson);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Validate Excel for HTML and URL";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnGetJson;
        private System.Windows.Forms.Button btnNewJson;
        private System.Windows.Forms.Button btnURLReplace;
        private System.Windows.Forms.Button btnJSON;
        private System.Windows.Forms.Button btnDownloadFile;
    }
}

