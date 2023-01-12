
namespace ExcelToXML
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TxtInputExcelPath = new System.Windows.Forms.TextBox();
            this.TxtOutputPath = new System.Windows.Forms.TextBox();
            this.BtnBrowseInputFile = new System.Windows.Forms.Button();
            this.BtnBrowsePath = new System.Windows.Forms.Button();
            this.BtnStart = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.BtnSelectPymatFile = new System.Windows.Forms.Button();
            this.TxtPymatFile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.BtnExcelToPymat = new System.Windows.Forms.Button();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 80);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 20);
            this.label1.TabIndex = 33;
            this.label1.Text = "Excel Input:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 25);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 20);
            this.label2.TabIndex = 34;
            this.label2.Text = "Output Path:";
            // 
            // TxtInputExcelPath
            // 
            this.TxtInputExcelPath.Location = new System.Drawing.Point(122, 75);
            this.TxtInputExcelPath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.TxtInputExcelPath.Name = "TxtInputExcelPath";
            this.TxtInputExcelPath.Size = new System.Drawing.Size(368, 26);
            this.TxtInputExcelPath.TabIndex = 31;
            // 
            // TxtOutputPath
            // 
            this.TxtOutputPath.Location = new System.Drawing.Point(122, 20);
            this.TxtOutputPath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.TxtOutputPath.Name = "TxtOutputPath";
            this.TxtOutputPath.Size = new System.Drawing.Size(368, 26);
            this.TxtOutputPath.TabIndex = 32;
            // 
            // BtnBrowseInputFile
            // 
            this.BtnBrowseInputFile.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.BtnBrowseInputFile.Location = new System.Drawing.Point(501, 75);
            this.BtnBrowseInputFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnBrowseInputFile.Name = "BtnBrowseInputFile";
            this.BtnBrowseInputFile.Size = new System.Drawing.Size(90, 25);
            this.BtnBrowseInputFile.TabIndex = 29;
            this.BtnBrowseInputFile.Text = "Browse";
            this.BtnBrowseInputFile.UseVisualStyleBackColor = true;
            this.BtnBrowseInputFile.Click += new System.EventHandler(this.BtnBrowseInputFile_Click);
            // 
            // BtnBrowsePath
            // 
            this.BtnBrowsePath.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.BtnBrowsePath.Location = new System.Drawing.Point(501, 20);
            this.BtnBrowsePath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnBrowsePath.Name = "BtnBrowsePath";
            this.BtnBrowsePath.Size = new System.Drawing.Size(90, 25);
            this.BtnBrowsePath.TabIndex = 30;
            this.BtnBrowsePath.Text = "Browse";
            this.BtnBrowsePath.UseVisualStyleBackColor = true;
            this.BtnBrowsePath.Click += new System.EventHandler(this.BtnBrowsePath_Click);
            // 
            // BtnStart
            // 
            this.BtnStart.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.BtnStart.Location = new System.Drawing.Point(332, 176);
            this.BtnStart.Name = "BtnStart";
            this.BtnStart.Size = new System.Drawing.Size(158, 45);
            this.BtnStart.TabIndex = 35;
            this.BtnStart.Text = "Pymat To Excel";
            this.BtnStart.UseVisualStyleBackColor = true;
            this.BtnStart.Click += new System.EventHandler(this.BtnStart_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // BtnSelectPymatFile
            // 
            this.BtnSelectPymatFile.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.BtnSelectPymatFile.Location = new System.Drawing.Point(501, 124);
            this.BtnSelectPymatFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnSelectPymatFile.Name = "BtnSelectPymatFile";
            this.BtnSelectPymatFile.Size = new System.Drawing.Size(90, 25);
            this.BtnSelectPymatFile.TabIndex = 29;
            this.BtnSelectPymatFile.Text = "Browse";
            this.BtnSelectPymatFile.UseVisualStyleBackColor = true;
            this.BtnSelectPymatFile.Click += new System.EventHandler(this.BtnSelectPymatFile_Click);
            // 
            // TxtPymatFile
            // 
            this.TxtPymatFile.Location = new System.Drawing.Point(122, 124);
            this.TxtPymatFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.TxtPymatFile.Name = "TxtPymatFile";
            this.TxtPymatFile.Size = new System.Drawing.Size(368, 26);
            this.TxtPymatFile.TabIndex = 31;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 129);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 20);
            this.label3.TabIndex = 33;
            this.label3.Text = "Pymat Input:";
            // 
            // BtnExcelToPymat
            // 
            this.BtnExcelToPymat.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.BtnExcelToPymat.Location = new System.Drawing.Point(122, 176);
            this.BtnExcelToPymat.Name = "BtnExcelToPymat";
            this.BtnExcelToPymat.Size = new System.Drawing.Size(158, 45);
            this.BtnExcelToPymat.TabIndex = 35;
            this.BtnExcelToPymat.Text = "Excel To Pymat";
            this.BtnExcelToPymat.UseVisualStyleBackColor = true;
            this.BtnExcelToPymat.Click += new System.EventHandler(this.BtnExcelToPymat_Click);
            // 
            // backgroundWorker2
            // 
            this.backgroundWorker2.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker2_DoWork);
            this.backgroundWorker2.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker2_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(635, 233);
            this.Controls.Add(this.BtnExcelToPymat);
            this.Controls.Add(this.BtnStart);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TxtPymatFile);
            this.Controls.Add(this.TxtInputExcelPath);
            this.Controls.Add(this.TxtOutputPath);
            this.Controls.Add(this.BtnSelectPymatFile);
            this.Controls.Add(this.BtnBrowseInputFile);
            this.Controls.Add(this.BtnBrowsePath);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Pymat - Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtInputExcelPath;
        private System.Windows.Forms.TextBox TxtOutputPath;
        private System.Windows.Forms.Button BtnBrowseInputFile;
        private System.Windows.Forms.Button BtnBrowsePath;
        private System.Windows.Forms.Button BtnStart;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button BtnSelectPymatFile;
        private System.Windows.Forms.TextBox TxtPymatFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BtnExcelToPymat;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
    }
}

